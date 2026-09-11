"""Local Ollama only: no Graph client, tenant access, scoring, or mail.

Run from P2: .venv/Scripts/python docs/benchmarks/stage2a/benchmark.py STAGE
Each stage measures three smart extractions, then each auxiliary call once.
The auxiliary inputs stay fixed at the last baseline extraction for comparison.
"""
import hashlib
import json
from pathlib import Path
import statistics
import sys
import time
from unittest.mock import patch

import requests

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))
from hiring_agent import extraction as ex

HERE = Path(__file__).resolve().parent
stage = sys.argv[1]
out = HERE / stage
out.mkdir(exist_ok=False)
resume = (HERE / "resume.txt").read_text(encoding="utf-8")
original_post = requests.post
records = []
active = ""


def measured_post(url, **kwargs):
    start = time.perf_counter()
    record = {"call": active, "request": kwargs}
    try:
        response = original_post(url, **kwargs)
        record["raw"] = response.json()
        record["status_code"] = response.status_code
        return response
    except Exception as error:
        record["error"] = repr(error)
        raise
    finally:
        record["seconds"] = time.perf_counter() - start
        records.append(record)
        (out / "raw.json").write_text(json.dumps(records, indent=2), encoding="utf-8")
        raw = record.get("raw", {})
        print(f"{active}: {record['seconds']:.3f}s; tokens={raw.get('eval_count')}; "
              f"done={raw.get('done_reason')}; thinking_chars="
              f"{len(raw.get('message', {}).get('thinking', ''))}", flush=True)


metadata = {
    "resume_chars": len(resume),
    "resume_sha256": hashlib.sha256(resume.encode()).hexdigest(),
    "model": ex.OLLAMA_MODEL,
    "timeout": ex.OLLAMA_TIMEOUT,
    "version": requests.get(f"{ex.OLLAMA_HOST}/api/version", timeout=10).json(),
    "model_info": original_post(f"{ex.OLLAMA_HOST}/api/show",
                                json={"model": ex.OLLAMA_MODEL}, timeout=10).json(),
}
(out / "metadata.json").write_text(json.dumps(metadata, indent=2), encoding="utf-8")
for name in ("hiring_agent/extraction.py", "config.yaml"):
    (out / Path(name).name).write_bytes((ROOT / name).read_bytes())
print(f"Stage {stage}: {len(resume)} chars; timeout {ex.OLLAMA_TIMEOUT}s", flush=True)
outputs = {}
with patch.object(requests, "post", measured_post):
    for run in range(1, 4):
        active = f"extract_{run}"
        start = time.perf_counter()
        fields = ex.extract_candidate_details_smart(resume)
        outputs[active] = {"seconds": time.perf_counter() - start, "fields": fields}
        (out / "outputs.json").write_text(json.dumps(outputs, indent=2), encoding="utf-8")
    if stage == "01_baseline":
        (HERE / "auxiliary_fields.json").write_text(json.dumps(fields, indent=2), encoding="utf-8")
    fields = json.loads((HERE / "auxiliary_fields.json").read_text(encoding="utf-8"))
    active = "recheck"
    outputs[active] = ex.ai_recheck_fields(fields, resume, "")
    active = "role_note"
    outputs[active] = ex.infer_looking_for_role(resume, "", fields["skills"])
    active = "portfolios"
    outputs[active] = ex.infer_missing_portfolios(resume, "N/A", "N/A", "N/A")
(out / "outputs.json").write_text(json.dumps(outputs, indent=2), encoding="utf-8")
median = statistics.median(outputs[f"extract_{i}"]["seconds"] for i in range(1, 4))
print(f"MEDIAN smart extraction: {median:.3f}s", flush=True)
