"""Print focused CV evidence for one CandidateList row without JD or mail work."""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env"))

from repair_candidate_list_batches import _deep_text, _download_current, _real_extension, _row_by_id
from sharepoint_client import SharePointClient

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main(app_id: str, context: int) -> int:
    client = SharePointClient()
    row = _row_by_id(client, app_id)
    if row is None:
        raise RuntimeError(f"{app_id} is not in CandidateList")
    name, raw = _download_current(client, row["values"])
    ext = _real_extension(raw, name)
    text, forced_ocr = _deep_text(raw, f"resume{ext}")
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines()]
    lines = [line for line in lines if line]
    print(f"FILE={name} BYTES={len(raw)} CHARS={len(text.strip())} OCR={forced_ocr}")
    print("\n--- HEADER ---")
    for number, line in enumerate(lines[:80], 1):
        print(f"{number:04}: {line}")
    patterns = re.compile(
        r"education|university|college|bachelor|master|b\.?s\.?|m\.?s\.?|mba|ph\.?d|degree",
        re.I,
    )
    hits = [i for i, line in enumerate(lines) if patterns.search(line)]
    print("\n--- EDUCATION / DEGREE CONTEXT ---")
    shown: set[int] = set()
    for hit in hits:
        for i in range(max(0, hit - context), min(len(lines), hit + context + 1)):
            if i not in shown:
                print(f"{i + 1:04}: {lines[i]}")
                shown.add(i)
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--app-id", required=True)
    parser.add_argument("--context", type=int, default=3)
    args = parser.parse_args()
    raise SystemExit(main(args.app_id, args.context))
