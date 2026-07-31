"""Remove a proven candidate resume from the local active-JD data sources."""
from __future__ import annotations

import argparse
import datetime as dt
import json
import shutil
from pathlib import Path


BASE = Path(__file__).parent
FILES = (BASE / "jd_sources.json", BASE / "jd_database.json", BASE / "jd_roles_cache.json")
INVALID_TITLE = "Parsika Shah Data  BusinessAnalyst"


def run(apply: bool) -> int:
    source = json.loads(FILES[0].read_text(encoding="utf-8"))
    database = json.loads(FILES[1].read_text(encoding="utf-8"))
    cache = json.loads(FILES[2].read_text(encoding="utf-8"))
    source_before = len(source.get("descriptions", []))
    database_before = len(database.get("records", []))
    cache_before = len(cache.get("roles", []))
    source["descriptions"] = [
        item for item in source.get("descriptions", [])
        if str(item.get("title", "")) != INVALID_TITLE
    ]
    database["records"] = [
        item for item in database.get("records", [])
        if str(item.get("title", "")) != INVALID_TITLE
    ]
    database["total_roles"] = len(database["records"])
    cache["roles"] = [
        item for item in cache.get("roles", [])
        if str(item.get("title", "")) != INVALID_TITLE
    ]
    counts = {
        "jd_sources": (source_before, len(source["descriptions"])),
        "jd_database": (database_before, len(database["records"])),
        "jd_cache": (cache_before, len(cache["roles"])),
    }
    print(json.dumps({"invalid_title": INVALID_TITLE, "counts": counts,
                      "mode": "APPLY" if apply else "DRY-RUN"}, indent=2))
    if not apply:
        return 0
    if any(before - after != 1 for before, after in counts.values()):
        raise RuntimeError("expected exactly one invalid JD in each local source")
    stamp = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    backup = BASE / "P2_Logs" / "backups" / f"jd_cleanup_{stamp}"
    backup.mkdir(parents=True, exist_ok=False)
    for file in FILES:
        shutil.copy2(file, backup / file.name)
    FILES[0].write_text(json.dumps(source, indent=2, ensure_ascii=False), encoding="utf-8")
    FILES[1].write_text(json.dumps(database, indent=2, ensure_ascii=False), encoding="utf-8")
    # The cache will be rebuilt below; write the filtered form first so a failed rebuild
    # can never continue exposing the invalid candidate-as-JD role.
    FILES[2].write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    from hiring_agent.jd_sources import build_jd_cache
    roles = build_jd_cache()
    if len(roles) != 85 or any(r.get("title") == INVALID_TITLE for r in roles):
        raise RuntimeError("JD cache rebuild did not produce the expected clean 85-role set")
    print(f"BACKUP={backup}")
    print("REBUILT_ROLES=85")
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()
    raise SystemExit(run(args.apply))
