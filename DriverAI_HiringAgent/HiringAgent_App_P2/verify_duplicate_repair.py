"""Verify one removed duplicate and its retained CandidateList application."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env"))

from sharepoint_client import SharePointClient


def main(removed: str, keeper: str) -> int:
    rows = SharePointClient().list_rows()
    by_id = {
        str(row["values"].get("Application ID", "")): row["values"] for row in rows
    }
    result = {
        "removed_application_id": removed,
        "removed_absent": removed not in by_id,
        "keeper_application_id": keeper,
        "keeper_present": keeper in by_id,
        "keeper_phone": by_id.get(keeper, {}).get("Phone", ""),
        "keeper_resume_url_present": bool(by_id.get(keeper, {}).get("Resume URL")),
    }
    print(json.dumps(result, indent=2))
    return 0 if all((result["removed_absent"], result["keeper_present"],
                     result["keeper_resume_url_present"])) else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--removed", required=True)
    parser.add_argument("--keeper", required=True)
    args = parser.parse_args()
    raise SystemExit(main(args.removed, args.keeper))
