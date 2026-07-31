"""Compatibility entry point for the safe P2 geo-recovery workflow.

Run without arguments for a dry run. Pass ``--live`` only after reviewing the
summary. No decline email is sent by this workflow.
"""

import sys

from hiring_agent.sharepoint_scoring import recover_location_rejections


def run_rejected_audit(dry_run: bool = True) -> dict:
    return recover_location_rejections(dry_run=dry_run)


if __name__ == "__main__":
    run_rejected_audit(dry_run="--live" not in sys.argv)
