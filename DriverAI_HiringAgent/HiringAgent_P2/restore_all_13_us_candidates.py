"""Compatibility entry point for historical location-rejection recovery.

The original script promoted every uncertain candidate to Scored and wrote
"United States" into Country. That behavior is retired. This wrapper now uses
P2's tri-state recovery and defaults to a read-only dry run.
"""

import sys

from hiring_agent.sharepoint_scoring import recover_location_rejections


def restore_all_us_candidates(dry_run: bool = True) -> dict:
    return recover_location_rejections(dry_run=dry_run)


if __name__ == "__main__":
    restore_all_us_candidates(dry_run="--live" not in sys.argv)
