"""Compatibility entry point for historical location-rejection recovery.

Named-candidate restoration used to fabricate US country/location values. The
shared tri-state recovery now handles these rows and defaults to read-only mode.
"""

import sys

from hiring_agent.sharepoint_scoring import recover_location_rejections


def restore_us_candidates(dry_run: bool = True) -> dict:
    return recover_location_rejections(dry_run=dry_run)


if __name__ == "__main__":
    restore_us_candidates(dry_run="--live" not in sys.argv)
