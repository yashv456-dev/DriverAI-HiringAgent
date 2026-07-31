import sys, os, json, time
from collections import defaultdict, Counter
from pathlib import Path

env_path = Path('HiringAgent_App_P2/.env')
for line in env_path.read_text(encoding='utf-8').splitlines():
    if '=' in line and not line.startswith('#'):
        k, v = line.split('=', 1)
        os.environ[k.strip()] = v.strip()

sys.path.insert(0, 'HiringAgent_App_P2')

from sharepoint_client import SharePointClient
import hiring_agent.config as _cfg
from hiring_agent.config import logger, COLUMNS, STATUS_LOCATION_REVIEW, STATUS_SCORED
from hiring_agent.geo import GeoDecision, classify_location_usa, reconcile_us_country
from hiring_agent.sharepoint_scoring import score_from_sharepoint, _ensure_schema

def reconcile_full_sheet():
    client = SharePointClient()
    logger.info("=" * 80)
    logger.info("  STARTING CANDIDATELIST RECONCILIATION, DEDUPLICATION & GEO-FIX")
    logger.info("=" * 80)

    # Step 0: Ensure schema is intact
    _ensure_schema(client)

    # Fetch fresh rows
    rows = client.list_rows()
    logger.info(f"Loaded {len(rows)} total rows from SharePoint table.")

    # -------------------------------------------------------------------------
    # PHASE 1: DEDUPLICATION & MERGING
    # -------------------------------------------------------------------------
    # Group rows by Application ID
    app_id_map = defaultdict(list)
    for r in rows:
        app_id = r['values'].get('Application ID', '').strip()
        if app_id:
            app_id_map[app_id].append(r)

    duplicate_app_ids = {k: v for k, v in app_id_map.items() if len(v) > 1}
    logger.info(f"Found {len(duplicate_app_ids)} Application IDs with duplicate rows.")

    rows_to_delete = []
    rows_to_update = []

    for app_id, dup_group in duplicate_app_ids.items():
        # Canonical row is the one with highest completion or lowest index
        canonical_row = dup_group[0]
        canonical_values = dict(canonical_row['values'])

        # Merge fields from remaining duplicate rows
        for extra_row in dup_group[1:]:
            rows_to_delete.append(extra_row['index'])
            extra_vals = extra_row['values']
            for col, val in extra_vals.items():
                curr = canonical_values.get(col, '')
                if (not str(curr or '').strip() or str(curr).strip() in ('N/A', 'Not extracted')) and str(val or '').strip() and str(val).strip() not in ('N/A', 'Not extracted'):
                    canonical_values[col] = val

        rows_to_update.append((canonical_row['index'], canonical_values))

    # Update canonical rows first
    for idx, vals in rows_to_update:
        logger.info(f"Updating canonical row index {idx} ({vals.get('Application ID')})...")
        client.patch_row(client._main_table, idx, vals)

    # Delete extra duplicate rows in reverse order of index to preserve indices
    rows_to_delete.sort(reverse=True)
    logger.info(f"Deleting {len(rows_to_delete)} duplicate rows from table...")
    for idx in rows_to_delete:
        logger.info(f"Deleting duplicate row index {idx}...")
        client.delete_row(client._main_table, idx)

    time.sleep(2)
    rows = client.list_rows()
    logger.info(f"Post-deduplication total rows: {len(rows)}")

    # -------------------------------------------------------------------------
    # PHASE 2: GEO-LOCATION & COUNTRY RECONCILIATION
    # -------------------------------------------------------------------------
    geo_updated = 0
    for r in rows:
        idx = r['index']
        v = dict(r['values'])
        loc = str(v.get('Location', '') or '').strip()
        ctry = str(v.get('Country', '') or '').strip()
        phone = str(v.get('Phone', '') or '').strip()

        education = str(v.get('Education', '') or '').strip()
        decision, reason = classify_location_usa(
            loc, ctry, phone=phone, education=education)

        needs_patch = False

        # Rule A: canonicalize Country only after a confirmed current US location.
        reconciled_country = reconcile_us_country(
            ctry, loc, decision == GeoDecision.CONFIRMED_US)
        if (decision == GeoDecision.CONFIRMED_US
                and reconciled_country != ctry):
            v['Country'] = reconciled_country
            needs_patch = True

        # Rule B: uncertainty stays visible for review; no school/employer city is
        # rewritten as the candidate's current residence.
        if (decision == GeoDecision.UNKNOWN
                and str(v.get('Status', '') or '').strip() != STATUS_LOCATION_REVIEW):
            v['Status'] = STATUS_LOCATION_REVIEW
            needs_patch = True

        if needs_patch:
            client.patch_row(client._main_table, idx, v)
            geo_updated += 1

    logger.info(f"Updated geo/country on {geo_updated} candidate rows.")

    # -------------------------------------------------------------------------
    # PHASE 3: NAME & FIELD HEALING
    # -------------------------------------------------------------------------
    name_healed = 0
    for r in rows:
        idx = r['index']
        v = dict(r['values'])
        name = str(v.get('Full Name', '') or '').strip()
        email = str(v.get('Email', '') or '').strip()

        if not name or name.lower() == 'not extracted':
            if email and '@' in email:
                localpart = email.split('@')[0]
                clean_name = localpart.replace('.', ' ').replace('_', ' ').replace('-', ' ').title()
                v['Full Name'] = clean_name
                client.patch_row(client._main_table, idx, v)
                name_healed += 1

    logger.info(f"Healed missing names on {name_healed} rows.")

    # -------------------------------------------------------------------------
    # PHASE 4: SCORE PENDING ROWS
    # -------------------------------------------------------------------------
    logger.info("Scoring pending 'New Email Received' rows using score_from_sharepoint()...")
    try:
        score_from_sharepoint()
    except Exception as e:
        logger.warning(f"SharePoint scoring pass finished: {e}")

    logger.info("Candidate list reconciliation & geo-fix complete!")

if __name__ == '__main__':
    reconcile_full_sheet()
