import sys, os, json, time, re
from collections import defaultdict, Counter
from pathlib import Path

env_path = Path('HiringAgent_App_P2/.env')
for line in env_path.read_text(encoding='utf-8').splitlines():
    if '=' in line and not line.startswith('#'):
        k, v = line.split('=', 1)
        os.environ[k.strip()] = v.strip()

sys.path.insert(0, 'HiringAgent_App_P2')

from sharepoint_client import SharePointClient
from hiring_agent.config import logger, COLUMNS, STATUS_LOCATION_REVIEW, STATUS_SCORED
from hiring_agent.geo import GeoDecision, classify_location_usa, reconcile_us_country
from hiring_agent.sharepoint_scoring import _ensure_schema, score_from_sharepoint

def delete_table_row(client: SharePointClient, index: int):
    """Delete a table row at itemAt(index)."""
    url = f"{client._wb_base()}/tables/{client.table}/rows/itemAt(index={index})"
    try:
        client._req("DELETE", url)
        return True
    except Exception as e:
        logger.warning(f"Could not delete row index {index}: {e}")
        return False

def run_row_by_row_healer():
    client = SharePointClient()
    print("=" * 90)
    print("  EXHAUSTIVE ROW-BY-ROW AND COLUMN-BY-COLUMN SHAREPOINT HEALER & AUDITOR")
    print("=" * 90)

    # 1. Ensure Table Schema
    _ensure_schema(client)

    # 2. Fetch fresh rows
    rows = client.list_rows()
    print(f"\n[INFO] Loaded {len(rows)} live rows from SharePoint master table.")

    stats = {
        'total_rows_inspected': len(rows),
        'geo_location_fixes': 0,
        'country_fixes': 0,
        'name_fixes': 0,
        'education_fixes': 0,
        'portfolio_fixes': 0,
        'app_updates_fixes': 0,
        'duplicate_rows_deleted': 0,
        'rows_patched': 0,
    }

    # -------------------------------------------------------------------------
    # STEP 1: DEDUPLICATION & MERGING (BY APPLICATION ID & EMAIL)
    # -------------------------------------------------------------------------
    app_id_map = defaultdict(list)
    for r in rows:
        app_id = str(r['values'].get('Application ID', '') or '').strip()
        if app_id:
            app_id_map[app_id].append(r)

    duplicate_groups = {k: v for k, v in app_id_map.items() if len(v) > 1}
    print(f"\n[STEP 1] Found {len(duplicate_groups)} Application IDs with duplicate rows.")

    rows_to_delete = []
    rows_to_patch = []

    for app_id, dup_group in duplicate_groups.items():
        canonical_row = dup_group[0]
        canonical_idx = canonical_row['index']
        canonical_vals = dict(canonical_row['values'])

        print(f"  - App ID '{app_id}' has {len(dup_group)} duplicate rows. Merging into Row index {canonical_idx} ({canonical_vals.get('Full Name')})...")

        for extra_row in dup_group[1:]:
            extra_idx = extra_row['index']
            rows_to_delete.append(extra_idx)
            extra_vals = extra_row['values']

            # Merge missing fields from duplicate row into canonical row
            for col, val in extra_vals.items():
                curr = canonical_vals.get(col, '')
                if (not str(curr or '').strip() or str(curr).strip() in ('N/A', 'Not extracted')) and str(val or '').strip() and str(val).strip() not in ('N/A', 'Not extracted'):
                    canonical_vals[col] = val

        rows_to_patch.append((canonical_idx, canonical_vals, canonical_row['values']))

    # Patch canonical rows
    for idx, vals, orig in rows_to_patch:
        client.update_row(idx, vals, current_values=orig)

    # Delete duplicate rows in reverse order of row index
    rows_to_delete.sort(reverse=True)
    for idx in rows_to_delete:
        print(f"  - Deleting duplicate Row index {idx}...")
        if delete_table_row(client, idx):
            stats['duplicate_rows_deleted'] += 1

    time.sleep(2)
    # Reload cleaned rows
    rows = client.list_rows()
    print(f"\n[INFO] Post-deduplication table size: {len(rows)} rows.")

    # -------------------------------------------------------------------------
    # STEP 2: ROW-BY-ROW COLUMN-BY-COLUMN INSPECTION & FIXES
    # -------------------------------------------------------------------------
    print(f"\n[STEP 2] Inspecting & healing every single row (1 to {len(rows)})...")

    for r in rows:
        idx = r['index']
        v = dict(r['values'])
        app_id = str(v.get('Application ID', '') or '').strip()
        name = str(v.get('Full Name', '') or '').strip()
        email = str(v.get('Email', '') or '').strip()
        phone = str(v.get('Phone', '') or '').strip()
        loc = str(v.get('Location', '') or '').strip()
        ctry = str(v.get('Country', '') or '').strip()
        edu = str(v.get('Education', '') or '').strip()
        mail_body = str(v.get('Mail Body', '') or '')
        status = str(v.get('Status', '') or '').strip()

        row_modified = False
        changes = []

        # 1. Full Name Fix
        if not name or name.lower() == 'not extracted':
            if email and '@' in email:
                localpart = email.split('@')[0]
                healed_name = localpart.replace('.', ' ').replace('_', ' ').replace('-', ' ').title()
                v['Full Name'] = healed_name
                changes.append(f"Full Name: '{name}' -> '{healed_name}'")
                stats['name_fixes'] += 1
                row_modified = True

        # 2. Geo fields: only a confirmed current US location may canonicalize Country.
        decision, reason = classify_location_usa(
            loc, ctry, phone=phone, education=edu)
        reconciled_country = reconcile_us_country(
            ctry, loc, decision == GeoDecision.CONFIRMED_US)
        if (decision == GeoDecision.CONFIRMED_US
                and reconciled_country != ctry):
            v['Country'] = reconciled_country
            changes.append(f"Country: '{ctry}' -> 'United States'")
            stats['country_fixes'] += 1
            row_modified = True

        # Supporting phone/education evidence prevents rejection but never fabricates
        # a current address from a university or employer location.
        if decision == GeoDecision.UNKNOWN and status != STATUS_LOCATION_REVIEW:
            v['Status'] = STATUS_LOCATION_REVIEW
            changes.append(f"Status: '{status}' -> '{STATUS_LOCATION_REVIEW}' ({reason})")
            stats['geo_location_fixes'] += 1
            row_modified = True

        # 4. Education Healing
        if not edu or edu.lower() in ('n/a', 'not extracted'):
            if 'Master' in mail_body or 'MS' in mail_body or 'B.S.' in mail_body or 'Bachelor' in mail_body:
                lines = [l.strip() for l in mail_body.split('|') if 'master' in l.lower() or 'bachelor' in l.lower() or 'university' in l.lower() or 'graduat' in l.lower()]
                if lines:
                    v['Education'] = lines[0][:120]
                    changes.append(f"Education: '{edu}' -> '{v['Education']}'")
                    stats['education_fixes'] += 1
                    row_modified = True

        # 5. Portfolio Slots Formatting ('N/A' default)
        for p_col in ['Portfolio 1', 'Portfolio 2', 'Portfolio 3']:
            p_val = str(v.get(p_col, '') or '').strip()
            if not p_val or p_val.lower() == 'not extracted':
                v[p_col] = 'N/A'
                row_modified = True
                stats['portfolio_fixes'] += 1

        # 6. Application Updates Default (0)
        app_upd = str(v.get('Application Updates', '') or '').strip()
        if not app_upd:
            v['Application Updates'] = 0
            row_modified = True
            stats['app_updates_fixes'] += 1

        # Apply patch if modifications were made
        if row_modified:
            client.update_row(idx, v, current_values=r['values'])
            stats['rows_patched'] += 1
            print(f"  [FIXED Row index #{idx:3}] AppID: {app_id} | Name: {v.get('Full Name')} | Changes: {', '.join(changes)}")

    # -------------------------------------------------------------------------
    # STEP 3: RE-SCORE PENDING & UNSCORED ROWS WITH P2 AI SCORER
    # -------------------------------------------------------------------------
    print("\n[STEP 3] Triggering Phase 2 AI scoring pass on pending rows...")
    try:
        score_from_sharepoint()
        print("  - AI Scoring pass complete!")
    except Exception as e:
        print(f"  - AI Scoring pass finished: {e}")

    print("\n" + "=" * 90)
    print("  EXHAUSTIVE ROW-BY-ROW HEALER COMPLETE SUMMARY")
    print("=" * 90)
    print(f"  - Total Rows Inspected : {stats['total_rows_inspected']}")
    print(f"  - Duplicate Rows Deleted: {stats['duplicate_rows_deleted']}")
    print(f"  - Rows Patched / Healed : {stats['rows_patched']}")
    print(f"  - Geo Location Fixes    : {stats['geo_location_fixes']}")
    print(f"  - Country Fixes         : {stats['country_fixes']}")
    print(f"  - Name Healing          : {stats['name_fixes']}")
    print(f"  - Education Healing     : {stats['education_fixes']}")
    print("=" * 90)

if __name__ == '__main__':
    run_row_by_row_healer()
