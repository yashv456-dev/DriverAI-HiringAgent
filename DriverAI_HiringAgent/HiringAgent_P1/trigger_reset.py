"""
trigger_reset.py  —  Mark archived emails as unread and move them back to Inbox
so the Power Automate trigger re-fires on them.

HOW IT WORKS
------------
The PA "When a new email arrives in a shared mailbox (V2)" trigger uses a delta
query against the Inbox folder. Moving a message FROM Archive BACK TO Inbox is
detected as a "new change" by the delta — the next 1-minute poll picks it up and
runs the full flow (spam gates, duplicate check, reply, SharePoint row) as if the
email just arrived.

WHEN TO USE
-----------
  - Testing the flow against real historical mail without forwarding new test emails.
  - Reprocessing a batch that the flow may have missed (e.g. flow was off, misconfigured).
  - Validating the 30-column schema against live data.

PREREQUISITES
-------------
The Entra app registration (same app used by the P2 scoring worker) must have:
  Mail.ReadWrite  (Application permission, admin-consented)
Without it, Graph returns 403 on the mailbox endpoints.

USAGE
-----
  # Preview — no changes made:
  python trigger_reset.py --dry-run

  # Move last 31 days of Archive mail back to Inbox as unread (default):
  python trigger_reset.py

  # Custom window (any number of days):
  python trigger_reset.py --days 7
  python trigger_reset.py --days 90

  # ALL emails in Archive regardless of age (e.g. June mails, old test runs):
  python trigger_reset.py --all
  python trigger_reset.py --all --dry-run

  # Different mailbox:
  python trigger_reset.py --mailbox hiring@example.com

Credentials are loaded from  ../HiringAgent_App_P2/.env  (same as the P2 worker).
"""
import argparse
import datetime
import os
import sys
from pathlib import Path

# ── load .env from the P2 sibling directory (same Entra app credentials) ─────
try:
    from dotenv import load_dotenv
    _env = Path(__file__).resolve().parent.parent / "HiringAgent_App_P2" / ".env"
    if _env.exists():
        load_dotenv(_env)
        print(f"[env] loaded credentials from {_env}")
    else:
        load_dotenv()  # fallback: local .env or system env
except ImportError:
    pass  # no python-dotenv; rely on os.environ being pre-set

import requests

GRAPH = "https://graph.microsoft.com/v1.0"
_TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
_TIMEOUT = 30
_PAGE_SIZE = 50


# ── Graph helpers ─────────────────────────────────────────────────────────────

def _get_token(tenant: str, client_id: str, client_secret: str) -> str:
    resp = requests.post(
        _TOKEN_URL.format(tenant=tenant),
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=_TIMEOUT,
    )
    if not resp.ok:
        print(f"[ERROR] Token request failed ({resp.status_code}): {resp.text[:300]}",
              file=sys.stderr)
        sys.exit(1)
    return resp.json()["access_token"]


def _get(session, url: str, params: dict = None) -> dict:
    r = session.get(url, params=params, timeout=_TIMEOUT)
    r.raise_for_status()
    return r.json()


def _patch(session, url: str, body: dict) -> dict:
    r = session.patch(url, json=body, timeout=_TIMEOUT)
    r.raise_for_status()
    return r.json()


def _post(session, url: str, body: dict) -> dict:
    r = session.post(url, json=body, timeout=_TIMEOUT)
    r.raise_for_status()
    return r.json()


def _list_archive_messages(session, mailbox: str, since: "datetime.datetime | None",
                           until: "datetime.datetime | None" = None) -> list:
    """Page through Archive and return messages.

    since=None  → no lower bound; since=<dt> → only messages received >= that datetime.
    until=None  → no upper bound; until=<dt> → only messages received <  that datetime.
    (since + until together give a closed window, e.g. one calendar month.)
    """
    url = f"{GRAPH}/users/{mailbox}/mailFolders/Archive/messages"
    params = {
        "$select": "id,subject,from,receivedDateTime,isRead",
        "$top": _PAGE_SIZE,
        "$orderby": "receivedDateTime asc",
    }
    clauses = []
    if since is not None:
        clauses.append(f"receivedDateTime ge {since.strftime('%Y-%m-%dT%H:%M:%SZ')}")
    if until is not None:
        clauses.append(f"receivedDateTime lt {until.strftime('%Y-%m-%dT%H:%M:%SZ')}")
    if clauses:
        params["$filter"] = " and ".join(clauses)
    messages = []
    while url:
        data = _get(session, url, params)
        messages.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None  # nextLink already encodes all params
    return messages


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--days", type=int, default=31,
        help="Look back this many days from today (default: 31)",
    )
    parser.add_argument(
        "--all", action="store_true", dest="all_emails",
        help="Reset ALL emails in Archive regardless of age (overrides --days)",
    )
    parser.add_argument(
        "--month", metavar="YYYY-MM", default=None,
        help="Reset exactly one calendar month of Archive mail (e.g. --month 2026-03). "
             "Overrides --days/--all. Use for month-by-month replay testing.",
    )
    parser.add_argument(
        "--mailbox", default="apply@driverai.io",
        help="Shared mailbox to operate on (default: apply@driverai.io)",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="List matching emails without making any changes",
    )
    args = parser.parse_args()

    # ── credentials check ───────────────────────────────────────────────────
    tenant = os.getenv("TENANT_ID", "").strip()
    client_id = os.getenv("CLIENT_ID", "").strip()
    client_secret = os.getenv("CLIENT_SECRET", "").strip()
    missing = [k for k, v in {
        "TENANT_ID": tenant,
        "CLIENT_ID": client_id,
        "CLIENT_SECRET": client_secret,
    }.items() if not v]
    if missing:
        print(f"\n[ERROR] Missing environment variables: {', '.join(missing)}")
        print("Set them in HiringAgent_App_P2/.env or export them before running.")
        sys.exit(1)

    until = None
    if args.month:
        try:
            year, mon = (int(x) for x in args.month.split("-"))
            since = datetime.datetime(year, mon, 1)
            until = datetime.datetime(year + 1, 1, 1) if mon == 12 else datetime.datetime(year, mon + 1, 1)
        except (ValueError, TypeError):
            print(f"\n[ERROR] --month must be YYYY-MM (got {args.month!r})")
            sys.exit(1)
        window_str = f"calendar month {args.month}  ({since:%Y-%m-%d} .. {until:%Y-%m-%d})"
    elif args.all_emails:
        since = None
        window_str = "ALL time (no date limit)"
    else:
        since = datetime.datetime.utcnow() - datetime.timedelta(days=args.days)
        window_str = f"last {args.days} days  (since {since.strftime('%Y-%m-%d %H:%M UTC')})"

    print(f"\n{'=' * 60}")
    print(f"  PA trigger reset utility")
    print(f"{'=' * 60}")
    print(f"  Mailbox  : {args.mailbox}")
    print(f"  Window   : {window_str}")
    print(f"  Mode     : {'DRY RUN (no changes)' if args.dry_run else 'LIVE — will move + mark unread'}")
    print(f"{'=' * 60}\n")

    # ── authenticate ────────────────────────────────────────────────────────
    print("Authenticating with Microsoft Graph...")
    token = _get_token(tenant, client_id, client_secret)
    session = requests.Session()
    session.headers.update({
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    })
    print("  OK\n")

    # ── list messages in Archive from window ────────────────────────────────
    print(f"Searching Archive ({window_str})...")
    try:
        messages = _list_archive_messages(session, args.mailbox, since, until)
    except requests.HTTPError as e:
        sc = e.response.status_code
        if sc == 403:
            print(
                f"\n[ERROR] 403 Forbidden — the app registration needs "
                f"Mail.ReadWrite (Application) permission admin-consented.\n"
                f"  Go to Entra admin centre → App registrations → your app "
                f"→ API permissions → Add Microsoft Graph → Application → Mail.ReadWrite → Grant admin consent.",
                file=sys.stderr,
            )
        else:
            print(f"\n[ERROR] {sc}: {e.response.text[:300]}", file=sys.stderr)
        sys.exit(1)

    if not messages:
        print(f"No emails found in Archive ({window_str}). Nothing to do.")
        return

    print(f"Found {len(messages)} email(s).\n")

    # ── print table ─────────────────────────────────────────────────────────
    col_w = (4, 22, 38, 45, 9)
    hdr = f"{'#':<{col_w[0]}}  {'Received (UTC)':<{col_w[1]}}  {'From':<{col_w[2]}}  {'Subject':<{col_w[3]}}  {'Was Read'}"
    print(hdr)
    print("-" * (sum(col_w) + 8))
    for i, msg in enumerate(messages, 1):
        from_addr = (msg.get("from") or {}).get("emailAddress", {}).get("address", "?")
        subject = (msg.get("subject") or "(no subject)")
        recv = msg.get("receivedDateTime", "")[:19].replace("T", " ")
        was_read = "Yes" if msg.get("isRead") else "No"
        print(
            f"{i:<{col_w[0]}}  {recv:<{col_w[1]}}  "
            f"{from_addr[:col_w[2]]:<{col_w[2]}}  "
            f"{subject[:col_w[3]]:<{col_w[3]}}  {was_read}"
        )

    if args.dry_run:
        print(
            f"\n[DRY RUN] Would mark {len(messages)} email(s) as unread and move to Inbox.\n"
            f"Re-run without --dry-run to apply.\n"
            f"  Tip: use --all to include emails older than {args.days} days."
        )
        return

    # ── move + mark unread ──────────────────────────────────────────────────
    print(f"\nProcessing {len(messages)} email(s)...")
    ok, err = 0, 0
    for msg in messages:
        mid = msg["id"]
        subj = (msg.get("subject") or "(no subject)")[:60]
        try:
            # Mark unread BEFORE the move (move changes the message ID in Exchange).
            _patch(
                session,
                f"{GRAPH}/users/{args.mailbox}/messages/{mid}",
                {"isRead": False},
            )
            # Move from Archive to Inbox; PA delta query sees this as a new Inbox message.
            _post(
                session,
                f"{GRAPH}/users/{args.mailbox}/messages/{mid}/move",
                {"destinationId": "Inbox"},
            )
            print(f"  [OK ] {subj}")
            ok += 1
        except requests.HTTPError as e:
            print(f"  [ERR] {subj}: {e.response.status_code} {e.response.text[:100]}")
            err += 1

    # ── summary ─────────────────────────────────────────────────────────────
    print(f"\n{'=' * 60}")
    print(f"  Done.  {ok} email(s) moved to Inbox (unread).  {err} error(s).")
    if ok:
        print(
            f"\n  The PA flow will pick up these {ok} email(s) on the next\n"
            f"  1-minute poll. Check Power Automate run history after ~5 min.\n"
            f"\n  Note: emails that are exact duplicates (same sender, < 90 days)\n"
            f"  will get the 'already on file' reply — not a new row."
        )
    print(f"{'=' * 60}\n")


if __name__ == "__main__":
    main()
