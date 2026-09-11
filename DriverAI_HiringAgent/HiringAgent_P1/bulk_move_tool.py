"""
bulk_move_tool.py  —  Helper utility to move emails between Inbox, Junk, and Archive folders.
Use this script to bypass Outlook Web App (OWA) batch selection limits.

SAFETY (added 2026-08-25)
  --to-inbox re-injects the ENTIRE Archive with no date window, so P1 re-runs its full
  intake on every message and re-sends all six applicant emails. It is therefore gated on
  the same replay posture as trigger_reset.py and refuses to run while
  email.send_applicant_emails is true. A dry run is always allowed; --force-send overrides.

  --to-archive no longer moves Junk Email by default. Junk holds screened spam and vendor
  mail; re-filing it to Archive re-presents it as legitimate applicant mail. Pass
  --include-junk if that is genuinely what you want.

USAGE:
  # Move Inbox into Archive (Junk is left alone; no read/unread changes):
  python bulk_move_tool.py --to-archive

  # ...and Junk too (re-files screened spam as legitimate mail):
  python bulk_move_tool.py --to-archive --include-junk

  # Move all messages in Archive into Inbox AND mark them all UNREAD (gated):
  python bulk_move_tool.py --to-inbox

  # Dry-run (Preview counts without moving anything):
  python bulk_move_tool.py --to-archive --dry-run
  python bulk_move_tool.py --to-inbox --dry-run
"""
import argparse
import os
import sys
from pathlib import Path

try:
    from dotenv import load_dotenv
    _env = Path(__file__).resolve().parent.parent / "HiringAgent_P2" / ".env"
    if _env.exists():
        load_dotenv(_env)
        print(f"[env] Loaded credentials from {_env}")
    else:
        load_dotenv()
except ImportError:
    pass

import requests

GRAPH = "https://graph.microsoft.com/v1.0"
_TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
_TIMEOUT = 30
_PAGE_SIZE = 50


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
        print(f"[ERROR] Token request failed ({resp.status_code}): {resp.text[:300]}", file=sys.stderr)
        sys.exit(1)
    return resp.json()["access_token"]


def _list_messages(session, mailbox: str, folder: str) -> list:
    """Enumerate all messages in a specific well-known folder."""
    url = f"{GRAPH}/users/{mailbox}/mailFolders/{folder}/messages"
    params = {
        "$select": "id,subject,receivedDateTime,isRead,from",
        "$top": _PAGE_SIZE
    }
    messages = []
    while url:
        resp = session.get(url, params=params if "$top" in url or "?" not in url else None, timeout=_TIMEOUT)
        if not resp.ok:
            print(f"[ERROR] Failed to fetch folder {folder} ({resp.status_code}): {resp.text[:300]}", file=sys.stderr)
            break
        data = resp.json()
        messages.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None
    return messages


def _only_from(messages: list, sender: str) -> list:
    """Messages whose From address matches `sender` (case-insensitive, exact address).

    Exists so ONE candidate can be requeued without re-injecting the whole Archive. The
    bulk --to-inbox path re-runs P1's full intake on every message it moves, which is how
    498 duplicate applicant emails were sent across three replay runs; requeueing a single
    person is a different operation with a blast radius of one, and needs to be expressible
    as such rather than approximated by moving everything.
    """
    want = sender.strip().lower()
    out = []
    for m in messages:
        addr = (((m.get("from") or {}).get("emailAddress") or {}).get("address") or "").lower()
        if addr == want:
            out.append(m)
    return out


def _move_message(session, mailbox: str, message_id: str, dest_folder: str):
    url = f"{GRAPH}/users/{mailbox}/messages/{message_id}/move"
    resp = session.post(url, json={"destinationId": dest_folder}, timeout=_TIMEOUT)
    resp.raise_for_status()


def _mark_unread(session, mailbox: str, message_id: str):
    url = f"{GRAPH}/users/{mailbox}/messages/{message_id}"
    resp = session.patch(url, json={"isRead": False}, timeout=_TIMEOUT)
    resp.raise_for_status()


# ── shared replay gate ───────────────────────────────────────────────────────
# --to-inbox is trigger_reset.py with no date window: it re-injects the ENTIRE
# Archive, so P1 re-runs its full intake - including all six applicant sends -
# on every message. That is the same failure that produced 498 duplicate sends
# (of 1048 total) up to 2026-08-24. The gate lives in trigger_reset.py; import
# it rather than copy it, so the two tools can never drift apart.
def _replay_gate(force: bool) -> None:
    import importlib.util
    _p = Path(__file__).resolve().parent / "trigger_reset.py"
    try:
        _spec = importlib.util.spec_from_file_location("_tr", _p)
        _tr = importlib.util.module_from_spec(_spec)
        _argv, sys.argv = sys.argv, [str(_p)]      # its module body parses no args, but be safe
        try:
            _spec.loader.exec_module(_tr)
        finally:
            sys.argv = _argv
        _tr._enforce_replay_posture(force)
    except SystemExit:
        raise
    except Exception as e:
        # Never fail OPEN: if the gate cannot be evaluated, refuse the replay.
        print(f"\n[BLOCKED] could not evaluate the replay posture ({e}).")
        print("          Run trigger_reset.py --dry-run to check, or pass --force-send.\n")
        if not force:
            sys.exit(2)


def main():
    parser = argparse.ArgumentParser(description="Bulk Move Utility for hiring mailbox")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--to-archive", action="store_true", help="Move Inbox into Archive (add --include-junk for Junk too)")
    group.add_argument("--to-inbox", action="store_true", help="Move Archive to Inbox and mark unread")
    parser.add_argument("--mailbox", default="apply@driverai.io", help="Target mailbox address")
    parser.add_argument("--dry-run", action="store_true", help="Preview counts without applying changes")
    parser.add_argument("--include-junk", action="store_true",
                        help="With --to-archive, ALSO move Junk Email into Archive. Off by default: "
                             "Junk holds screened spam and vendor mail, and re-filing it to Archive "
                             "re-presents it as legitimate applicant mail.")
    parser.add_argument("--sender", metavar="ADDRESS",
                        help="Only act on messages from this exact address. With --to-inbox "
                             "this requeues ONE candidate instead of the whole Archive - the "
                             "right way to re-run a single person. Note P1 keys the row's "
                             "Full Name and Email off the From header, so requeue the "
                             "candidate's OWN message; forwarding it files the row under YOU.")
    parser.add_argument("--force-send", action="store_true",
                        help="With --to-inbox, replay even though P1 will re-email real candidates.")
    args = parser.parse_args()

    # A dry run changes nothing, so it never needs the gate. Neither does a --sender move:
    # the gate exists to stop a WHOLE-ARCHIVE replay re-emailing hundreds of real candidates,
    # and requeueing one named person is a deliberate act with a blast radius of one - which
    # is usually done precisely BECAUSE the acknowledgment should go out again.
    if args.to_inbox and not args.dry_run and not args.sender:
        _replay_gate(args.force_send)

    tenant = os.environ.get("TENANT_ID")
    client_id = os.environ.get("CLIENT_ID")
    client_secret = os.environ.get("CLIENT_SECRET")

    if not tenant or not client_id or not client_secret:
        print("[ERROR] Missing AAD environment credentials (TENANT_ID, CLIENT_ID, CLIENT_SECRET).", file=sys.stderr)
        print("Please ensure they are defined in HiringAgent_P2/.env", file=sys.stderr)
        sys.exit(1)

    print(f"Connecting to Graph API for mailbox: {args.mailbox}...")
    token = _get_token(tenant, client_id, client_secret)
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {token}"})

    if args.to_archive:
        print("\n--- Scanning folders for Archive move ---")
        inbox_messages = _list_messages(session, args.mailbox, "inbox")
        junk_messages = (_list_messages(session, args.mailbox, "junkemail")
                         if args.include_junk else [])

        print(f"Found {len(inbox_messages)} emails in Inbox.")
        if args.include_junk:
            print(f"Found {len(junk_messages)} emails in Junk Email.")
            print("  [--include-junk] screened spam WILL be re-filed to Archive, where it")
            print("  reads as legitimate applicant mail. Only do this if you mean to.")
        else:
            _junk_n = len(_list_messages(session, args.mailbox, "junkemail"))
            print(f"Skipping Junk Email ({_junk_n} messages) - pass --include-junk to move those too.")
        
        total = len(inbox_messages) + len(junk_messages)
        if total == 0:
            print("Nothing to move. Folders are already empty!")
            return

        if args.dry_run:
            print(f"[DRY-RUN] Would move {total} emails to Archive.")
            return

        print(f"Moving {total} emails to Archive...")
        moved_count = 0
        for msg in inbox_messages:
            try:
                _move_message(session, args.mailbox, msg["id"], "archive")
                moved_count += 1
                if moved_count % 10 == 0 or moved_count == total:
                    print(f"  Processed {moved_count}/{total}...")
            except Exception as e:
                print(f"  [WARN] Failed to move Inbox msg {msg['id']} ({msg.get('subject', 'No Subject')}): {e}")

        for msg in junk_messages:
            try:
                _move_message(session, args.mailbox, msg["id"], "archive")
                moved_count += 1
                if moved_count % 10 == 0 or moved_count == total:
                    print(f"  Processed {moved_count}/{total}...")
            except Exception as e:
                print(f"  [WARN] Failed to move Junk msg {msg['id']} ({msg.get('subject', 'No Subject')}): {e}")
                
        print(f"[OK] Successfully moved {moved_count} emails to Archive.")

    elif args.to_inbox:
        print("\n--- Scanning Archive for Inbox restoration ---")
        archive_messages = _list_messages(session, args.mailbox, "archive")
        if args.sender:
            _before = len(archive_messages)
            archive_messages = _only_from(archive_messages, args.sender)
            print(f"[--sender {args.sender}] {len(archive_messages)} of {_before} Archive "
                  f"message(s) match; the rest are left untouched.")
            if len(archive_messages) == 0:
                print("Nothing to requeue - check the address is exactly as it appears in the "
                      "From header.")
                return
        
        print(f"Found {len(archive_messages)} emails in Archive.")
        if len(archive_messages) == 0:
            print("Nothing to move. Archive is empty!")
            return

        if args.dry_run:
            print(f"[DRY-RUN] Would mark unread and move {len(archive_messages)} emails from Archive to Inbox.")
            return

        print(f"Restoring {len(archive_messages)} emails to Inbox (marking unread)...")
        restored_count = 0
        total = len(archive_messages)
        for msg in archive_messages:
            try:
                # 1. Mark unread
                _mark_unread(session, args.mailbox, msg["id"])
                # 2. Move back to inbox
                _move_message(session, args.mailbox, msg["id"], "inbox")
                restored_count += 1
                if restored_count % 10 == 0 or restored_count == total:
                    print(f"  Processed {restored_count}/{total}...")
            except Exception as e:
                print(f"  [WARN] Failed to restore msg {msg['id']} ({msg.get('subject', 'No Subject')}): {e}")

        print(f"[OK] Successfully restored and marked unread {restored_count} emails to Inbox.")


if __name__ == "__main__":
    main()
