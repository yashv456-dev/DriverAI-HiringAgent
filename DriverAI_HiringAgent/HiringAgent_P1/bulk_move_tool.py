"""
bulk_move_tool.py  —  Helper utility to move emails between Inbox, Junk, and Archive folders.
Use this script to bypass Outlook Web App (OWA) batch selection limits.

USAGE:
  # Move all messages in Inbox and Junk Email folders into Archive (No changes to read/unread status):
  python bulk_move_tool.py --to-archive

  # Move all messages in Archive into Inbox AND mark them all UNREAD:
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
    _env = Path(__file__).resolve().parent.parent / "HiringAgent_App_P2" / ".env"
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
        "$select": "id,subject,receivedDateTime,isRead",
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


def _move_message(session, mailbox: str, message_id: str, dest_folder: str):
    url = f"{GRAPH}/users/{mailbox}/messages/{message_id}/move"
    resp = session.post(url, json={"destinationId": dest_folder}, timeout=_TIMEOUT)
    resp.raise_for_status()


def _mark_unread(session, mailbox: str, message_id: str):
    url = f"{GRAPH}/users/{mailbox}/messages/{message_id}"
    resp = session.patch(url, json={"isRead": False}, timeout=_TIMEOUT)
    resp.raise_for_status()


def main():
    parser = argparse.ArgumentParser(description="Bulk Move Utility for hiring mailbox")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--to-archive", action="store_true", help="Move Inbox + Junk into Archive")
    group.add_argument("--to-inbox", action="store_true", help="Move Archive to Inbox and mark unread")
    parser.add_argument("--mailbox", default="apply@driverai.io", help="Target mailbox address")
    parser.add_argument("--dry-run", action="store_true", help="Preview counts without applying changes")
    args = parser.parse_args()

    tenant = os.environ.get("TENANT_ID")
    client_id = os.environ.get("CLIENT_ID")
    client_secret = os.environ.get("CLIENT_SECRET")

    if not tenant or not client_id or not client_secret:
        print("[ERROR] Missing AAD environment credentials (TENANT_ID, CLIENT_ID, CLIENT_SECRET).", file=sys.stderr)
        print("Please ensure they are defined in HiringAgent_App_P2/.env", file=sys.stderr)
        sys.exit(1)

    print(f"Connecting to Graph API for mailbox: {args.mailbox}...")
    token = _get_token(tenant, client_id, client_secret)
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {token}"})

    if args.to_archive:
        print("\n--- Scanning folders for Archive move ---")
        inbox_messages = _list_messages(session, args.mailbox, "inbox")
        junk_messages = _list_messages(session, args.mailbox, "junkemail")
        
        print(f"Found {len(inbox_messages)} emails in Inbox.")
        print(f"Found {len(junk_messages)} emails in Junk Email.")
        
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
