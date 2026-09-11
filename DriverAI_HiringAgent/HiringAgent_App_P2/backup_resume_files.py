"""Download every resume BINARY from SharePoint into a local, verifiable mirror.

backup_sharepoint.py snapshots the SHEET (rows as JSON) and writes resume_inventory.json -
but that inventory is only a LIST of filenames. The actual PDF/DOCX bytes were never stored
locally, so a resume deleted or overwritten in SharePoint was gone. That is not
hypothetical: on 2026-08-06 three resumes had to be reconstructed from the original Archive
mail (CC14 Hamza Asif, 82D3 Vaibhav Pawade, and 44A6, which a case-insensitive path bug in
repair_rejected_month_records.py deleted moments after uploading it). The mailbox saved all
three - but that is luck, not a backup: P1 archives can be cleared, and a resume that
arrived before the retention window simply is not there.

Read-only against SharePoint. The only writes are local files.

DESIGN NOTES (why this differs from the version reverted on 2026-08-04)

  * INCREMENTAL BY CONTENT, NOT BY NAME. The old script skipped whenever a file of the same
    name already existed locally, so a resume REPLACED in SharePoint - exactly what happens
    when a candidate resends, or when P2 renames after a re-score - kept its stale backup
    forever and the skip counter reported success. Here a file is skipped only when its
    recorded size AND sha256 both still match; otherwise it is re-downloaded, and the
    superseded copy is kept under <name>.superseded-<sha8> so a bad overwrite is survivable.

  * VERIFIED AFTER WRITE. Bytes are re-read from disk and hashed before the manifest records
    them, so a truncated or partial write is caught at backup time rather than at restore
    time, when it is too late to matter.

  * ONE ROLLING MIRROR, NOT A TIMESTAMPED COPY PER RUN. A timestamped folder per run meant
    re-downloading all 127 files each time, which is why nobody would run it often enough
    for it to help. This mirrors the live tree in place, so a daily run costs only what
    actually changed.

Usage:
    python backup_resume_files.py              # sync the mirror (default)
    python backup_resume_files.py --verify     # check the mirror, download nothing
    python backup_resume_files.py --into DIR   # use a different mirror location
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path
from urllib.parse import quote

import hiring_agent.config  # noqa: F401  - loads .env and logging
from hiring_agent.config import logger, BASE_DIR
from sharepoint_client import GRAPH, SharePointClient

DEFAULT_MIRROR = BASE_DIR / "P2_Logs" / "resume_backup"
MANIFEST = "manifest.json"


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _enumerate(client: SharePointClient) -> list[dict]:
    """Recursive walk returning {'folder', 'name', 'size'} per file.

    Uses its own $select so `size` comes back - the client's list_folder_children selects
    only name+folder, and this tool must not change a production primitive just to read one
    extra field. Same direct-_req pattern backup_sharepoint.py already uses.
    """
    out: list[dict] = []
    stack = [client.resumes_folder.strip("/")]
    seen: set[str] = set()
    while stack:
        folder = stack.pop()
        if folder.lower() in seen:
            continue
        seen.add(folder.lower())
        url = (f"{GRAPH}/sites/{client.site_id()}/drive/root:/"
               f"{quote(folder, safe='/')}:/children?$select=name,folder,size&$top=200")
        while url:
            try:
                data = client._req("GET", url).json()
            except Exception as e:
                logger.warning(f"       WARNING  : could not list '{folder}': {e}")
                break
            for it in data.get("value", []):
                name = str(it.get("name", "") or "")
                if not name:
                    continue
                if "folder" in it:
                    stack.append(f"{folder}/{name}")
                else:
                    out.append({"folder": folder, "name": name,
                                "size": int(it.get("size", 0) or 0)})
            url = data.get("@odata.nextLink")
    return sorted(out, key=lambda f: f"{f['folder']}/{f['name']}".lower())


def _load_manifest(mirror: Path) -> dict:
    path = mirror / MANIFEST
    if not path.exists():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
        return {e["path"]: e for e in raw.get("files", [])}
    except Exception as e:
        logger.warning(f"       WARNING  : manifest unreadable ({e}) - treating as empty")
        return {}


def _write_manifest(mirror: Path, entries: dict, source: str) -> None:
    (mirror / MANIFEST).write_text(json.dumps({
        "updated_at": dt.datetime.now(dt.timezone.utc).isoformat(),
        "source_folder": source,
        "file_count": len(entries),
        "files": sorted(entries.values(), key=lambda e: e["path"].lower()),
    }, indent=2, ensure_ascii=False), encoding="utf-8")


def verify(mirror: Path) -> int:
    """Check every mirrored file against its recorded hash. Downloads nothing."""
    entries = _load_manifest(mirror)
    if not entries:
        logger.info("  No manifest found - nothing to verify.")
        return 1
    bad = missing = 0
    for rel, e in sorted(entries.items()):
        local = mirror / rel
        if not local.exists():
            logger.warning(f"  MISSING  {rel}")
            missing += 1
            continue
        if _sha256(local.read_bytes()) != e.get("sha256"):
            logger.warning(f"  CORRUPT  {rel} (hash differs from manifest)")
            bad += 1
    logger.info(f"  {len(entries)} in manifest | {missing} missing | {bad} corrupt")
    return 1 if (bad or missing) else 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--into", metavar="DIR", help=f"Mirror location (default: {DEFAULT_MIRROR}).")
    ap.add_argument("--verify", action="store_true",
                    help="Verify the existing mirror against its manifest; download nothing.")
    args = ap.parse_args()

    mirror = Path(args.into) if args.into else DEFAULT_MIRROR
    mirror.mkdir(parents=True, exist_ok=True)

    logger.info("=" * 64)
    logger.info(f"  Resume binary backup - {'VERIFY' if args.verify else 'SYNC'}")
    logger.info("=" * 64)
    logger.info(f"  Mirror: {mirror}")

    if args.verify:
        return verify(mirror)

    client = SharePointClient()
    logger.info(f"  Connected: {client.hostname}")
    remote = _enumerate(client)
    root = client.resumes_folder.strip("/")
    logger.info(f"  {len(remote)} file(s) under '{client.resumes_folder}'")

    entries = _load_manifest(mirror)
    saved = skipped = replaced = failed = 0

    for item in remote:
        folder, name, size = item["folder"], item["name"], item["size"]
        rel_dir = folder[len(root):].strip("/")
        rel = f"{rel_dir}/{name}" if rel_dir else name
        dest = mirror / rel
        prior = entries.get(rel)

        # Skip only when BOTH the recorded size and the on-disk hash still agree - a name
        # match alone is what let the old version keep a stale copy of a replaced resume.
        if prior and dest.exists() and prior.get("size") == size:
            if _sha256(dest.read_bytes()) == prior.get("sha256"):
                skipped += 1
                continue

        try:
            data = client.download_file(folder, name)
        except Exception as e:
            failed += 1
            logger.warning(f"  FAILED   {rel}: {e}")
            continue

        dest.parent.mkdir(parents=True, exist_ok=True)
        digest = _sha256(data)

        # Preserve whatever was there before rather than overwriting it - if SharePoint's
        # copy was itself clobbered, the previous good bytes are the only ones left.
        if dest.exists():
            old = dest.read_bytes()
            if _sha256(old) != digest:
                keep = dest.with_name(f"{dest.name}.superseded-{_sha256(old)[:8]}")
                if not keep.exists():
                    keep.write_bytes(old)
                logger.info(f"  replaced {rel} (previous copy kept as {keep.name})")
                replaced += 1

        dest.write_bytes(data)
        # Verify what actually landed on disk, not what we believe we wrote.
        if _sha256(dest.read_bytes()) != digest:
            failed += 1
            logger.warning(f"  FAILED   {rel}: write verification mismatch")
            continue

        entries[rel] = {"path": rel, "folder": folder, "name": name,
                        "size": len(data), "sha256": digest,
                        "backed_up_at": dt.datetime.now(dt.timezone.utc).isoformat()}
        saved += 1
        logger.info(f"  saved    {rel} ({len(data):,} bytes)")

    _write_manifest(mirror, entries, client.resumes_folder)

    # A file in the manifest but no longer in SharePoint is not an error - it is the whole
    # point. Surface it so a deletion is visible rather than silently normalised away.
    live = set()
    for item in remote:
        rel_dir = item["folder"][len(root):].strip("/")
        live.add(f"{rel_dir}/{item['name']}" if rel_dir else item["name"])
    gone = sorted(set(entries) - live)

    logger.info("")
    logger.info(f"  {saved} saved | {skipped} unchanged | {replaced} replaced | {failed} failed")
    if gone:
        logger.info(f"  {len(gone)} file(s) in the backup are NO LONGER in SharePoint "
                    f"(retained here):")
        for rel in gone[:20]:
            logger.info(f"     {rel}")
    logger.info(f"  Mirror: {mirror}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
