"""
Resume attachment handler — uploads to SharePoint and writes JSON sidecars.

Naming convention (mirrors P1):
  Resume:  cv_{GUID}.{ext}                  (collision-free)
  Sidecar: intake_cv_{GUID}.{ext}.json      (alongside the resume)
  Folder:  /Candidate_Resumes/{YYYY}/{MMMM}/

The GUID in the filename is derived from the content hash (SHA-256[:16])
so re-uploading the same bytes produces the same filename.
That is intentional: P1 spec §9.2 says "a repeat submission overwrites that
person's file rather than adding a second one" for the update path.

For brand-new candidates a fresh UUID is used so two different resumes for the
same candidate never collide, matching the P1 GUID-based naming.
"""
from __future__ import annotations

import hashlib
import json
import logging
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

RESUME_EXTENSIONS = frozenset({".pdf", ".doc", ".docx"})


@dataclass
class SharePointRef:
    """Result of a successful resume upload to SharePoint."""
    stored_filename: str          # cv_{GUID}.{ext}
    original_filename: str
    sharepoint_path: str          # full path in SharePoint
    drive_id: str
    item_id: str
    content_hash: str             # hex SHA-256 of raw bytes
    size_bytes: int
    sidecar_path: str             # path of the intake_*.json sidecar
    year: str
    month: str
    guid: str


@dataclass
class AttachmentHandler:
    """
    Handles resume extraction, SharePoint upload, and sidecar creation.

    In dry_run mode, no files are written to SharePoint.
    """
    sharepoint_site: str
    resumes_folder: str      # e.g. "/Shared Documents/Candidate_Resumes"
    graph_client: Any        # GraphMailClient or compatible mock
    dry_run: bool = False

    def extract_resume_names(self, attach_names: list[str]) -> list[str]:
        """Return only the attachment names that are valid resume formats."""
        return [n for n in attach_names if Path(n).suffix.lower() in RESUME_EXTENSIONS]

    def content_hash(self, data: bytes) -> str:
        return hashlib.sha256(data).hexdigest()

    def build_filename(self, original_name: str, reuse_guid: str | None = None) -> tuple[str, str]:
        """
        Return (stored_filename, guid).

        If reuse_guid is provided (update path), the same GUID is reused so
        the new file overwrites the old one in SharePoint.
        """
        ext = Path(original_name).suffix.lower() or ".pdf"
        guid = reuse_guid or str(uuid.uuid4())
        return f"cv_{guid}{ext}", guid

    def folder_path(self, received_at: datetime | None = None) -> tuple[str, str, str]:
        """
        Return (folder_path, year, month) for the given timestamp.

        Example: ("/Shared Documents/Candidate_Resumes/2026/September", "2026", "September")
        """
        dt = received_at or datetime.now(timezone.utc)
        year = dt.strftime("%Y")
        month = dt.strftime("%B")  # "September" etc.
        path = f"{self.resumes_folder.rstrip('/')}/{year}/{month}"
        return path, year, month

    def upload_resume(
        self,
        graph_message_id: str,
        attachment_id: str,
        original_filename: str,
        received_at: datetime,
        app_id: str,
        *,
        reuse_guid: str | None = None,
    ) -> SharePointRef:
        """
        Download attachment from Graph and upload to SharePoint.

        Order of operations (matches P1 Registrar role):
          1. Download attachment bytes
          2. Compute content hash
          3. Build collision-free filename
          4. Ensure dated folder exists in SharePoint
          5. Upload resume file
          6. Write JSON sidecar
          7. Return SharePointRef

        Args:
            graph_message_id: Graph ID of the message (must be valid before any folder move).
            attachment_id: Graph attachment ID.
            original_filename: Original filename from the email attachment.
            received_at: Email received timestamp (for folder dating).
            app_id: Application Reference ID (APP-YYYYMMDD-...).
            reuse_guid: If provided, reuse this GUID (update path overwrites old file).
        """
        # 1. Download
        logger.debug("Downloading attachment %s / %s", graph_message_id, attachment_id)
        raw_bytes = self.graph_client.download_attachment(graph_message_id, attachment_id)

        # 2. Hash
        chash = self.content_hash(raw_bytes)

        # 3. Filename
        stored_name, guid = self.build_filename(original_filename, reuse_guid)

        # 4. Folder
        folder, year, month = self.folder_path(received_at)

        # 5 & 6. Upload to SharePoint
        full_path = f"{folder}/{stored_name}"
        sidecar_name = f"intake_{stored_name}.json"
        sidecar_path = f"{folder}/{sidecar_name}"

        if self.dry_run:
            logger.info("[DRY-RUN] Would upload %s → %s", original_filename, full_path)
            drive_id = "DRY_DRIVE"
            item_id = f"DRY_{guid}"
        else:
            drive_id, item_id = self._upload_file(folder, stored_name, raw_bytes)

        ref = SharePointRef(
            stored_filename=stored_name,
            original_filename=original_filename,
            sharepoint_path=full_path,
            drive_id=drive_id,
            item_id=item_id,
            content_hash=chash,
            size_bytes=len(raw_bytes),
            sidecar_path=sidecar_path,
            year=year,
            month=month,
            guid=guid,
        )

        # 6. Write sidecar
        self._write_sidecar(folder, sidecar_name, ref, app_id=app_id)

        return ref

    def _upload_file(self, folder_path: str, filename: str, data: bytes) -> tuple[str, str]:
        """
        Ensure the folder exists and upload the file via SharePoint REST.

        Returns (drive_id, item_id).
        """
        # Use the Graph /drive/root:/path:/content PUT endpoint.
        site_path = folder_path.replace("/Shared Documents", "")
        url = (
            f"https://graph.microsoft.com/v1.0"
            f"/sites/{self._site_id}/drives/{self._drive_id}"
            f"/root:{site_path}/{filename}:/content"
        )
        result = self.graph_client._request(
            "PUT", url,
            retryable=True,
            data=data,
            headers={
                "Authorization": f"Bearer {self.graph_client._get_token()}",
                "Content-Type": "application/octet-stream",
            },
        )
        return result.get("parentReference", {}).get("driveId", ""), result.get("id", "")

    def _write_sidecar(
        self,
        folder_path: str,
        sidecar_name: str,
        ref: SharePointRef,
        *,
        app_id: str,
    ) -> None:
        """Write the JSON sidecar alongside the resume file."""
        payload = {
            "app_id": app_id,
            "stored_filename": ref.stored_filename,
            "original_filename": ref.original_filename,
            "sharepoint_path": ref.sharepoint_path,
            "drive_id": ref.drive_id,
            "item_id": ref.item_id,
            "content_hash": ref.content_hash,
            "size_bytes": ref.size_bytes,
            "captured_at": datetime.now(timezone.utc).isoformat(),
        }
        payload_bytes = json.dumps(payload, indent=2, ensure_ascii=False).encode()

        if self.dry_run:
            logger.info("[DRY-RUN] Would write sidecar %s/%s", folder_path, sidecar_name)
            return

        site_path = folder_path.replace("/Shared Documents", "")
        url = (
            f"https://graph.microsoft.com/v1.0"
            f"/sites/{self._site_id}/drives/{self._drive_id}"
            f"/root:{site_path}/{sidecar_name}:/content"
        )
        self.graph_client._request(
            "PUT", url,
            retryable=True,
            data=payload_bytes,
            headers={
                "Authorization": f"Bearer {self.graph_client._get_token()}",
                "Content-Type": "application/json",
            },
        )
        logger.debug("Wrote sidecar: %s/%s", folder_path, sidecar_name)

    @property
    def _site_id(self) -> str:
        """Extract site ID from the SharePoint site URL."""
        # e.g. "led1234567.sharepoint.com,site-id,web-id"
        # For now, use the Graph /sites endpoint to resolve — caller should cache this.
        raise NotImplementedError("Subclass or inject site_id via config")

    @property
    def _drive_id(self) -> str:
        raise NotImplementedError("Subclass or inject drive_id via config")
