"""
Microsoft Graph mail client for the Python P1 intake engine.

Handles:
- Polling the apply@driverai.io inbox (top=1, like the P1 recurrence)
- Downloading attachment bytes
- Moving messages to Archive or Junk Email
- Marking messages as read
- Sending auto-reply emails from the shared mailbox

CRITICAL idempotency rule:
  Use internetMessageId (immutable) as the canonical message key, NEVER
  the Graph Message ID (which changes on folder move). See TDD §11 and
  P2_SQLITE_ARCHITECTURE_PLAN.md.

Auth: App-only client credentials (MSAL). Requires:
  TENANT_ID, CLIENT_ID, CLIENT_SECRET in .env
  Exchange RBAC scope: Mail.ReadWrite scoped to apply@driverai.io only
                       Mail.Send on apply@driverai.io
"""
from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from typing import Any

logger = logging.getLogger(__name__)

# ─── Data models ──────────────────────────────────────────────────────────────

@dataclass
class Attachment:
    attachment_id: str          # Graph attachment ID (mutable; download before move)
    name: str
    content_type: str
    size: int
    is_inline: bool = False
    # Populated after download
    content_bytes: bytes = field(default_factory=bytes, repr=False)


@dataclass
class Message:
    """Immutable snapshot of a Graph mailbox message."""
    # Graph mutable ID (changes on folder move — do NOT use as idempotency key)
    graph_id: str
    # Stable across folder moves — THE idempotency key
    internet_message_id: str
    sender_email: str
    sender_display: str
    subject: str
    body_text: str          # plain-text body (HTML stripped server-side)
    body_html: str          # raw HTML body
    received_at: str        # ISO-8601 UTC string from Graph
    is_read: bool
    attachments: list[Attachment] = field(default_factory=list)
    # Raw Graph JSON (for audit/sidecar)
    _raw: dict = field(default_factory=dict, repr=False)


# ─── Graph client ─────────────────────────────────────────────────────────────

class GraphMailClient:
    """
    Thin wrapper around Microsoft Graph mail endpoints.

    Uses app-only auth (client credentials). Folder moves and mark-as-read
    disable automatic retries to prevent racing on mutable Graph IDs.
    """

    GRAPH_BASE = "https://graph.microsoft.com/v1.0"
    # Retry on throttle (HTTP 429) with exponential backoff + jitter.
    MAX_RETRIES = 5
    RETRY_BASE_SECONDS = 2

    def __init__(self, mailbox: str, tenant_id: str, client_id: str, client_secret: str):
        self.mailbox = mailbox
        self._tenant_id = tenant_id
        self._client_id = client_id
        self._client_secret = client_secret
        self._token: str | None = None
        self._token_expiry: float = 0.0

        # Lazy import so the module loads without msal when running tests with mocks.
        try:
            import msal as _msal
            self._msal = _msal
        except ImportError:
            self._msal = None

        try:
            import requests as _requests
            self._requests = _requests
        except ImportError:
            self._requests = None

    # ── Auth ──────────────────────────────────────────────────────────────────

    def _get_token(self) -> str:
        """Acquire or refresh an app-only access token via MSAL."""
        if self._token and time.time() < self._token_expiry - 60:
            return self._token

        if self._msal is None:
            raise RuntimeError("msal not installed — run: pip install msal")

        app = self._msal.ConfidentialClientApplication(
            self._client_id,
            authority=f"https://login.microsoftonline.com/{self._tenant_id}",
            client_credential=self._client_secret,
        )
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if "access_token" not in result:
            raise RuntimeError(
                f"MSAL token acquisition failed: {result.get('error_description', result)}"
            )
        self._token = result["access_token"]
        self._token_expiry = time.time() + result.get("expires_in", 3600)
        return self._token

    # ── HTTP helpers ──────────────────────────────────────────────────────────

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
        }

    def _get(self, url: str, *, retryable: bool = True, **kwargs) -> Any:
        return self._request("GET", url, retryable=retryable, **kwargs)

    def _post(self, url: str, json_body: dict, *, retryable: bool = True) -> Any:
        return self._request("POST", url, json=json_body, retryable=retryable)

    def _patch(self, url: str, json_body: dict, *, retryable: bool = False) -> Any:
        """Patch is NOT retried by default — mutable IDs may change after a move."""
        return self._request("PATCH", url, json=json_body, retryable=retryable)

    def _request(self, method: str, url: str, *, retryable: bool, **kwargs) -> Any:
        if self._requests is None:
            raise RuntimeError("requests not installed — run: pip install requests")

        attempt = 0
        while True:
            resp = self._requests.request(
                method, url, headers=self._headers(), timeout=30, **kwargs
            )

            if resp.status_code == 429 and retryable and attempt < self.MAX_RETRIES:
                retry_after = int(resp.headers.get("Retry-After", self.RETRY_BASE_SECONDS * (2 ** attempt)))
                import random
                jitter = random.uniform(0, retry_after * 0.1)
                wait = retry_after + jitter
                logger.warning("Graph 429 throttle — retrying in %.1fs (attempt %d/%d)",
                               wait, attempt + 1, self.MAX_RETRIES)
                time.sleep(wait)
                attempt += 1
                continue

            if resp.status_code == 204:
                return None  # no content

            resp.raise_for_status()

            try:
                return resp.json()
            except Exception:
                return resp.content

    # ── Mail operations ───────────────────────────────────────────────────────

    def poll_inbox(self, top: int = 1) -> list[Message]:
        """
        Fetch the top N messages from the inbox (oldest unread first).

        fetchOnlyUnread=false matches P1 — opening a message does not remove it
        from the processing queue. The worker marks it read explicitly after
        all writes succeed.
        """
        url = (
            f"{self.GRAPH_BASE}/users/{self.mailbox}/mailFolders/Inbox/messages"
            f"?$top={top}&$orderby=receivedDateTime asc"
            f"&$select=id,internetMessageId,from,subject,body,bodyPreview,"
            f"receivedDateTime,isRead,hasAttachments"
            f"&$expand=attachments($select=id,name,contentType,size,isInline)"
        )
        data = self._get(url)
        return [self._parse_message(m) for m in data.get("value", [])]

    def get_message(self, internet_message_id: str) -> Message | None:
        """Look up a message by its immutable internetMessageId."""
        # Filter syntax requires URL-encoding of the angle brackets.
        filter_val = f"internetMessageId eq '{internet_message_id}'"
        url = (
            f"{self.GRAPH_BASE}/users/{self.mailbox}/messages"
            f"?$filter={filter_val}&$top=1"
            f"&$expand=attachments($select=id,name,contentType,size,isInline)"
        )
        data = self._get(url)
        msgs = data.get("value", [])
        return self._parse_message(msgs[0]) if msgs else None

    def download_attachment(self, graph_message_id: str, attachment_id: str) -> bytes:
        """Download raw attachment bytes. Call BEFORE moving the message."""
        url = (
            f"{self.GRAPH_BASE}/users/{self.mailbox}/messages/"
            f"{graph_message_id}/attachments/{attachment_id}/$value"
        )
        return self._get(url)

    def mark_read(self, graph_message_id: str) -> None:
        """Mark a message as read. NOT retried — mutable ID safety."""
        url = f"{self.GRAPH_BASE}/users/{self.mailbox}/messages/{graph_message_id}"
        self._patch(url, {"isRead": True}, retryable=False)
        logger.debug("Marked read: %s", graph_message_id)

    def move_message(self, graph_message_id: str, destination_folder: str) -> str:
        """
        Move a message to a well-known folder.

        Returns the NEW Graph message ID (the old one is now invalid).
        Folder names: "Archive", "JunkEmail", "Inbox", "DeletedItems"

        NOT retried — folder moves assign a new mutable Graph ID and retrying
        risks double-moves on a transient failure. The worker handles recovery
        via internetMessageId lookup.
        """
        url = f"{self.GRAPH_BASE}/users/{self.mailbox}/messages/{graph_message_id}/move"
        result = self._post(url, {"destinationId": destination_folder}, retryable=False)
        new_id = result.get("id", "") if result else ""
        logger.debug("Moved %s → %s (new id: %s)", graph_message_id, destination_folder, new_id)
        return new_id

    def send_mail(
        self,
        to_address: str,
        subject: str,
        body_html: str,
        bcc_address: str | None = None,
        save_to_sent: bool = False,
    ) -> None:
        """
        Send an email from the shared mailbox.

        Requires Mail.Send scoped to the apply mailbox.
        BCC: yashv@driverai.io is added to all applicant-facing emails (P1 spec §6).
        """
        message: dict[str, Any] = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": to_address}}],
        }
        if bcc_address:
            message["bccRecipients"] = [{"emailAddress": {"address": bcc_address}}]

        url = f"{self.GRAPH_BASE}/users/{self.mailbox}/sendMail"
        self._post(url, {"message": message, "saveToSentItems": save_to_sent})
        logger.info("Sent mail to %s (subject: %s)", to_address, subject[:60])

    # ── Folder resolution ─────────────────────────────────────────────────────

    def get_folder_id(self, folder_name: str) -> str:
        """Resolve a folder display name to its Graph ID."""
        url = f"{self.GRAPH_BASE}/users/{self.mailbox}/mailFolders"
        data = self._get(url)
        for folder in data.get("value", []):
            if folder.get("displayName", "").lower() == folder_name.lower():
                return folder["id"]
        raise KeyError(f"Folder not found: {folder_name!r}")

    # ── Parsing ───────────────────────────────────────────────────────────────

    @staticmethod
    def _parse_message(raw: dict) -> Message:
        from_addr = raw.get("from", {}).get("emailAddress", {})
        body = raw.get("body", {})
        attachments = [
            Attachment(
                attachment_id=a.get("id", ""),
                name=a.get("name", ""),
                content_type=a.get("contentType", ""),
                size=a.get("size", 0),
                is_inline=a.get("isInline", False),
            )
            for a in raw.get("attachments", [])
            if not a.get("isInline", False)  # skip inline images
        ]
        return Message(
            graph_id=raw.get("id", ""),
            internet_message_id=raw.get("internetMessageId", ""),
            sender_email=(from_addr.get("address") or "").lower().strip(),
            sender_display=from_addr.get("name", ""),
            subject=raw.get("subject", "") or "",
            body_text=raw.get("bodyPreview", ""),
            body_html=body.get("content", ""),
            received_at=raw.get("receivedDateTime", ""),
            is_read=raw.get("isRead", False),
            attachments=attachments,
            _raw=raw,
        )
