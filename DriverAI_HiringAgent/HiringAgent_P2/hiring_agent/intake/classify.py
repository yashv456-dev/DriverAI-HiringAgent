"""
Intake security classifier — ports all three P1 screening gates from flow_config.json.

Gate 1: Sender substring check (43 terms) — blocks bots, platforms, self-loop
Gate 2: Subject substring check (27 terms) — blocks system mail, OOO, invoices
Gate 3: Content phrase check (123 phrases) — blocks scam, abuse, malware, vendor pitches

Allow-list (4 addresses) — exact match, bypasses all gates.
Link-shortener check — blocked unless a real PDF/DOCX is also attached.

IMPORTANT: Matching is LOWERCASE SUBSTRING for Gates 1/2/3.
Allow-list matching is EXACT EQUALITY (changed from substring 2026-09-08
after integrity_fixes.py audit — see flow_config.json _comment_allow_senders).

Malware extension phrases carry a TRAILING SPACE (e.g. ".exe ") by design.
The scan field appends a trailing space so ".exe " matches "invoice.exe" but
not "www.exeter.ac.uk" (.exet). Do NOT strip these trailing spaces.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Sequence

logger = logging.getLogger(__name__)

# ─── Classification result ─────────────────────────────────────────────────────

class Classification(str, Enum):
    ALLOW_LISTED    = "allow_listed"      # passes all gates unconditionally
    BAD_SENDER      = "bad_sender"        # Gate 1 blocked
    BAD_SUBJECT     = "bad_subject"       # Gate 2 blocked
    SPAM            = "spam"              # Gate 3 blocked
    SHORTENER_ONLY  = "shortener_only"    # link shortener with no real resume
    # Passed gates — routed by route.py
    NEW             = "new"
    DUPLICATE       = "duplicate"
    RESUME_UPDATE   = "resume_update"
    FOLLOWUP        = "followup"
    CV_REQUEST      = "cv_request"
    WRONG_FORMAT    = "wrong_format"
    IGNORED         = "ignored"

BLOCKED = {
    Classification.BAD_SENDER,
    Classification.BAD_SUBJECT,
    Classification.SPAM,
    Classification.SHORTENER_ONLY,
}


@dataclass
class ClassifyResult:
    classification: Classification
    gate_hit: str = ""          # which specific term/phrase matched
    allow_listed: bool = False


# ─── Filter lists — sourced verbatim from flow_config.json (Sep 2026) ─────────

# Exact-equality allow-list (checked BEFORE all gates).
ALLOW_SENDERS: tuple[str, ...] = (
    "yashv@driverai.io",
    "tracys@driverai.io",
    "akhilb@driverai.io",
    "samruddhac@driverai.io",
)

# Gate 1: substring match against lowercased From header.
BAD_SENDERS: tuple[str, ...] = (
    "noreply", "no-reply", "donotreply", "notifications", "mailer-daemon",
    "postmaster", "bounce-", "bounces.", "newsletter", "billing@",
    "teams.mail.microsoft.com", "notify.microsoft.com", "email.microsoftonline.com",
    "linkedin.com", "accounts.google.com", "slack.com", "slack-mail.com",
    "zoom.us", "zoom.com", "calendly.com", "atlassian.net", "github.com",
    "docusign.net", "mailchimp", "sendgrid.net", "salesforce.com",
    "amazonses.com", "facebookmail.com", "dropbox.com", "notion.so",
    "asana.com", "bacancy.com", "equestsolutions.net", "sigmasolve.com",
    "logicrays.com", "infilon.com", "quebecsol.com", "synapsetechservice.com",
    "itidoltechnologies.com", "atgcorp.com", "pushcam-solution.com",
    "fixelsmedia.com",
    "apply@driverai.io",  # self-loop guard — MUST stay here, never in allow-list
)

# Gate 2: substring match against lowercased subject.
# Last 3 entries are P1's own auto-reply subjects — prevent the flow answering itself.
BAD_SUBJECTS: tuple[str, ...] = (
    "sent a message", "notification", "security alert",
    "daily digest", "weekly digest", "monthly digest", "your digest",
    "invoice attached", "your receipt", "unsubscribe",
    "password reset", "confirm your email", "verify your",
    "out of office", "delivery failed", "returned mail",
    "webinar", "meeting recording", "your order", "payment declined",
    "security code", "verification code", "free trial", "take our survey",
    # P1's own auto-reply subjects (self-loop guard for subject line)
    "thanks for applying to driver ai",
    "your driver ai application is already",
    "quick favor for your driver ai application",
)

# Gate 3: phrase match against lowercased(subject + " " + body + " " + joined_attach_names + " ")
# NOTE: malware extension entries carry TRAILING SPACES — do not strip.
SPAM_PHRASES: tuple[str, ...] = (
    # Scam / phishing (30)
    "click here to claim", "verify your account", "account has been suspended",
    "urgent action required", "wire transfer", "lottery", "you have won",
    "crypto investment", "loan offer", "earn from home", "inheritance fund",
    "claim your prize", "click the link below to verify", "click here to confirm",
    "click below to unlock", "click to restore access", "update your payment",
    "update your billing", "your account will be closed", "your account will be suspended",
    "confirm your identity", "login immediately", "sign in immediately",
    "unusual sign-in activity", "unauthorized access detected", "reset your password now",
    "your package is waiting", "you have a pending delivery", "we could not deliver",
    "track your shipment",
    # Offensive / abusive (36)
    "send nudes", "send pics", "send pic", "dick pic", "sex chat",
    "sexual favor", "sexual favour", "sexual favors", "blowjob", "blow job",
    "horny", "pornography", "porn site", "child porn", "xxx video",
    "naked pic", "naked pics", "motherfucker", "mother fucker", "fuck you",
    "fucking idiot", "fucking bitch", "bitch", "slut", "whore",
    "rape you", "i will rape", "molest", "pedophile", "kill yourself",
    "i will kill you", "i will find you", "go to hell", "burn in hell",
    "faggot", "nigger",
    # Malware file extensions — TRAILING SPACE is intentional (20)
    ".exe ", ".scr ", ".bat ", ".cmd ", ".vbs ", ".jar ", ".apk ",
    ".msi ", ".ps1 ", ".lnk ", ".iso ", ".dll ", ".hta ", ".cpl ", ".pif ",
    "download and run", "enable macros", "enable content",
    "disable antivirus", "disable your antivirus",
    # Foreign scam phrases (25)
    "lotería", "has ganado", "transferencia de dinero", "herencia",
    "vous avez gagné", "transfert d'argent", "héritage", "loterie",
    "sie haben gewonnen", "geld überweisen", "erbschaft",
    "você ganhou", "transferência de dinheiro",
    "вы выиграли", "перевод денег", "наследство",
    "中奖", "转账", "彩票", "当選しました", "送金",
    "लॉटरी", "पैसे ट्रांसफर", "لقد فزت", "تحويل الأموال",
    # Vendor solicitation phrases (12) — validated against 401 archive msgs, zero false positives
    "company profile and corporate deck", "capability deck",
    "engagement models", "consultant details", "we have an excellent",
    "resume of my developer", "attached their resume",
    "available for your project", "software development agency",
    "handful of my contacts", "we work as a", "development partner",
)

# Link shortener domains — blocked UNLESS a real PDF/DOCX resume is also present.
LINK_SHORTENER_PHRASES: tuple[str, ...] = (
    "bit.ly/", "tinyurl.com/", "is.gd/", "t.me/", "rebrand.ly/",
    "cutt.ly/", "shorturl.at/", "ow.ly/", "v.gd/", "rb.gy/",
    "s.id/", "urlzs.com/",
)

# Keywords that indicate an application intent (used when no resume attached).
APP_KEYWORDS: tuple[str, ...] = (
    "resume", "cv", "curriculum vitae", "cover letter", "job application",
    "application for", "applying for", "position of", "role of",
    "apply for", "apply to", "apply as", "hiring", "pfa",
    "attached resume", "attached cv", "my resume", "my cv",
    "candidate", "applicant", "job opening",
)

# Accepted resume attachment extensions.
RESUME_EXTENSIONS: frozenset[str] = frozenset({".pdf", ".doc", ".docx"})

# APP-Ref detection pattern — see flow_config.json appref.detect_pattern
_APPREF_RE = re.compile(r"app-20\d{6}-\d{4}-[a-zA-Z0-9]{4}", re.IGNORECASE)


# ─── Classifier ───────────────────────────────────────────────────────────────

class Classifier:
    """
    Stateless security classifier.

    Usage::

        c = Classifier()
        result = c.classify(
            sender_email="jane@example.com",
            subject="Application for Senior Engineer",
            body="Please find my resume attached.",
            attach_names=["jane_doe_resume.pdf"],
        )
        if result.classification in BLOCKED:
            # route to junk / silence
    """

    # ── Public helpers ───────────────────────────────────────────────────────

    def is_allow_listed(self, sender_email: str) -> bool:
        """Exact match against allow-list (case-insensitive)."""
        return sender_email.strip().lower() in {s.lower() for s in ALLOW_SENDERS}

    def gate1_sender(self, sender_email: str) -> str | None:
        """Return the matching bad-sender term, or None if the sender passes."""
        lower = sender_email.lower()
        for term in BAD_SENDERS:
            if term in lower:
                return term
        return None

    def gate2_subject(self, subject: str) -> str | None:
        """Return the matching bad-subject term, or None if the subject passes."""
        lower = subject.lower()
        for term in BAD_SUBJECTS:
            if term in lower:
                return term
        return None

    def gate3_spam(self, body: str, subject: str, attach_names: Sequence[str]) -> str | None:
        """
        Return the matching spam phrase, or None if the content passes.

        The scan field is: lower(subject) + " " + lower(body) + " " +
        lower(join(attach_names)) + " "   ← trailing space terminates attach names.

        The trailing space is required for malware-extension matching:
        ".exe " matches "invoice.exe " but not "www.exeter.ac.uk".
        """
        # Build scan field exactly as the Power Automate flow does.
        attach_str = " ".join(n for n in attach_names)
        scan = f"{subject.lower()} {body.lower()} {attach_str.lower()} "
        for phrase in SPAM_PHRASES:
            if phrase in scan:
                return phrase
        return None

    def has_link_shortener(self, body: str, subject: str) -> str | None:
        """Return the matching shortener domain, or None."""
        scan = f"{subject.lower()} {body.lower()} "
        for phrase in LINK_SHORTENER_PHRASES:
            if phrase in scan:
                return phrase
        return None

    def has_resume_attachment(self, attach_names: Sequence[str]) -> bool:
        """True if at least one attachment has an accepted resume extension."""
        return any(
            Path(n).suffix.lower() in RESUME_EXTENSIONS
            for n in attach_names
        )

    def has_non_resume_attachment(self, attach_names: Sequence[str]) -> bool:
        """True if attachments exist but none are accepted resume formats."""
        return bool(attach_names) and not self.has_resume_attachment(attach_names)

    def has_app_keywords(self, body: str, subject: str) -> bool:
        """True if the message body or subject contains application intent keywords."""
        scan = f"{subject.lower()} {body.lower()}"
        return any(kw in scan for kw in APP_KEYWORDS)

    def extract_quoted_ref(self, subject: str, body: str) -> str | None:
        """
        Scan subject and body for a quoted APP-Ref (e.g. APP-20260909-1230-8A1).
        Returns the first match normalized to uppercase, or None.
        """
        for text in (subject, body):
            m = _APPREF_RE.search(text)
            if m:
                return m.group(0).upper()
        return None

    # ── Gate runner ──────────────────────────────────────────────────────────

    def run_gates(
        self,
        sender_email: str,
        subject: str,
        body: str,
        attach_names: Sequence[str],
    ) -> ClassifyResult:
        """
        Run allow-list + Gates 1–3 + link-shortener check.
        Returns a ClassifyResult. Does NOT decide routing (that's route.py).
        """
        # Allow-list: exact equality, bypasses everything.
        if self.is_allow_listed(sender_email):
            return ClassifyResult(Classification.ALLOW_LISTED, allow_listed=True)

        # Gate 1 — sender
        hit = self.gate1_sender(sender_email)
        if hit:
            logger.debug("Gate1 BLOCKED %s — matched %r", sender_email, hit)
            return ClassifyResult(Classification.BAD_SENDER, gate_hit=hit)

        # Gate 2 — subject
        hit = self.gate2_subject(subject)
        if hit:
            logger.debug("Gate2 BLOCKED subject %r — matched %r", subject[:60], hit)
            return ClassifyResult(Classification.BAD_SUBJECT, gate_hit=hit)

        # Gate 3 — content
        hit = self.gate3_spam(body, subject, attach_names)
        if hit:
            logger.debug("Gate3 BLOCKED — matched %r", hit)
            return ClassifyResult(Classification.SPAM, gate_hit=hit)

        # Link-shortener gate: blocked UNLESS a real resume is attached.
        shortener = self.has_link_shortener(body, subject)
        if shortener and not self.has_resume_attachment(attach_names):
            logger.debug("Shortener BLOCKED (no resume) — matched %r", shortener)
            return ClassifyResult(Classification.SHORTENER_ONLY, gate_hit=shortener)

        # Passed all gates — route.py decides the routing path.
        return ClassifyResult(Classification.NEW)


# Singleton for import convenience.
classifier = Classifier()


# ─── Compatibility shim ───────────────────────────────────────────────────────
# Allow `from pathlib import Path` inside this file without circular import.
from pathlib import Path  # noqa: E402 — must be after the class definition
