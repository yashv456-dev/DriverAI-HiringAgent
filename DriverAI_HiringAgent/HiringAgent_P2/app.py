"""DriverAI Hiring Agent - desktop app (online-only: SharePoint is the sole store).

A single window over the local bot. Tabs:
  Home             - run/preview/stop, queue count, scheduler, metrics + chart, main log
  Job Descriptions - job-post URLs, one SharePoint/OneDrive JD folder, pasted JD text
  SharePoint       - connection editor (.env) — hostname/site/folder via one pasted URL,
                     credentials, table (sheet) name, workbook filename
  Resumes          - upload local files into the online scoring queue; no-write preview
  Candidates       - loaded live from SharePoint; search, sort, filter, detail, export
  Settings         - Ollama model/host, AI + geo toggles

No candidate data (resumes, workbooks, scored fields) is ever written to this
computer — only operational logs and run-history counts stay local, for diagnostics.
Heavy scoring runs shell out to bot.py so the window stays responsive; the live log
streams here. The scoring engine (hiring_agent/, bot.py) is never modified by the GUI.
"""

import datetime
import json
import os
import queue
import re
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from collections import Counter
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
SETTINGS_FILE = APP_DIR / "app_settings.json"
HISTORY_FILE = APP_DIR / "run_history.jsonl"
APP_VERSION = "1.2"
_INSTANCE_PORT = 47653          # single-instance guard (released automatically on exit)
_HISTORY_MAX = 1000             # rotate the run-history file at this many entries
CLIENT_MODE = True
CLIENT_DISABLED_TABS = (
    "Job Descriptions", "SharePoint", "Resumes", "Candidates", "Settings",
)
_BATCH_SIZE_CHOICES = ("1", "10", "25", "50", "100", "250", "500")

DEFAULT_SETTINGS = {
    "ollama_model": "llama3.2",
    "ollama_host": "http://localhost:11434",
    "use_ai": True,
    "usa_only": True,
    "auto_start": False,
    "trigger_interval_min": 5,
    "scheduled_times": [],
    "schedule_enabled": False,
    "backdate_enabled": True,
    "last_scheduled_run": "",
    "launch_on_startup": False,
    "batch_size": 25,
    "window_geometry": "",
}


def _parse_batch_size(value) -> int:
    """Return a positive candidate limit or raise ValueError."""
    try:
        size = int(str(value).strip())
    except (TypeError, ValueError) as e:
        raise ValueError("Batch size must be a whole number.") from e
    if size < 1:
        raise ValueError("Batch size must be at least 1.")
    return size


def _score_sharepoint_args(batch_size, dry_run: bool = False) -> list[str]:
    """Build the one-shot P2 command shared by manual and scheduled GUI runs."""
    args = ["--score-sharepoint", "--batch-size", str(_parse_batch_size(batch_size))]
    if dry_run:
        args.append("--dry-run")
    return args


def _safe_window_geometry(value: str) -> str:
    """Keep a stale multi-monitor position from opening the app off-screen."""
    raw = str(value or "").strip()
    size = re.match(r"^(\d{3,4})x(\d{3,4})", raw)
    width = int(size.group(1)) if size else 1200
    height = int(size.group(2)) if size else 800
    width = min(max(width, 980), 1800)
    height = min(max(height, 660), 1200)
    # A negative coordinate is often a disconnected left/top monitor. Client mode
    # prioritizes a reliably visible primary-screen launch over preserving that position.
    if not raw or "-" in raw:
        return f"{width}x{height}+80+80"
    return raw

def load_settings() -> dict:
    s = dict(DEFAULT_SETTINGS)
    if SETTINGS_FILE.exists():
        try:
            s.update(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
        except Exception:
            pass
    return s

def save_settings(s: dict) -> None:
    SETTINGS_FILE.write_text(json.dumps(s, indent=2), encoding="utf-8")

_S = load_settings()
os.environ.setdefault("HIRING_OLLAMA_MODEL", _S["ollama_model"])
os.environ.setdefault("HIRING_OLLAMA_HOST", _S["ollama_host"])
os.environ.setdefault("HIRING_OLLAMA_ENABLED", "true" if _S["use_ai"] else "false")
os.environ.setdefault("HIRING_OLLAMA_SCORING", "true" if _S["use_ai"] else "false")
os.environ.setdefault("HIRING_GEO_USA_ONLY", "true" if _S["usa_only"] else "false")

sys.path.insert(0, str(APP_DIR))

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

# ── run history (GUI-side only; the bot is never modified) ───────────────────

def append_history(entry: dict) -> None:
    """Append one run record; rotate the file when it grows past _HISTORY_MAX."""
    try:
        lines = []
        if HISTORY_FILE.exists():
            lines = HISTORY_FILE.read_text(encoding="utf-8").splitlines()
        lines.append(json.dumps(entry))
        if len(lines) > _HISTORY_MAX:
            lines = lines[-_HISTORY_MAX:]
        HISTORY_FILE.write_text("\n".join(lines) + "\n", encoding="utf-8")
    except OSError:
        pass

def load_history(limit: int = 200) -> list[dict]:
    out = []
    try:
        if HISTORY_FILE.exists():
            for ln in HISTORY_FILE.read_text(encoding="utf-8").splitlines()[-limit:]:
                try:
                    out.append(json.loads(ln))
                except Exception:
                    continue
    except OSError:
        pass
    return out

def parse_run_summary(text: str) -> dict:
    """Pull processed/rejected/errors counts out of a bot run's streamed output."""
    counts = {}
    m = re.search(r"Scored\s*:\s*(\d+)", text)
    if m:
        counts["processed"] = int(m.group(1))
        m2 = re.search(r"Rejected\s*:\s*(\d+)", text)
        counts["rejected"] = int(m2.group(1)) if m2 else 0
        m3 = re.search(r"Errors\s*:\s*(\d+)", text)
        counts["errors"] = int(m3.group(1)) if m3 else 0
        return counts
    m = re.search(r"(\d+) added, (\d+) updated, (\d+) rejected", text)
    if m:
        counts["processed"] = int(m.group(1)) + int(m.group(2))
        counts["rejected"] = int(m.group(3))
        m3 = re.search(r"(\d+) error", text)
        counts["errors"] = int(m3.group(1)) if m3 else 0
    return counts


def parse_p2_progress_line(line: str) -> dict:
    """Parse a P2 Progress log line into values the Home progress panel can use."""
    marker = "P2 Progress |"
    if marker not in str(line or ""):
        return {}
    payload = str(line).split(marker, 1)[1]
    event = {}
    for part in payload.split("|"):
        if "=" not in part:
            continue
        key, value = part.split("=", 1)
        key, value = key.strip(), value.strip()
        if key:
            event[key] = value
    for key in (
        "waiting", "batch", "after_batch", "current", "batch_left",
        "queue_left", "row", "attempted", "remaining", "errors",
    ):
        if key in event:
            try:
                event[key] = int(event[key])
            except (TypeError, ValueError):
                pass
    return event if event.get("state") else {}


# ── chart palette (validated against the #0f1117 canvas: CVD dE 66, contrast >=3:1) ──
CHART_SURFACE = "#0f1117"
CHART_BLUE = "#3987e5"     # scored candidates / run duration series
CHART_RED = "#e66767"      # rejected candidates / failed runs
CHART_GRID = "#232733"
CHART_MUTED = "#898781"
CHART_INK = "#c3c2b7"
APP_BG = "#11141c"
CARD_BG = "#171b24"
CARD_ALT = "#1d2230"
CARD_EDGE = "#2a3142"
CARD_TEXT = "#edf2ff"
CARD_SUBTEXT = "#96a0b8"
ACCENT = "#56a4ff"
SUCCESS = "#6ed98d"
WARNING = "#f2be5c"
ERROR = "#f07b72"

# ── type scale (plain specs here; turned into live CTkFont objects in __init__,
# since CTkFont needs a running Tk root to construct) ─────────────────────────
_FONT_SPECS = {
    "caption": ("Segoe UI", 11, "normal"),      # muted help text, KPI captions
    "small_bold": ("Segoe UI", 11, "bold"),     # detail-row labels
    "heading": ("Segoe UI", 15, "bold"),        # card / section titles
    "display": ("Segoe UI", 24, "bold"),        # KPI big numbers
}

def _blend(fg: str, bg: str = CHART_SURFACE, alpha: float = 0.30) -> str:
    """Pre-blend fg toward bg (tk Canvas has no alpha) — used for area fills."""
    f = [int(fg[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(bg[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(fv * alpha + bv * (1 - alpha))
                                   for fv, bv in zip(f, b))

def backfill_history_from_logs() -> int:
    """One-time seed of run_history.jsonl from past P2.log files (summary lines)."""
    try:
        from hiring_agent.config import LOGS_DIR
    except Exception:
        return 0
    added = []
    try:
        logs = sorted(Path(LOGS_DIR).rglob("P2.log*"))
    except OSError:
        return 0
    for log in logs:
        try:
            text = log.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        for m in re.finditer(
                r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}).{0,40}Scored\s*:\s*(\d+)", text):
            seg = text[m.end():m.end() + 400]
            rej = re.search(r"Rejected\s*:\s*(\d+)", seg)
            err = re.search(r"Errors\s*:\s*(\d+)", seg)
            added.append({"ts": m.group(1).replace(" ", "T"), "rc": 0, "secs": 0,
                          "mode": "score-sharepoint", "backfilled": True,
                          "processed": int(m.group(2)),
                          "rejected": int(rej.group(1)) if rej else 0,
                          "errors": int(err.group(1)) if err else 0})
    for e in added:
        append_history(e)
    return len(added)

# ── SharePoint URL parsing (one-line URL -> hostname / site / folder / file) ──

def parse_sharepoint_url(url: str) -> dict:
    """Parse a full SharePoint browser or share URL into its Graph parts.

    Handles %20 encoding, ?csf=1&web=1 query junk, /:f:/r/ share-link prefixes,
    'Shared Documents' (team sites) and 'Documents' (personal/OneDrive) roots,
    and Forms/AllItems.aspx view suffixes. Returns {hostname, site_path,
    folder, file} — folder is drive-relative with a leading slash; any part
    may be "" when absent.
    """
    from urllib.parse import urlparse, unquote
    raw = (url or "").strip()
    if not raw:
        return {"hostname": "", "site_path": "", "folder": "", "file": ""}
    if "://" not in raw:
        raw = "https://" + raw
    p = urlparse(raw)
    hostname = p.hostname or ""    # drops a port if the user pasted one
    path = unquote(p.path)
    # Share-link prefix, e.g. /:f:/r/... or /:x:/g/... — one letter (content type)
    # then one letter (share mode); covers every SharePoint/OneDrive variant seen.
    path = re.sub(r"^/:[a-z]:/[a-z](?=/|$)", "", path)
    segs = [s for s in path.split("/") if s]
    site_path, rest = "", segs
    if segs[:1] in (["sites"], ["teams"], ["personal"]) and len(segs) >= 2:
        site_path = f"/{segs[0]}/{segs[1]}"
        rest = segs[2:]
    if rest and rest[0].lower() in ("shared documents", "documents"):
        rest = rest[1:]
    rest = [s for s in rest
            if s.lower() != "forms" and not s.lower().endswith(".aspx")]
    # Only a recognized document extension counts as "this URL points at a file" —
    # a folder literally named 'Resumes.old' or 'Q1.2024' must stay a folder.
    file = ""
    if rest and re.search(r"\.(xlsx|xls|csv|pdf|docx|doc)$", rest[-1], re.IGNORECASE):
        file = rest[-1]
        rest = rest[:-1]
    folder = "/" + "/".join(rest) if rest else ""
    return {"hostname": hostname, "site_path": site_path, "folder": folder, "file": file}

def compose_sharepoint_url(hostname: str, site_path: str, folder: str) -> str:
    """Rebuild a human-pasteable URL from stored .env parts (best effort)."""
    if not hostname:
        return ""
    lib = "Documents" if site_path.startswith("/personal/") else "Shared Documents"
    folder = ("/" + folder.strip("/")) if folder.strip("/") else ""
    return f"https://{hostname}{site_path}/{lib}{folder}"

def acquire_single_instance() -> bool:
    """Bind a localhost port for the app's lifetime; a second instance fails the bind."""
    global _instance_sock
    try:
        _instance_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        _instance_sock.bind(("127.0.0.1", _INSTANCE_PORT))
        _instance_sock.listen(1)
        return True
    except OSError:
        return False

_instance_sock = None

class _ExemptingScrollableFrame(ctk.CTkScrollableFrame):
    """CTkScrollableFrame whose mousewheel scroll defers to nested widgets that
    scroll themselves (tk.Text, tk.Listbox, ttk.Treeview, ttk.Combobox) instead
    of hijacking the wheel for the outer tab - matching the exemption the old
    hand-rolled canvas-based _scroll_tab() provided. CTkScrollableFrame's own
    _check_if_valid_scroll() walks up from the event widget's .master chain
    looking for this frame's canvas, but only exempts CTkScrollbar/CTkSlider/
    CTkTextbox - it has no notion of the plain ttk/tk widgets this app keeps."""
    _EXEMPT = (tk.Text, tk.Listbox, ttk.Treeview, ttk.Combobox)

    def _check_if_valid_scroll(self, widget):
        if isinstance(widget, self._EXEMPT):
            return False
        return super()._check_if_valid_scroll(widget)

class HiringApp:
    def __init__(self, root):
        self.root = root
        self.fonts = {name: ctk.CTkFont(family=fam, size=sz, weight=w)
                      for name, (fam, sz, w) in _FONT_SPECS.items()}
        self._style_ttk_widgets()
        self.settings = load_settings()
        if CLIENT_MODE:
            # Client mode is deliberately one-shot/scheduled only. Continuous watch,
            # run-on-open, and missed-run catch-up remain available in code but cannot
            # start unexpectedly from an older settings file.
            self.settings["auto_start"] = False
            self.settings["backdate_enabled"] = False
        self._sync_startup_shortcut()
        self.proc = None
        self.log_q: queue.Queue = queue.Queue()
        if not HISTORY_FILE.exists():
            backfill_history_from_logs()   # seed charts from past P2.log summaries
        self._history = load_history()
        self._insights = {
            "queue": None,
            "rows": 0,
            "rejected": 0,
            "categories": [],
            "roles": [],
            "recent_candidates": [],
            "sharepoint_error": "",
            "last_refresh": "",
        }
        self._sp_live_meta = {}
        self._jd_sp_preview = []
        self._jd_items = []
        self._jd_url_status = {}
        root.title(f"DriverAI Hiring Agent  v{APP_VERSION}")
        geo = _safe_window_geometry(self.settings.get("window_geometry"))
        try:
            root.geometry(geo)
        except Exception:
            root.geometry("1200x800+80+80")
        # Tall tabs scroll, so the window may shrink well below the design size
        # without hiding controls (the old 1120x760 floor exceeded the default
        # geometry and still clipped three tabs).
        root.minsize(980, 660)

        self.nb = ctk.CTkTabview(root, command=self._on_tab_changed)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self._build_home_tab()
        self._build_jd_tab()
        self._build_sharepoint_tab()
        self._build_resumes_tab()
        self._build_candidates_tab()
        self._build_settings_tab()
        self._apply_client_mode()

        self._build_status_bar()
        self._drain_log()

        self._sched_fired_today: set = set()
        self._sched_thread = None

        self.root.after(300, self._run_health_check)
        self.root.after(600, self._refresh_queue)
        self.root.after(400, self._refresh_metrics)
        self.root.after(900, self._refresh_insights)
        # Say out loud, at launch, whether opening this window is about to start writing to
        # live SharePoint. auto_start silently kicks off a real scoring loop 500ms after the
        # window appears - no click, no prompt - and nothing on screen said so, which makes
        # "just opening the app to look at the queue" an action with side effects. Logged
        # either way so the quiet case is explicit rather than merely absent.
        _interval = self.settings.get("trigger_interval_min", 5)
        if self.settings.get("auto_start"):
            self._log(self.sp_log,
                      f"AUTO-RUN IS ON - scoring will start automatically in a moment and "
                      f"repeat every {_interval} min, writing to live SharePoint. "
                      f"Turn it off in Settings > Auto start.")
            self.root.after(500, self._sp_autorun)
        else:
            self._log(self.sp_log,
                      "Auto-run is off - this app will not touch SharePoint until you "
                      "press Run. (Settings > Auto start enables it.)")
        if self.settings.get("schedule_enabled") and self.settings.get("scheduled_times"):
            self._log(self.sp_log,
                      f"Scheduler is ON for {self.settings.get('scheduled_times')} - "
                      f"runs will fire unattended at those times.")
            self.root.after(800, self._start_scheduler)
            if self.settings.get("backdate_enabled"):
                self.root.after(1200, self._backdate_check)

    # ── shared helpers ───────────────────────────────────────────────────────
    def _style_ttk_widgets(self):
        """Retheme the ttk widgets kept permanently alongside CustomTkinter
        (Treeview, Combobox, Spinbox, Scrollbar, Progressbar - no CTk
        equivalent exists for any of them). CTk only styles its own CTk*
        widget classes - it never touches the ttk Style engine, so without
        this these widgets would revert to the light OS-default ttk theme.
        "clam" is the only stock base theme that reliably honors
        background/foreground overrides on Windows ("vista"/"winnative"
        mostly ignore them)."""
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure("Treeview", background=CARD_ALT, fieldbackground=CARD_ALT,
                         foreground=CARD_TEXT, bordercolor=CARD_EDGE,
                         borderwidth=0, rowheight=26)
        style.map("Treeview", background=[("selected", ACCENT)],
                   foreground=[("selected", "#0b0d12")])
        style.configure("Treeview.Heading", background=CARD_BG, foreground=CARD_SUBTEXT,
                         relief="flat", borderwidth=1)
        style.map("Treeview.Heading", background=[("active", CARD_ALT)])

        style.configure("TCombobox", fieldbackground=CARD_ALT, background=CARD_BG,
                         foreground=CARD_TEXT, arrowcolor=CARD_SUBTEXT,
                         bordercolor=CARD_EDGE, selectbackground=CARD_ALT,
                         selectforeground=CARD_TEXT)
        style.map("TCombobox", fieldbackground=[("readonly", CARD_ALT)],
                   foreground=[("readonly", CARD_TEXT)])

        style.configure("TSpinbox", fieldbackground=CARD_ALT, background=CARD_BG,
                         foreground=CARD_TEXT, arrowcolor=CARD_SUBTEXT,
                         bordercolor=CARD_EDGE)

        style.configure("TScrollbar", background=CARD_BG, troughcolor=APP_BG,
                         bordercolor=APP_BG, arrowcolor=CARD_SUBTEXT)
        style.map("TScrollbar", background=[("active", CARD_ALT)])

        style.configure("TProgressbar", troughcolor=CARD_ALT, background=ACCENT,
                         bordercolor=CARD_ALT, lightcolor=ACCENT, darkcolor=ACCENT)

    def _make_log(self, parent, height: int = 12) -> tk.Text:
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        bar = ctk.CTkFrame(frame, fg_color="transparent")
        bar.pack(fill=tk.X)
        txt = tk.Text(frame, height=height, wrap=tk.WORD, state=tk.DISABLED,
                      bg=CHART_SURFACE, fg=CARD_TEXT, insertbackground=CARD_TEXT,
                      relief=tk.FLAT, borderwidth=0, highlightthickness=0)
        txt.tag_configure("err", foreground=ERROR)
        txt.tag_configure("warn", foreground=WARNING)
        ctk.CTkButton(bar, text="Clear", width=60, height=24,
                      command=lambda: self._log_clear(txt)).pack(side=tk.RIGHT, padx=2)
        ctk.CTkButton(bar, text="Copy", width=54, height=24,
                      command=lambda: self._log_copy(txt)).pack(side=tk.RIGHT, padx=2)
        ctk.CTkButton(bar, text="Open log folder", width=120, height=24,
                      command=self._open_log_folder).pack(side=tk.RIGHT, padx=2)
        sb = ctk.CTkScrollbar(frame, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        return txt

    def _log(self, widget, line: str):
        text = line.rstrip()
        low = text.lower()
        tag = ()
        if "error" in low or "failed" in low or "[err" in low:
            tag = ("err",)
        elif "warning" in low:
            tag = ("warn",)
        widget.configure(state=tk.NORMAL)
        widget.insert(tk.END, text + "\n", tag)
        widget.see(tk.END)
        widget.configure(state=tk.DISABLED)

    def _log_clear(self, widget):
        widget.configure(state=tk.NORMAL)
        widget.delete("1.0", tk.END)
        widget.configure(state=tk.DISABLED)

    def _log_copy(self, widget):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(widget.get("1.0", tk.END))
        except Exception:
            pass

    def _open_log_folder(self):
        try:
            from hiring_agent.config import LOGS_DIR
            folder = Path(LOGS_DIR)
        except Exception:
            folder = APP_DIR
        try:
            folder.mkdir(parents=True, exist_ok=True)
            # os.startfile exists only on Windows. The team runs macOS too, and an
            # AttributeError here used to be swallowed by the except below, so the
            # button silently did nothing rather than opening anything.
            if sys.platform == "win32":
                os.startfile(str(folder))  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.run(["open", str(folder)], check=False)
            else:
                subprocess.run(["xdg-open", str(folder)], check=False)
        except Exception:
            pass

    def _copy_text(self, value: str):
        text = (value or "").strip()
        if not text:
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except Exception:
            pass

    def _open_link(self, value: str):
        target = (value or "").strip()
        if not target:
            return
        if "://" not in target:
            target = "https://" + target
        try:
            webbrowser.open(target)
        except Exception:
            pass

    def _card(self, parent, title: str, subtitle: str = "", minheight: int = 96,
              compact: bool = False):
        """Rounded card that grows with its content. minheight is kept for call
        compatibility but only content decides the height; side-by-side cards
        equalize via fill=BOTH. compact=True tightens padding for dense tiles
        (KPI row). Returns (frame, frame) - callers historically got a separate
        outer/body pair (the old outer+body double-Frame border hack); a single
        CTkFrame with corner_radius/border_width replaces both."""
        del minheight
        pad = (10, 8) if compact else (14, 12)
        frame = ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=10,
                              border_width=1, border_color=CARD_EDGE)
        inner = ctk.CTkFrame(frame, fg_color="transparent")
        inner.pack(fill=tk.BOTH, expand=True, padx=pad[0], pady=pad[1])
        ctk.CTkLabel(inner, text=title, fg_color="transparent", text_color=CARD_TEXT,
                     font=self.fonts["small_bold"] if compact else self.fonts["heading"],
                     anchor="w").pack(anchor=tk.W, fill=tk.X)
        if subtitle:
            ctk.CTkLabel(inner, text=subtitle, fg_color="transparent", text_color=CARD_SUBTEXT,
                         font=self.fonts["caption"], anchor="w", justify="left",
                         wraplength=360).pack(anchor=tk.W, fill=tk.X, pady=(2, 0))
        return frame, inner

    def _scroll_tab(self, tab_text: str) -> ctk.CTkFrame:
        """Notebook tab whose content scrolls vertically when it can't fit.

        Home / Job Descriptions / SharePoint stack more content than a laptop
        window is tall; without this their bottom sections (run log, source
        details, the Save/Test buttons) were simply cut off.
        """
        self.nb.add(tab_text)
        holder = self.nb.tab(tab_text)
        scrollable = _ExemptingScrollableFrame(holder, fg_color="transparent")
        scrollable.pack(fill=tk.BOTH, expand=True)
        return scrollable

    def _on_tab_changed(self, _event=None):
        """First visit to Candidates auto-loads the table (when configured)."""
        try:
            current = self.nb.get()
        except Exception:
            return
        if CLIENT_MODE and current != "Home":
            self.nb.set("Home")
            return
        if current != "Candidates" or getattr(self, "_cand_autoloaded", False):
            return
        if self.candidate_rows or getattr(self, "_cand_loading", False):
            return
        try:
            from hiring_agent.config import SHAREPOINT_CONFIGURED
        except Exception:
            return
        if SHAREPOINT_CONFIGURED:
            self._cand_autoloaded = True
            self._load_candidates_online()

    def _apply_client_mode(self):
        """Lock navigation to the client-facing Home run/schedule workflow."""
        if not CLIENT_MODE:
            return
        self.nb.set("Home")
        buttons = getattr(getattr(self.nb, "_segmented_button", None),
                          "_buttons_dict", {})
        for tab_name in CLIENT_DISABLED_TABS:
            button = buttons.get(tab_name)
            if button is not None:
                button.configure(state=tk.DISABLED)

    def _detail_row(self, parent, label: str, variable: tk.StringVar,
                    *, copyable: bool = False, linkable: bool = False):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill=tk.X, pady=2)
        ctk.CTkLabel(row, text=label, fg_color="transparent", text_color=CARD_SUBTEXT,
                     font=self.fonts["small_bold"], width=140, anchor="nw",
                     justify="left").pack(side=tk.LEFT, anchor=tk.N, padx=(0, 8))
        value_box = ctk.CTkFrame(row, fg_color=CARD_ALT, corner_radius=6,
                                  border_width=1, border_color=CARD_EDGE)
        value_box.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ctk.CTkLabel(value_box, textvariable=variable, fg_color="transparent", text_color=CARD_TEXT,
                     font=self.fonts["caption"], wraplength=760, justify="left",
                     anchor="w").pack(fill=tk.X, expand=True, padx=10, pady=4)
        if not (copyable or linkable):
            return
        buttons = ctk.CTkFrame(row, fg_color="transparent")
        buttons.pack(side=tk.RIGHT, anchor=tk.N)
        open_btn = None
        if linkable:
            open_btn = ctk.CTkButton(
                buttons, text="Open", width=54, height=24,
                command=lambda v=variable: self._open_link(v.get()))
            open_btn.pack(side=tk.LEFT, padx=(0, 4))
        copy_btn = None
        if copyable:
            copy_btn = ctk.CTkButton(
                buttons, text="Copy", width=54, height=24,
                command=lambda v=variable: self._copy_text(v.get()))
            copy_btn.pack(side=tk.LEFT)

        def refresh_buttons(*_args):
            value = (variable.get() or "").strip()
            enabled = value not in {"", "-", "(none)", "(none found)",
                                    "(discover after Test Connection)"}
            if copy_btn is not None:
                copy_btn.configure(state=tk.NORMAL if enabled else tk.DISABLED)
            if open_btn is not None:
                open_btn.configure(state=tk.NORMAL if enabled else tk.DISABLED)

        variable.trace_add("write", refresh_buttons)
        refresh_buttons()

    def _compose_workbook_path(self) -> str:
        explicit = self.sp_workbook_var.get().strip()
        if explicit:
            return explicit if explicit.startswith("/") else "/" + explicit
        return "/P1P2_SharePoint_Master_Files/Sharepoint_Master_File.xlsx"

    def _compose_workbook_url(self) -> str:
        return compose_sharepoint_url(
            self.sp_hostname_var.get().strip(),
            self.sp_site_var.get().strip(),
            self._compose_workbook_path(),
        )

    def _common_flags(self) -> list:
        s = self.settings
        flags = ["--model", s["ollama_model"], "--host", s["ollama_host"]]
        flags += ["--ai"] if s["use_ai"] else ["--no-ai"]
        if not s["usa_only"]:
            flags += ["--no-geo"]
        return flags

    def _run_bot(self, args: list, log_widget, interactive: bool = True):
        if getattr(self, "_bot_running", False):
            if interactive:
                messagebox.showwarning("Busy", "A run is already in progress.")
            else:
                self.log_q.put((log_widget, "[skip] A run is already in progress."))
            return
        self._bot_running = True  # cleared by _on_run_complete on the UI thread
        self._set_busy(True)
        cmd = [sys.executable, str(APP_DIR / "bot.py")] + args + self._common_flags()
        self._log(log_widget, f"$ {' '.join(cmd)}\n")
        mode = args[0].lstrip("-") if args else "run"
        if mode == "score-sharepoint":
            self._reset_p2_progress()
        started = time.monotonic()

        def worker():
            buffered = []
            try:
                self.proc = subprocess.Popen(
                    cmd, cwd=str(APP_DIR), stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT, text=True, bufsize=1,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                for line in self.proc.stdout:
                    buffered.append(line)
                    self.log_q.put((log_widget, line))
                self.proc.wait()
                rc = self.proc.returncode
                self.log_q.put((log_widget, f"\n[exit code {rc}]\n"))
            except Exception as e:
                self.log_q.put((log_widget, f"\n[error] {e}\n"))
                rc = -1
            finally:
                # The completion payload is what clears _bot_running on the UI thread, so it
                # must be posted even if summarising the output throws - otherwise the app
                # stays stuck on "a run is already in progress" until it is restarted.
                secs = round(time.monotonic() - started, 1)
                try:
                    counts = parse_run_summary("".join(buffered[-80:]))
                except Exception:
                    counts = {}
                self.log_q.put((None, {"rc": rc, "secs": secs, "mode": mode,
                                       "dry": "--dry-run" in args, **counts}))

        threading.Thread(target=worker, daemon=True).start()

    def _drain_log(self):
        try:
            while True:
                widget, payload = self.log_q.get_nowait()
                if widget is not None:
                    progress = parse_p2_progress_line(payload)
                    if progress:
                        self._apply_p2_progress(progress)
                    self._log(widget, payload)
                else:
                    self._on_run_complete(payload)
        except queue.Empty:
            pass
        self.root.after(120, self._drain_log)

    def _reset_p2_progress(self):
        if not hasattr(self, "run_progress_title_var"):
            return
        self.run_progress_title_var.set("Starting P2...")
        self.run_progress_detail_var.set("Checking SharePoint and reading the queue")
        self.run_progress_left_var.set("Waiting for queue count")
        self.run_progress_bar.set(0)

    def _apply_p2_progress(self, event: dict):
        if not hasattr(self, "run_progress_title_var"):
            return
        state = event.get("state", "")
        if state == "queue":
            waiting = int(event.get("waiting", 0) or 0)
            batch = int(event.get("batch", 0) or 0)
            after = int(event.get("after_batch", 0) or 0)
            self.run_progress_title_var.set(f"{batch} candidate(s) selected")
            self.run_progress_detail_var.set(f"{waiting} waiting at the start of this run")
            self.run_progress_left_var.set(f"{after} expected to remain after this batch")
            self.run_progress_bar.set(0)
            return
        if state in {"candidate", "done"}:
            current = int(event.get("current", 0) or 0)
            batch = max(1, int(event.get("batch", 1) or 1))
            batch_left = int(event.get("batch_left", 0) or 0)
            queue_left = int(event.get("queue_left", 0) or 0)
            ref = str(event.get("ref", "") or "(no reference)")
            name = str(event.get("name", "") or "(name not extracted)")
            row = event.get("row", "?")
            if state == "done":
                result = str(event.get("result", "") or "Completed")
                self.run_progress_title_var.set(f"Completed {current} of {batch}: {result}")
                activity = "Candidate complete"
            else:
                self.run_progress_title_var.set(f"Candidate {current} of {batch}")
                activity = str(event.get("activity", "") or "Processing")
            self.run_progress_detail_var.set(f"{ref} | {name} | Row at start: {row}")
            self.run_progress_left_var.set(
                f"{activity} | {batch_left} left in this run | "
                f"about {queue_left} left in the queue"
            )
            self.run_progress_bar.set(min(max(current / batch, 0), 1))
            return
        if state == "finish":
            attempted = int(event.get("attempted", 0) or 0)
            remaining = int(event.get("remaining", 0) or 0)
            errors = int(event.get("errors", 0) or 0)
            self.run_progress_title_var.set("Batch finished")
            self.run_progress_detail_var.set(
                f"{attempted} attempted" + (f" | {errors} error(s)" if errors else "")
            )
            self.run_progress_left_var.set(f"{remaining} candidate(s) still waiting")
            self.run_progress_bar.set(1 if attempted else 0)
            return
        if state == "empty":
            self.run_progress_title_var.set("Queue is empty")
            self.run_progress_detail_var.set("No new or updated candidates to process")
            self.run_progress_left_var.set("0 candidates waiting")
            self.run_progress_bar.set(0)

    def _set_busy(self, busy: bool):
        try:
            if busy:
                self.run_btn.configure(state=tk.DISABLED)
                self.run_spinner.start(12)
            else:
                self.run_btn.configure(state=tk.NORMAL)
                self.run_spinner.stop()
        except Exception:
            pass

    # ═══════════════════════════════════════════════════════════════════════════
    # Tab 1: Home — run, queue, scheduler, metrics, main log
    # ═══════════════════════════════════════════════════════════════════════════

    def _daily_counts(self, days: int = 14) -> list:
        scored: dict = {}
        rejected: dict = {}
        for h in self._history:
            if h.get("dry"):
                continue
            d = str(h.get("ts", ""))[:10]
            if len(d) == 10:
                scored[d] = scored.get(d, 0) + int(h.get("processed", 0) or 0)
                rejected[d] = rejected.get(d, 0) + int(h.get("rejected", 0) or 0)
        today = datetime.date.today()
        out = []
        for i in range(days - 1, -1, -1):
            d = (today - datetime.timedelta(days=i)).isoformat()
            out.append((d, scored.get(d, 0), rejected.get(d, 0)))
        return out

    def _draw_area_chart(self):
        c = self.area_canvas
        c.delete("all")
        self._area_pts = []
        w = c.winfo_width() or 460
        h = int(c.winfo_height() or 110)
        data = self._daily_counts(14)
        max_y = max((s + r) for _, s, r in data)
        if max_y <= 0:
            c.create_text(w // 2, h // 2, fill="#556", font=("Segoe UI", 8),
                          text="No run data yet")
            return
        pl, pr, pt, pb = 8, 8, 16, 14
        n = len(data)

        def X(i):
            return pl + i * (w - pl - pr) / (n - 1)

        def Y(v):
            return h - pb - v * (h - pt - pb) / max_y

        base = h - pb
        for gv in (0.5, 1.0):
            gy = Y(max_y * gv)
            c.create_line(pl, gy, w - pr, gy, fill=CHART_GRID)
        # scored band: filled polygon + 2px top line
        top_s = [(X(i), Y(s)) for i, (_, s, _r) in enumerate(data)]
        c.create_polygon(pl, base, *[v for p in top_s for v in p], w - pr, base,
                         fill=_blend(CHART_BLUE), outline="")
        c.create_line(*[v for p in top_s for v in p], fill=CHART_BLUE, width=2)
        # rejected band stacked above, with a 2px surface gap
        top_t = [(X(i), Y(s + r)) for i, (_, s, r) in enumerate(data)]
        bottom = [(X(i), Y(s) - 2) for i, (_, s, _r) in enumerate(data)]
        c.create_polygon(*[v for p in (top_t + bottom[::-1]) for v in p],
                         fill=_blend(CHART_RED), outline="")
        c.create_line(*[v for p in top_t for v in p], fill=CHART_RED, width=2)
        for i, (d, s, r) in enumerate(data):
            if s or r:
                c.create_oval(X(i) - 2, Y(s + r) - 2, X(i) + 2, Y(s + r) + 2,
                              fill=CHART_RED if r else CHART_BLUE, outline="")
            label = datetime.date.fromisoformat(d).strftime("%b %d")
            self._area_pts.append((X(i), f"{label} — {s} scored, {r} rejected"))
        lx = w - pr - 132
        c.create_rectangle(lx, 4, lx + 8, 12, fill=CHART_BLUE, outline="")
        c.create_text(lx + 12, 8, anchor=tk.W, text="Scored", fill=CHART_INK,
                      font=("Segoe UI", 8))
        c.create_rectangle(lx + 66, 4, lx + 74, 12, fill=CHART_RED, outline="")
        c.create_text(lx + 78, 8, anchor=tk.W, text="Rejected", fill=CHART_INK,
                      font=("Segoe UI", 8))
        c.create_text(pl, h - 6, anchor=tk.W, fill=CHART_MUTED, font=("Segoe UI", 7),
                      text=datetime.date.fromisoformat(data[0][0]).strftime("%b %d"))
        c.create_text(w - pr, h - 6, anchor=tk.E, fill=CHART_MUTED,
                      font=("Segoe UI", 7), text="today")

    def _draw_duration_chart(self):
        c = self.dur_canvas
        c.delete("all")
        self._dur_pts = []
        w = c.winfo_width() or 460
        h = int(c.winfo_height() or 110)
        runs = [x for x in self._history if not x.get("dry") and x.get("secs")][-20:]
        if not runs:
            c.create_text(w // 2, h // 2, fill="#556", font=("Segoe UI", 8),
                          text="Timed runs appear here after the next run")
            return
        pl, pr, pt, pb = 8, 8, 16, 14
        max_s = max(x["secs"] for x in runs)
        n = len(runs)

        def X(i):
            return pl + (i * (w - pl - pr) / (n - 1) if n > 1 else (w - pl - pr) / 2)

        def Y(v):
            return h - pb - v * (h - pt - pb) / max_s

        base = h - pb
        for gv in (0.5, 1.0):
            gy = Y(max_s * gv)
            c.create_line(pl, gy, w - pr, gy, fill=CHART_GRID)
        pts = [(X(i), Y(x["secs"])) for i, x in enumerate(runs)]
        c.create_polygon(pl, base, *[v for p in pts for v in p], X(n - 1), base,
                         fill=_blend(CHART_BLUE), outline="")
        if n > 1:
            c.create_line(*[v for p in pts for v in p], fill=CHART_BLUE, width=2)
        for i, x in enumerate(runs):
            ok = x.get("rc") == 0
            c.create_oval(X(i) - 3, Y(x["secs"]) - 3, X(i) + 3, Y(x["secs"]) + 3,
                          fill=CHART_BLUE if ok else CHART_RED, outline="")
            when = str(x.get("ts", ""))[5:16].replace("T", " ")
            self._dur_pts.append(
                (X(i), f"{when} — {x['secs']:.0f}s, {'ok' if ok else 'FAILED'}"))
        c.create_text(pl, 8, anchor=tk.W, fill=CHART_MUTED, font=("Segoe UI", 7),
                      text=f"max {max_s:.0f}s")

    def _chart_hover(self, canvas, points, event):
        canvas.delete("hover")
        if not points:
            return
        px, label = min(points, key=lambda p: abs(p[0] - event.x))
        if abs(px - event.x) > 40:
            return
        h = int(canvas.winfo_height() or 110)
        w = int(canvas.winfo_width() or 460)
        canvas.create_line(px, 14, px, h - 14, fill="#3a4152", tags="hover")
        tx = min(max(px, 90), w - 90)
        canvas.create_text(tx, 3, anchor=tk.N, text=label, fill="#d6deeb",
                           font=("Segoe UI", 8), tags="hover")

    # ── Manual run / interval ──
    def _selected_batch_size(self, interactive: bool = True):
        raw = (self.batch_size_var.get() if hasattr(self, "batch_size_var")
               else self.settings.get("batch_size", 25))
        try:
            size = _parse_batch_size(raw)
        except ValueError as e:
            if interactive:
                messagebox.showwarning("Invalid batch size", str(e))
            return None
        self.settings["batch_size"] = size
        save_settings(self.settings)
        return size

    def _scheduled_score_args(self) -> list[str]:
        try:
            size = _parse_batch_size(self.settings.get("batch_size", 25))
        except ValueError:
            size = 25
            self.settings["batch_size"] = size
        return _score_sharepoint_args(size)

    def _sp_run(self, dry):
        batch_size = self._selected_batch_size()
        if batch_size is None:
            return
        self.sp_status.configure(
            text="  Trying it out..." if dry else "  Scoring...",
            text_color=SUCCESS,
        )
        args = _score_sharepoint_args(batch_size, dry_run=dry)
        self._run_bot(args, self.sp_log)

    def _sp_autorun(self):
        try:
            mins = max(1, int(self.sp_interval.get()))
        except ValueError:
            mins = 5
            self.sp_interval.set("5")
        self._save_trigger_settings()
        batch_size = self.settings.get("batch_size", 25)
        self._run_bot(["--score-sharepoint", "--batch-size", str(batch_size),
                       "--watch", "--interval", str(mins * 60)],
                      self.sp_log, interactive=False)
        self.sp_status.configure(text=f"  Scoring every {mins} min", text_color=SUCCESS)

    def _sp_export_results(self):
        """Rebuild the client-facing result sheet without running a scoring pass.

        Goes through bot.py --export-results like every other action here, so it reuses the
        same busy-guard, log streaming and completion handling rather than doing SharePoint
        work on the UI thread.
        """
        self._run_bot(["--export-results"], self.sp_log)

    def _sp_stop(self):
        if not (self.proc and self.proc.poll() is None):
            self._log(self.sp_log, "Nothing running.")
            return
        # Rejecting or restoring a candidate is two Graph calls (add to one sheet, delete
        # from the other). Killing the worker between them leaves that candidate on BOTH
        # sheets, which then needs the dedup/merge pass or a manual repair to sort out - so
        # make the cost explicit instead of stopping instantly on a single click.
        if not messagebox.askyesno(
                "Stop the run?",
                "Stop scoring now?\n\n"
                "The candidate currently being written may be left half-recorded "
                "(on both the main and Rejected sheets), which needs a re-run or a "
                "repair pass to clean up.\n\n"
                "Candidates already finished are unaffected."):
            return
        self.proc.terminate()
        self._log(self.sp_log, "\n[stopped]\n")
        self.sp_status.configure(text="  Stopped", text_color=ERROR)

    # ── Scheduled triggers ──
    def _format_sched_label(self, hhmm: str) -> str:
        h, m = map(int, hhmm.split(":"))
        return datetime.time(h, m).strftime("%I:%M %p").lstrip("0")

    def _normalize_schedule_time(self, value: str):
        raw = value.strip().lower()
        if not raw:
            return None
        raw = raw.replace(".", ":").replace(" ", "")
        match = re.fullmatch(r"(\d{1,2})(?::(\d{1,2}))?([ap]m)?", raw)
        if not match:
            return None
        hour = int(match.group(1))
        minute = int(match.group(2) or "0")
        ampm = match.group(3)
        if minute > 59:
            return None
        if ampm:
            if hour < 1 or hour > 12:
                return None
            if ampm == "am":
                hour = 0 if hour == 12 else hour
            else:
                hour = 12 if hour == 12 else hour + 12
        elif hour > 23:
            return None
        return f"{hour:02d}:{minute:02d}"

    def _startup_script_path(self):
        appdata = os.getenv("APPDATA")
        if not appdata:
            return None
        return (Path(appdata) / "Microsoft" / "Windows" / "Start Menu" /
                "Programs" / "Startup" / "DriverAI Hiring Agent P2.cmd")

    def _sync_startup_shortcut(self):
        startup_script = self._startup_script_path()
        if not startup_script:
            return
        enabled = bool(self.settings.get("launch_on_startup"))
        try:
            if enabled:
                if getattr(sys, "frozen", False):
                    launch_cmd = f'start "" "{sys.executable}"'
                else:
                    pythonw = Path(sys.executable).with_name("pythonw.exe")
                    launcher = pythonw if pythonw.exists() else Path(sys.executable)
                    launch_cmd = f'start "" "{launcher}" "{APP_DIR / "app.py"}"'
                script = "\r\n".join([
                    "@echo off",
                    f'cd /d "{APP_DIR}"',
                    launch_cmd,
                ]) + "\r\n"
                startup_script.write_text(script, encoding="ascii")
            elif startup_script.exists():
                startup_script.unlink()
        except Exception:
            pass

    def _start_scheduler(self):
        if self._sched_thread and self._sched_thread.is_alive():
            return
        self._sched_stop_event = threading.Event()
        self._sched_last_reset_date = datetime.date.today()
        self._sched_thread = threading.Thread(target=self._scheduler_loop, daemon=True)
        self._sched_thread.start()
        self._sched_update_next()

    def _stop_scheduler(self):
        if hasattr(self, "_sched_stop_event"):
            self._sched_stop_event.set()
            # Join so the thread is actually gone before a quick re-enable can race
            # _start_scheduler's is_alive() guard and skip creating a new thread.
            if self._sched_thread:
                self._sched_thread.join(timeout=2)
        self._sched_update_next()

    def _backdate_check(self):
        times = sorted(self.settings.get("scheduled_times", []))
        if not times:
            return
        now = datetime.datetime.now()
        last_run_str = self.settings.get("last_scheduled_run", "")
        if last_run_str:
            try:
                last_run = datetime.datetime.fromisoformat(last_run_str)
            except Exception:
                last_run = None
        else:
            last_run = None

        missed = None
        for t in times:
            try:
                h, m = map(int, t.split(":"))
                sched_dt = now.replace(hour=h, minute=m, second=0, microsecond=0)
                if sched_dt > now:
                    continue
                if last_run and last_run >= sched_dt:
                    continue
                missed = t
            except Exception:
                continue

        if missed:
            label = self._format_sched_label(missed)
            self._log(self.sp_log,
                      f"[backdate] Missed trigger at {label} — running now")
            self._sched_fired_today.add(missed)
            self.settings["last_scheduled_run"] = now.isoformat()
            save_settings(self.settings)
            self._run_bot(self._scheduled_score_args(), self.sp_log, interactive=False)

    # ═══════════════════════════════════════════════════════════════════════════
    # Tab 2: Job Descriptions — online (URLs + SP folder) and local (pasted) JDs
    # ═══════════════════════════════════════════════════════════════════════════

    def _jd_save(self):
        from hiring_agent.jd_sources import save_jd_sources
        save_jd_sources(
            self.jd_enabled_var.get(), self._jd_urls, self._jd_descs,
            sharepoint_jd_folder=self.jd_sp_folder_var.get().strip(),
            sharepoint_jd_hostname=self.jd_sp_host_var.get().strip(),
            sharepoint_jd_site=self.jd_sp_site_var.get().strip(),
        )
        total = len(self._jd_urls) + len(self._jd_descs)
        state = "on" if self.jd_enabled_var.get() else "off"
        messagebox.showinfo("Saved", f"{total} job posting(s) saved (matching is {state}).\n\n"
                            "The next scoring run will use these.")

    def _jd_show_builtin(self):
        from hiring_agent.scoring import get_open_roles
        roles = get_open_roles()
        lines = []
        for r in roles:
            skills = ", ".join(r.get("skills", [])[:10])
            lines.append(f"{r['title']}:  {skills}")
        messagebox.showinfo("Built-in Default Roles", "\n".join(lines))

    # ═══════════════════════════════════════════════════════════════════════════
    # Tab 3: SharePoint — connection (.env), header check, local data folders
    # ═══════════════════════════════════════════════════════════════════════════

    def _sp_save_env(self):
        """Write the connection fields back to .env (preserving other keys)."""
        env_path = APP_DIR / ".env"
        existing: dict = {}
        try:
            if env_path.exists():
                for line in env_path.read_text(encoding="utf-8").splitlines():
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, _, v = line.partition("=")
                        existing[k.strip()] = v.strip()
        except Exception as e:
            # Don't overwrite an unreadable .env with a partial file — that would
            # silently drop every key this dialog doesn't know about.
            self.sp_conn_status.configure(
                text=f"Save aborted: could not read existing .env ({e})",
                text_color=ERROR)
            return
        for key, attr in self._sp_env_keys:
            val = getattr(self, attr).get().strip()
            if val:
                existing[key] = val
            elif key in existing and key == "SHAREPOINT_WORKBOOK":
                existing.pop(key)          # cleared optional key -> use the default
        try:
            env_path.write_text(
                "\n".join(f"{k}={v}" for k, v in existing.items()) + "\n",
                encoding="utf-8",
            )
            self.sp_conn_status.configure(
                text="Saved — restart DriverAI to use it.", text_color=WARNING)
            self._log(self.sp_tab_log, "[config] SharePoint settings saved to .env")
        except Exception as e:
            self.sp_conn_status.configure(text=f"Couldn't save: {e}", text_color=ERROR)

    # ═══════════════════════════════════════════════════════════════════════════
    # Tab 4: Resumes — local folder / existing workbook / single file
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_resumes_tab(self):
        self.nb.add("Resumes")
        tab = self.nb.tab("Resumes")

        ctk.CTkLabel(tab, text="Resumes that arrive by email are scored automatically "
                            "on the Home tab. Add resumes from your computer here and "
                            "they'll be scored the same way — nothing stays on this "
                            "computer.", fg_color="transparent", text_color=CARD_SUBTEXT,
                     wraplength=940, justify="left", anchor="w"
                     ).pack(anchor=tk.W, fill=tk.X, padx=10, pady=(10, 2))

        # ── Add local resumes to be scored ──
        up_card, up = self._card(tab, "Add Resumes to be Scored")
        up_card.pack(fill=tk.X, padx=10, pady=4)
        up_row = ctk.CTkFrame(up, fg_color="transparent")
        up_row.pack(fill=tk.X, pady=(6, 2))
        ctk.CTkButton(up_row, text="Choose Files...",
                      command=self._pick_upload_files).pack(side=tk.LEFT)
        ctk.CTkButton(up_row, text="Clear",
                      command=self._clear_upload_files).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(up_row, text="Send for Scoring",
                      command=self._upload_to_queue).pack(side=tk.LEFT, padx=6)
        self.upload_status = ctk.CTkLabel(up_row, text="", fg_color="transparent", text_color=ACCENT)
        self.upload_status.pack(side=tk.RIGHT, padx=8)
        self.upload_list = tk.Listbox(up, height=4, bg=CHART_SURFACE, fg=CARD_TEXT,
                                      selectbackground=CARD_ALT, relief=tk.FLAT,
                                      highlightthickness=0)
        self.upload_list.pack(fill=tk.X, pady=(2, 6))
        self._upload_files: list = []
        ctk.CTkLabel(up, text="They'll be scored the next time you click Score Now on the "
                           "Home tab.", fg_color="transparent", text_color=CARD_SUBTEXT,
                     font=self.fonts["caption"], anchor="w"
                     ).pack(anchor=tk.W, fill=tk.X, pady=(0, 6))

        # ── Quick preview: read a file in-memory, score it, write nothing ──
        one_card, one = self._card(tab, "Try a Resume First  (just to see how it's read — "
                                       "nothing is saved or sent anywhere)")
        one_card.pack(fill=tk.X, padx=10, pady=4)
        row4 = ctk.CTkFrame(one, fg_color="transparent")
        row4.pack(fill=tk.X, pady=(6, 6))
        ctk.CTkLabel(row4, text="Resume (PDF or Word):", fg_color="transparent",
                     text_color=CARD_TEXT).pack(side=tk.LEFT)
        self.file_var = tk.StringVar()
        ctk.CTkEntry(row4, textvariable=self.file_var).pack(side=tk.LEFT, fill=tk.X,
                                                         expand=True, padx=6)
        ctk.CTkButton(row4, text="Browse...", command=self._pick_file).pack(side=tk.LEFT)
        ctk.CTkButton(row4, text="Preview", command=self._score_one).pack(side=tk.LEFT, padx=6)

        self.resume_log = self._make_log(tab, height=14)

    def _pick_upload_files(self):
        files = filedialog.askopenfilenames(
            title="Choose resumes", filetypes=[("Resumes", "*.pdf *.docx")])
        for f in files:
            if f not in self._upload_files:
                self._upload_files.append(f)
                self.upload_list.insert(tk.END, f"  {Path(f).name}")
        if self._upload_files:
            self.upload_status.configure(
                text=f"{len(self._upload_files)} file(s) ready", text_color=ACCENT)

    def _clear_upload_files(self):
        self._upload_files.clear()
        self.upload_list.delete(0, tk.END)
        self.upload_status.configure(text="")

    def _upload_to_queue(self):
        files = list(self._upload_files)
        if not files:
            messagebox.showwarning("No files chosen", "Choose PDF or Word resumes first.")
            return
        if getattr(self, "_uploading", False):
            return
        self._uploading = True
        # Remove exactly this batch now (on the main thread) — anything the user adds
        # while the upload is in flight stays untouched instead of being wiped later.
        self._clear_upload_files()
        self.upload_status.configure(
            text=f"Sending {len(files)} file(s)...", text_color=WARNING)

        def worker():
            ok = err = 0
            try:
                import uuid
                from sharepoint_client import SharePointClient, SharePointError
                import hiring_agent.config as _cfg
                from hiring_agent.config import COLUMNS
                client = SharePointClient()
                client.ensure_workbook()
                for path in files:
                    now = datetime.datetime.now()
                    subpath = _cfg.dated_subpath(now)
                    ref = f"APP-{now:%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6].upper()}"
                    name = Path(path).name
                    stored_folder = f"{client.resumes_folder}/{subpath}"
                    stored_name = f"{ref}_{name}"
                    uploaded = False
                    try:
                        client.upload_file(stored_folder, stored_name,
                                           Path(path).read_bytes())
                        uploaded = True
                        row = {c: "" for c in COLUMNS}
                        row.update({
                            "Application ID": ref,
                            # ISO 'T' form on purpose: Excel Online keeps it as TEXT.
                            # The '%Y-%m-%d %H:%M:%S' space form gets coerced to a
                            # numeric date serial, which broke resume-path lookup and
                            # P1's ticks()/formatDateTime() reads of this cell.
                            "Received Date": now.strftime("%Y-%m-%dT%H:%M:%S"),
                            "Mail Subject": "Manual upload (GUI)",
                            "Status": "New Email Received",
                            "Has Resume": "Yes",
                            "Original Filename": name,
                            "Application Updates": 0,
                        })
                        from hiring_agent.sharepoint_scoring import (
                            _ensure_main_period_separator)
                        _ensure_main_period_separator(client, row["Received Date"])
                        client.add_main_row(row)
                        ok += 1
                        self.log_q.put((self.resume_log,
                                        f"[upload] {name} -> {ref} (queued)"))
                    except Exception as e:
                        err += 1
                        self.log_q.put((self.resume_log,
                                        f"[upload] FAILED {Path(path).name}: {e}"))
                        if uploaded:
                            # Row write failed after the file landed — remove the
                            # orphaned copy so it can't silently sit unscored forever.
                            try:
                                client.delete_file(stored_folder, stored_name)
                                self.log_q.put((self.resume_log,
                                                f"[upload]   cleaned up orphaned file"))
                            except SharePointError as ce:
                                self.log_q.put((self.resume_log,
                                                f"[upload]   WARNING: could not remove "
                                                f"the orphaned copy ({stored_name}): {ce} "
                                                f"— remove it manually on SharePoint."))
            except Exception as e:
                err += 1
                self.log_q.put((self.resume_log, f"[upload] ERROR: {e}"))

            def done():
                self._uploading = False
                self.upload_status.configure(
                    text=f"{ok} ready to score" + (f", {err} failed" if err else "")
                         + (" — go to the Home tab and click Score Now" if ok else ""),
                    text_color=SUCCESS if not err else ERROR)
                self._refresh_queue()
            self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def _pick_file(self):
        f = filedialog.askopenfilename(title="Choose a resume",
                                       filetypes=[("Resumes", "*.pdf *.docx")])
        if f:
            self.file_var.set(f)

    def _score_one(self):
        path = self.file_var.get().strip()
        if not path:
            messagebox.showwarning("Choose a file", "Choose a PDF or Word resume first.")
            return
        self._log(self.resume_log, f"Reading {Path(path).name} ...")

        def worker():
            try:
                from hiring_agent.local_scorer import score_resume_file
                r = score_resume_file(path)
                if r.get("error"):
                    self.log_q.put((self.resume_log, f"[error] {r['error']}"))
                    return
                c = r.get("candidate", {})
                self.log_q.put((self.resume_log, f"Name:     {c.get('full_name', '?')}"))
                self.log_q.put((self.resume_log, f"Location: {c.get('location', '?')}"))
                self.log_q.put((self.resume_log, f"Country:  {c.get('country', '?')}"))
                self.log_q.put((self.resume_log, f"Skills:   {c.get('skills', '?')}"))
                for m in r.get("role_matches", [])[:6]:
                    self.log_q.put((self.resume_log,
                                    f"  {m['score']:3d}%  {m['role_title']}"))
            except Exception as e:
                self.log_q.put((self.resume_log, f"[error] {e}"))

        threading.Thread(target=worker, daemon=True).start()

    # ═══════════════════════════════════════════════════════════════════════════
    # Tab 5: Candidates — browse, search, sort, filter, detail, export
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_candidates_tab(self):
        self.nb.add("Candidates")
        tab = self.nb.tab("Candidates")
        bar = ctk.CTkFrame(tab, fg_color="transparent"); bar.pack(fill=tk.X, padx=8, pady=6)
        ctk.CTkButton(bar, text="Refresh Candidates", command=self._load_candidates_online
                      ).pack(side=tk.LEFT)
        self.include_rejected_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(bar, text="Also show candidates outside the USA",
                        variable=self.include_rejected_var, onvalue=True, offvalue=False,
                        command=self._load_candidates_online).pack(side=tk.LEFT, padx=10)
        ctk.CTkButton(bar, text="Export...", command=self._export_candidates
                      ).pack(side=tk.RIGHT)
        ctk.CTkLabel(tab, text="Always shows the latest from SharePoint. "
                            "Double-click any candidate to see every stored detail.",
                     fg_color="transparent", text_color=CARD_SUBTEXT, font=self.fonts["caption"],
                     anchor="w").pack(anchor=tk.W, fill=tk.X, padx=8)

        self.candidate_rows = []
        self._filtered_rows = []
        self._sort_col = None
        self._sort_desc = False

        # ── Search + category filter ──
        filters_card, filters = self._card(tab, "Search & Filter")
        filters_card.pack(fill=tk.X, padx=8, pady=(0, 6))
        srow = ctk.CTkFrame(filters, fg_color="transparent"); srow.pack(fill=tk.X, pady=(6, 2))
        ctk.CTkLabel(srow, text="Search:", fg_color="transparent", text_color=CARD_TEXT
                     ).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_candidate_filter())
        ctk.CTkEntry(srow, textvariable=self.search_var, width=280).pack(side=tk.LEFT, padx=6)
        ctk.CTkLabel(srow, text="(name, email, skills, location, role)", fg_color="transparent",
                     text_color=CARD_SUBTEXT).pack(side=tk.LEFT)
        self.candidate_filter_status = ctk.CTkLabel(
            srow, text="Click Refresh Candidates to begin.", fg_color="transparent", text_color=CARD_TEXT)
        self.candidate_filter_status.pack(side=tk.RIGHT)

        filter_bar = ctk.CTkFrame(filters, fg_color="transparent")
        filter_bar.pack(fill=tk.X, pady=(0, 4))
        ctk.CTkLabel(filter_bar, text="Show categories:", fg_color="transparent", text_color=CARD_TEXT
                     ).pack(side=tk.LEFT)
        ctk.CTkButton(filter_bar, text="All", width=60,
                   command=self._select_all_candidate_categories).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(filter_bar, text="None", width=60,
                   command=self._clear_candidate_categories).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(filter_bar, text="What are these?",
                   command=self._show_candidate_categories).pack(side=tk.LEFT)

        self.category_listbox = tk.Listbox(filters, height=5, selectmode=tk.EXTENDED,
                                           exportselection=False, bg=CHART_SURFACE, fg=CARD_TEXT,
                                           selectbackground=ACCENT, selectforeground="#0b0d12",
                                           relief=tk.FLAT, highlightthickness=0)
        self.category_listbox.pack(fill=tk.X, pady=(0, 8))
        self.category_listbox.bind("<<ListboxSelect>>",
                                   lambda _e: self._apply_candidate_filter())

        cols = ("Full Name", "Location", "Country", "Category",
                "Suggested Role 1", "Suggested Role 2", "Suggested Role 3", "Status")
        headers = {"Suggested Role 1": "Best Match", "Suggested Role 2": "2nd Match",
                  "Suggested Role 3": "3rd Match"}
        self.tree = ttk.Treeview(tab, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c, text=headers.get(c, c), command=lambda col=c: self._sort_by(col))
            # 8 columns must fit a ~1000px window; wider defaults pushed the
            # Status column past the right edge with no way to reach it.
            self.tree.column(c, width=135, minwidth=90, stretch=True)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        self.tree.bind("<Double-1>", self._show_candidate_detail)

    def _fill_candidates(self, rows):
        self.tree.delete(*self.tree.get_children())
        self._filtered_rows = rows
        for i, r in enumerate(rows):
            self.tree.insert("", tk.END, iid=str(i), values=(
                r.get("Full Name", ""), r.get("Location", ""),
                r.get("Country", ""), r.get("Category", ""),
                r.get("Suggested Role 1", ""), r.get("Suggested Role 2", ""),
                r.get("Suggested Role 3", ""), r.get("Status", "")))

    def _sort_by(self, col):
        if self._sort_col == col:
            self._sort_desc = not self._sort_desc
        else:
            self._sort_col, self._sort_desc = col, False
        self._apply_candidate_filter()

    def _candidate_categories_defined(self) -> list[str]:
        try:
            from hiring_agent.config import ROLE_CATEGORY_RULES, ROLE_CATEGORY_DEFAULT
            cats = [str(rule.get("category", "")).strip() for rule in ROLE_CATEGORY_RULES]
            if ROLE_CATEGORY_DEFAULT:
                cats.append(str(ROLE_CATEGORY_DEFAULT).strip())
        except Exception:
            cats = []
        seen = set()
        out = []
        for cat in cats:
            if cat and cat not in seen:
                seen.add(cat)
                out.append(cat)
        return out

    def _loaded_candidate_categories(self) -> list[str]:
        seen = set()
        out = []
        for row in self.candidate_rows:
            cat = str(row.get("Category", "")).strip()
            if cat and cat not in seen:
                seen.add(cat)
                out.append(cat)
        return out

    def _refresh_candidate_categories(self):
        categories = self._loaded_candidate_categories()
        if not categories:
            categories = self._candidate_categories_defined()
        self.category_listbox.delete(0, tk.END)
        for cat in categories:
            self.category_listbox.insert(tk.END, cat)
        if categories:
            self.category_listbox.selection_set(0, tk.END)
            self.candidate_filter_status.configure(text=f"{len(categories)} categories loaded.")
        else:
            self.candidate_filter_status.configure(text="No categories found yet.")

    def _selected_candidate_categories(self) -> list[str]:
        return [self.category_listbox.get(i) for i in self.category_listbox.curselection()]

    def _select_all_candidate_categories(self):
        if self.category_listbox.size() > 0:
            self.category_listbox.selection_set(0, tk.END)
        self._apply_candidate_filter()

    def _clear_candidate_categories(self):
        self.category_listbox.selection_clear(0, tk.END)
        self._apply_candidate_filter()

    def _show_candidate_categories(self):
        loaded = self._loaded_candidate_categories()
        defined = self._candidate_categories_defined()
        cats = loaded or defined
        lines = ["These group candidates by the kind of role they best match:", ""]
        lines.extend(f"- {cat}" for cat in cats)
        messagebox.showinfo("What are categories?", "\n".join(lines))

    def _candidate_resume_url(self, row: dict) -> str:
        for col in ("Resume URL", "Resume Link"):
            val = str(row.get(col, "") or "").strip()
            if val.lower().startswith(("http://", "https://")):
                return val
        return ""

    def _apply_candidate_filter(self):
        selected = set(self._selected_candidate_categories())
        total = len(self.candidate_rows)
        if not self.candidate_rows:
            self._fill_candidates([])
            self.candidate_filter_status.configure(text="Click Refresh Candidates to begin.")
            return
        if not selected:
            rows = list(self.candidate_rows)
        else:
            rows = [r for r in self.candidate_rows
                    if str(r.get("Category", "")).strip() in selected]

        q = self.search_var.get().strip().lower()
        if q:
            def hits(r):
                hay = " ".join(str(r.get(k, "")) for k in (
                    "Full Name", "Email", "Current Skills", "Location", "Country",
                    "Looking For Role", "Suggested Role 1", "Suggested Role 2",
                    "Suggested Role 3", "Category")).lower()
                return q in hay
            rows = [r for r in rows if hits(r)]

        if self._sort_col:
            rows.sort(key=lambda r: str(r.get(self._sort_col, "")).lower(),
                      reverse=self._sort_desc)

        self._fill_candidates(rows)
        bits = [f"Showing {len(rows)} of {total} candidate(s)"]
        if selected and len(selected) != self.category_listbox.size():
            bits.append(f"{len(selected)} categories")
        if q:
            bits.append(f"search '{q}'")
        self.candidate_filter_status.configure(text="  |  ".join(bits) + ".")

    def _show_candidate_detail(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        try:
            row = self._filtered_rows[int(sel[0])]
        except (ValueError, IndexError):
            return
        try:
            from hiring_agent.config import COLUMNS
            cols = list(COLUMNS)
        except Exception:
            cols = list(row.keys())

        win = ctk.CTkToplevel(self.root)
        win.title(row.get("Full Name") or row.get("Application ID") or "Candidate")
        win.geometry("640x560")
        resume_url = self._candidate_resume_url(row)
        if resume_url:
            actions = ctk.CTkFrame(win, fg_color="transparent")
            actions.pack(fill=tk.X, padx=10, pady=(10, 0))
            ctk.CTkButton(actions, text="Open Resume",
                          command=lambda u=resume_url: webbrowser.open(u)
                          ).pack(side=tk.LEFT)
        inner = _ExemptingScrollableFrame(win, fg_color="transparent")
        inner.pack(fill=tk.BOTH, expand=True)

        for i, col in enumerate(cols):
            val = str(row.get(col, "") or "")
            ctk.CTkLabel(inner, text=f"{col}:", fg_color="transparent", text_color=CARD_TEXT,
                         font=self.fonts["small_bold"]).grid(
                row=i, column=0, sticky=tk.NW, padx=(10, 6), pady=2)
            target = val if val.lower().startswith("http") else (
                resume_url if col == "Resume Link" and resume_url else "")
            if target:
                link_text = val or (resume_url.rstrip("/").rsplit("/", 1)[-1].split("?")[0] if resume_url else "Resume")
                link = ctk.CTkLabel(inner, text=link_text, fg_color="transparent", text_color=ACCENT,
                                    cursor="hand2", wraplength=460, justify="left")
                link.grid(row=i, column=1, sticky=tk.W, pady=2)
                link.bind("<Button-1>", lambda _e, u=target: webbrowser.open(u))
            else:
                ctk.CTkLabel(inner, text=val or "-", fg_color="transparent", text_color=CARD_TEXT,
                             wraplength=460, justify="left").grid(
                    row=i, column=1, sticky=tk.W, pady=2)

    def _export_candidates(self):
        rows = self._filtered_rows
        if not rows:
            messagebox.showinfo("Nothing to export", "Click Refresh Candidates first.")
            return
        f = filedialog.asksaveasfilename(
            title="Save candidate list", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if not f:
            return
        try:
            import pandas as pd
            # Same contract as the automatic client export (the "result sheet"): the minimal
            # client column set, current-month-first order, a blank row under the header and
            # between calendar months, and a clickable Resume Link. Previously this button
            # dumped all 30 internal columns — Mail Body, Status, Retry Count and the rest —
            # with no separators and a dead Resume Link, so the file a user hand-exported
            # looked nothing like the one the client actually receives.
            from hiring_agent.sharepoint_scoring import (
                prepare_client_export_rows, write_client_export, _CLIENT_EXPORT_COLUMNS)
            prepared = prepare_client_export_rows(rows)
            if f.lower().endswith(".csv"):
                pd.DataFrame(prepared, columns=_CLIENT_EXPORT_COLUMNS).to_csv(f, index=False)
            else:
                write_client_export(prepared, f)
            messagebox.showinfo("Saved", f"{len(rows)} candidate(s) saved to:\n{f}")
        except Exception as e:
            messagebox.showerror("Couldn't save", str(e))

    def _load_candidates_online(self):
        """Load candidates straight from the live SharePoint table."""
        if getattr(self, "_cand_loading", False):
            return
        self._cand_loading = True
        self.candidate_filter_status.configure(text="Loading from SharePoint...")

        def worker():
            try:
                from sharepoint_client import SharePointClient
                client = SharePointClient()
                rows = [r["values"] for r in client.list_rows()]
                if self.include_rejected_var.get():
                    rows += [r["values"] for r in client.list_rejected_rows()]

                def done():
                    self._cand_loading = False
                    self.candidate_rows = rows
                    self._refresh_candidate_categories()
                    self._apply_candidate_filter()
                self.root.after(0, done)
            except Exception as e:
                msg = str(e)

                def fail():
                    self._cand_loading = False
                    self.candidate_filter_status.configure(
                        text="Couldn't connect to SharePoint.")
                    messagebox.showerror(
                        "Couldn't load candidates",
                        f"Check the connection on the SharePoint tab.\n\nDetails: {msg}")
                self.root.after(0, fail)

        threading.Thread(target=worker, daemon=True).start()

    # ═══════════════════════════════════════════════════════════════════════════
    # Tab 6: Settings
    # ═══════════════════════════════════════════════════════════════════════════
    def _build_settings_tab(self):
        self.nb.add("Settings")
        tab = self.nb.tab("Settings")
        form = ctk.CTkFrame(tab, fg_color="transparent")
        form.pack(fill=tk.X, padx=12, pady=12)

        ctk.CTkLabel(form, text="AI model", fg_color="transparent", text_color=CARD_TEXT
                     ).grid(row=0, column=0, sticky=tk.W, pady=4)
        self.model_var = tk.StringVar(value=self.settings["ollama_model"])
        self.model_combo = ctk.CTkComboBox(form, variable=self.model_var, width=260,
                                           values=[self.settings["ollama_model"]])
        self.model_combo.grid(row=0, column=1, pady=4, sticky=tk.W)
        ctk.CTkButton(form, text="Refresh List", command=self._refresh_models
                      ).grid(row=0, column=2, padx=6)
        ctk.CTkLabel(form, text="The AI that reads and scores resumes.", fg_color="transparent",
                     text_color=CARD_SUBTEXT, font=self.fonts["caption"]
                     ).grid(row=1, column=1, sticky=tk.W)

        ctk.CTkLabel(form, text="AI server address", fg_color="transparent", text_color=CARD_TEXT
                     ).grid(row=2, column=0, sticky=tk.W, pady=(10, 4))
        self.host_var = tk.StringVar(value=self.settings["ollama_host"])
        ctk.CTkEntry(form, textvariable=self.host_var, width=280).grid(
            row=2, column=1, pady=(10, 4), sticky=tk.W)
        ctk.CTkLabel(form, text="Leave this as-is unless your IT admin tells you to change it.",
                     fg_color="transparent", text_color=CARD_SUBTEXT, font=self.fonts["caption"]
                     ).grid(row=3, column=1, sticky=tk.W)

        self.ai_var = tk.BooleanVar(value=self.settings["use_ai"])
        ctk.CTkCheckBox(form, text="Use AI to score candidates (off = simple keyword matching)",
                        variable=self.ai_var, onvalue=True, offvalue=False
                        ).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(14, 4))
        self.geo_var = tk.BooleanVar(value=self.settings["usa_only"])
        ctk.CTkCheckBox(form, text="Only consider candidates based in the USA",
                        variable=self.geo_var, onvalue=True, offvalue=False
                        ).grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=4)

        ctk.CTkButton(form, text="Save", command=self._save
                      ).grid(row=6, column=0, sticky=tk.W, pady=10)
        ctk.CTkButton(form, text="Test AI Connection", command=self._test_ollama
                      ).grid(row=6, column=1, sticky=tk.W, pady=10)
        self.settings_log = self._make_log(tab)

    def _refresh_models(self):
        host = self.host_var.get().strip() or "http://localhost:11434"
        self._log(self.settings_log, f"Checking {host} for AI models...")

        def worker():
            try:
                import requests
                r = requests.get(f"{host.rstrip('/')}/api/tags", timeout=5)
                r.raise_for_status()
                models = [m.get("name", "") for m in r.json().get("models", []) if m.get("name")]
                if models:
                    self.root.after(0, lambda: self.model_combo.configure(values=models))
                    self.log_q.put((self.settings_log,
                                    f"Found {len(models)} model(s): {', '.join(models)}"))
                else:
                    self.log_q.put((self.settings_log,
                                    "The AI service is running, but no models are "
                                    "installed yet. Ask your IT admin to run: "
                                    "ollama pull llama3.2"))
            except Exception as e:
                self.log_q.put((self.settings_log,
                                f"Can't reach the AI service ({e}). Make sure it's "
                                f"running, then try again."))
        threading.Thread(target=worker, daemon=True).start()

    def _save(self):
        try:
            mins = max(1, int(self.sp_interval.get()))
        except ValueError:
            mins = self.settings.get("trigger_interval_min", 5)
        self.settings.update({
            "ollama_model": self.model_var.get().strip() or "llama3.2",
            "ollama_host": self.host_var.get().strip() or "http://localhost:11434",
            "use_ai": bool(self.ai_var.get()),
            "usa_only": bool(self.geo_var.get()),
            "auto_start": bool(self.auto_start_var.get()),
            "trigger_interval_min": mins,
            "schedule_enabled": bool(self.sched_enabled_var.get()),
            "backdate_enabled": bool(self.backdate_var.get()),
            "launch_on_startup": bool(self.launch_startup_var.get()),
        })
        save_settings(self.settings)
        self._sync_startup_shortcut()
        self._log(self.settings_log, "Saved.")

    def _test_ollama(self):
        host = self.host_var.get().strip() or "http://localhost:11434"
        self._log(self.settings_log, f"Checking {host} ...")

        def worker():
            try:
                import requests
                r = requests.get(f"{host.rstrip('/')}/api/tags", timeout=5)
                r.raise_for_status()
                models = [m.get("name", "?") for m in r.json().get("models", [])]
                self.log_q.put((self.settings_log,
                                f"Connected. Models available: "
                                f"{', '.join(models) or '(none installed yet)'}"))
            except Exception as e:
                self.log_q.put((self.settings_log,
                                f"Can't reach the AI service ({e}). DriverAI will use "
                                f"simple keyword matching instead."))

        threading.Thread(target=worker, daemon=True).start()

    # ═══════════════════════════════════════════════════════════════════════════
    # Status bar (bottom of window) + startup health check
    # ═══════════════════════════════════════════════════════════════════════════
    def _refresh_recent_activity(self):
        if not hasattr(self, "activity_list"):
            return
        self.activity_list.delete(0, tk.END)
        hist = [h for h in self._history if not h.get("dry")][-8:]
        if not hist:
            self.activity_list.insert(tk.END, "No completed live runs yet.")
            return
        for item in reversed(hist):
            ts = str(item.get("ts", ""))[5:16].replace("T", " ")
            if item.get("rc") == 0:
                secs = item.get("secs", 0) or 0
                summary = (f"[OK] {item.get('processed', 0)} scored / "
                           f"{item.get('rejected', 0)} rejected in {secs:.0f}s")
            else:
                summary = f"[FAIL] code {item.get('rc')} after {item.get('secs', 0):.0f}s"
            self.activity_list.insert(tk.END, f"{ts}  {summary}")

    def _refresh_queue(self):
        def worker():
            text = "Not connected"
            color = CARD_SUBTEXT
            big = "Not connected"
            try:
                from hiring_agent.config import SHAREPOINT_CONFIGURED
                if not SHAREPOINT_CONFIGURED:
                    raise RuntimeError("sharepoint not configured")
                from sharepoint_client import SharePointClient
                from hiring_agent.config import STATUS_NEEDS_REVIEW
                n = len(SharePointClient().list_unscored_rows(
                    ("New Email Received", STATUS_NEEDS_REVIEW)))
                if n:
                    text = f"{n} waiting"
                    color = WARNING
                    big = str(n)
                else:
                    text = "All caught up"
                    color = SUCCESS
                    big = "0"
            except Exception as e:
                if "sharepoint not configured" not in str(e).lower():
                    text = "Connection issue"
                    color = ERROR
                    big = "Unavailable"

            def done():
                if hasattr(self, "queue_lbl"):
                    self.queue_lbl.configure(text=text, text_color=color)
                if hasattr(self, "queue_big_var"):
                    self.queue_big_var.set(big)
            self.root.after(0, done)
        threading.Thread(target=worker, daemon=True).start()

    def _refresh_insights(self):
        def worker():
            data = {
                "queue": None,
                "rows": 0,
                "rejected": 0,
                "categories": [],
                "roles": [],
                "recent_candidates": [],
                "sharepoint_error": "",
                "last_refresh": datetime.datetime.now().strftime("%I:%M %p").lstrip("0"),
            }
            try:
                from hiring_agent.config import SHAREPOINT_CONFIGURED
                if not SHAREPOINT_CONFIGURED:
                    raise RuntimeError("SharePoint not configured")
                from sharepoint_client import SharePointClient
                client = SharePointClient()
                rows = [r["values"] for r in client.list_rows()]
                rejected_rows = [r["values"] for r in client.list_rejected_rows()]
                data["queue"] = sum(1 for r in rows if str(r.get("Status", "")).strip() == "New Email Received")
                data["rows"] = len(rows)
                data["rejected"] = len(rejected_rows)
                cats = Counter()
                roles = Counter()
                recent = []
                for row in rows:
                    cat = str(row.get("Category", "")).strip()
                    role = str(row.get("Suggested Role 1", "")).strip()
                    if cat:
                        cats[cat] += 1
                    if role:
                        roles[role] += 1
                    recent.append((
                        str(row.get("Received Date", "")).strip(),
                        str(row.get("Full Name", "")).strip() or str(row.get("Email", "")).strip(),
                        str(row.get("Status", "")).strip(),
                    ))
                data["categories"] = cats.most_common(6)
                data["roles"] = roles.most_common(6)
                recent.sort(reverse=True)
                data["recent_candidates"] = recent[:6]
                meta = client.workbook_metadata()
                self.root.after(0, lambda m=meta: self._apply_sharepoint_metadata(m))
            except Exception as e:
                data["sharepoint_error"] = str(e)

            def done():
                self._insights = data
                self._render_insights()
            self.root.after(0, done)
        threading.Thread(target=worker, daemon=True).start()

    def _render_insights(self):
        if hasattr(self, "category_list"):
            self.category_list.delete(0, tk.END)
            items = self._insights.get("categories") or []
            if not items:
                self.category_list.insert(tk.END, "No category data yet.")
                if self._insights.get("sharepoint_error"):
                    self.category_list.insert(tk.END, "Connect SharePoint and refresh insights.")
            else:
                total = max(1, sum(count for _, count in items))
                for name, count in items:
                    pct = round((count / total) * 100)
                    self.category_list.insert(tk.END, f"{count:>3}  {pct:>2}%  {name}")
        if hasattr(self, "role_list"):
            self.role_list.delete(0, tk.END)
            items = self._insights.get("roles") or []
            if not items:
                self.role_list.insert(tk.END, "No role data yet.")
                if self._insights.get("sharepoint_error"):
                    self.role_list.insert(tk.END, "Connect SharePoint and refresh insights.")
            else:
                total = max(1, sum(count for _, count in items))
                for name, count in items:
                    pct = round((count / total) * 100)
                    self.role_list.insert(tk.END, f"{count:>3}  {pct:>2}%  {name}")
        if hasattr(self, "snapshot_list"):
            self.snapshot_list.delete(0, tk.END)
            if self._insights.get("sharepoint_error"):
                self.snapshot_list.insert(tk.END, "SharePoint unavailable")
                self.snapshot_list.insert(tk.END, self._insights["sharepoint_error"][:120])
            else:
                self.snapshot_list.insert(tk.END, f"Queue waiting: {self._insights.get('queue', 0)}")
                self.snapshot_list.insert(tk.END, f"Workbook rows: {self._insights.get('rows', 0)}")
                self.snapshot_list.insert(tk.END, f"Rejected sheet: {self._insights.get('rejected', 0)}")
                self.snapshot_list.insert(tk.END, f"Insights refreshed: {self._insights.get('last_refresh', '-')}")
                for received, name, status in self._insights.get("recent_candidates", [])[:2]:
                    label = (name or "Unnamed")[:40]
                    self.snapshot_list.insert(tk.END, f"{received or '-'}  {label}  {status or '-'}")

    def _refresh_metrics(self):
        hist = [h for h in self._history if not h.get("dry")]
        today = datetime.date.today().isoformat()
        today_runs = [h for h in hist if str(h.get("ts", "")).startswith(today)]
        ok = sum(1 for h in hist if h.get("rc") == 0)
        failed = max(0, len(hist) - ok)
        timed = [h for h in hist[-20:] if h.get("secs")]
        avg = sum(h["secs"] for h in timed) / len(timed) if timed else 0
        scored = sum(h.get("processed", 0) or 0 for h in hist)
        rejected = sum(h.get("rejected", 0) or 0 for h in hist)
        last = hist[-1] if hist else None
        if hasattr(self, "_tile_vars"):
            self._tile_vars["today"].set(str(len(today_runs)))
            self._tile_vars["avg"].set(f"{avg:.0f}s" if avg else "-")
            self._tile_vars["cands"].set(f"{scored} / {rejected}" if hist else "-")
            if "okfail" in self._tile_vars:
                self._tile_vars["okfail"].set(f"{ok} / {failed}")
            if last:
                state = "Worked" if last.get("rc") == 0 else "Failed"
                self._tile_vars["last"].set(f"{state} in {last.get('secs', 0):.0f}s")
            else:
                self._tile_vars["last"].set("-")
        self._draw_area_chart()
        self._draw_duration_chart()
        self._refresh_recent_activity()

    def _save_trigger_settings(self, interactive: bool = False):
        try:
            mins = max(1, int(self.sp_interval.get()))
        except ValueError:
            mins = 5
        batch_size = self._selected_batch_size(interactive=interactive)
        if batch_size is None:
            return False
        self.settings["trigger_interval_min"] = mins
        self.settings["batch_size"] = batch_size
        self.settings["auto_start"] = bool(self.auto_start_var.get())
        self.settings["schedule_enabled"] = bool(self.sched_enabled_var.get())
        self.settings["backdate_enabled"] = bool(self.backdate_var.get())
        self.settings["launch_on_startup"] = bool(self.launch_startup_var.get())
        save_settings(self.settings)
        self._sync_startup_shortcut()
        self._sched_refresh_list()
        return True

    def _sync_scheduler_state(self):
        if self.settings.get("schedule_enabled") and self.settings.get("scheduled_times"):
            self._start_scheduler()
        else:
            self._stop_scheduler()

    def _sched_refresh_list(self):
        times = sorted(self.settings.get("scheduled_times", []))
        if hasattr(self, "sched_listbox") and self.sched_listbox.winfo_exists():
            self.sched_listbox.delete(0, tk.END)
            for t in times:
                self.sched_listbox.insert(tk.END, f"  {self._format_sched_label(t)}   ({t})")
        labels = [self._format_sched_label(t) for t in times]
        if hasattr(self, "sched_summary_var"):
            if self.settings.get("schedule_enabled") and labels:
                summary = "On - " + ", ".join(labels[:3])
                if len(labels) > 3:
                    summary += " ..."
                self.sched_summary_var.set(summary)
            elif labels:
                self.sched_summary_var.set("Off - times saved")
            else:
                self.sched_summary_var.set("Off")
        self._sched_update_next()

    def _sched_update_next(self):
        times = sorted(self.settings.get("scheduled_times", []))
        if not self.settings.get("schedule_enabled") or not times:
            next_text = "No next run"
        else:
            now = datetime.datetime.now()
            next_time = None
            for t in times:
                try:
                    h, m = map(int, t.split(":"))
                    sched_dt = now.replace(hour=h, minute=m, second=0, microsecond=0)
                    if sched_dt >= now:
                        next_time = sched_dt
                        break
                except Exception:
                    continue
            if next_time is None:
                h, m = map(int, times[0].split(":"))
                next_time = (now + datetime.timedelta(days=1)).replace(
                    hour=h, minute=m, second=0, microsecond=0)
            next_text = f"Next run: {next_time.strftime('%b %d, %I:%M %p').lstrip('0')}"
        if hasattr(self, "next_run_var"):
            self.next_run_var.set(next_text)
            # Color reflects actual state - "No next run" is neutral, not a success to celebrate.
            color = SUCCESS if next_text.startswith("Next run:") else CARD_SUBTEXT
            for attr in ("next_run_lbl", "next_run_lbl_popup"):
                lbl = getattr(self, attr, None)
                if lbl is not None and lbl.winfo_exists():
                    lbl.configure(text_color=color)
        last_run_str = self.settings.get("last_scheduled_run", "")
        if hasattr(self, "sched_last_var"):
            if last_run_str:
                try:
                    when = datetime.datetime.fromisoformat(last_run_str)
                    self.sched_last_var.set("Last scheduled run: " + when.strftime("%b %d, %I:%M %p").lstrip("0"))
                except Exception:
                    self.sched_last_var.set("Last scheduled run: recorded")
            else:
                self.sched_last_var.set("No scheduled run yet")

    def _scheduler_loop(self):
        while not self._sched_stop_event.is_set():
            if not self.settings.get("schedule_enabled"):
                break
            now = datetime.datetime.now()
            hhmm = now.strftime("%H:%M")
            times = self.settings.get("scheduled_times", [])
            if hhmm in times and hhmm not in self._sched_fired_today:
                self._sched_fired_today.add(hhmm)
                self.settings["last_scheduled_run"] = now.isoformat()
                save_settings(self.settings)
                label = self._format_sched_label(hhmm)
                self.log_q.put((self.sp_log, f"\n[schedule] Triggered at {label} ({hhmm})"))
                scheduled_args = self._scheduled_score_args()
                self.root.after(0, lambda args=scheduled_args: self._run_bot(
                    args, self.sp_log, interactive=False))
                self.root.after(0, self._sched_refresh_list)
            today = now.date()
            if today != self._sched_last_reset_date:
                self._sched_fired_today.clear()
                self._sched_last_reset_date = today
            self._sched_stop_event.wait(30)

    def _sched_prefill_selected(self, *_args):
        if not hasattr(self, "sched_listbox"):
            return
        sel = self.sched_listbox.curselection()
        if not sel:
            return
        times = sorted(self.settings.get("scheduled_times", []))
        idx = sel[0]
        if idx >= len(times):
            return
        hhmm = times[idx]
        self.sched_time_var.set(hhmm)
        try:
            hour, minute = hhmm.split(":")
            self.sched_hour_var.set(hour)
            self.sched_min_var.set(minute)
        except ValueError:
            pass

    def _sched_edit_time(self):
        sel = self.sched_listbox.curselection() if hasattr(self, "sched_listbox") else ()
        if not sel:
            messagebox.showinfo("Select", "Select a time to edit.")
            return
        typed_time = self.sched_time_var.get().strip()
        hhmm = self._normalize_schedule_time(typed_time) if typed_time else None
        if not hhmm:
            try:
                hhmm = f"{int(self.sched_hour_var.get()) % 24:02d}:{int(self.sched_min_var.get()) % 60:02d}"
            except ValueError:
                messagebox.showwarning("Invalid time", "Enter a valid replacement time.")
                return
        times = sorted(self.settings.get("scheduled_times", []))
        idx = sel[0]
        if idx < len(times):
            times[idx] = hhmm
            self.settings["scheduled_times"] = sorted(set(times))
            save_settings(self.settings)
            self._save_trigger_settings()
            self._sync_scheduler_state()
            self._sched_refresh_list()

    def _sched_add_time(self):
        typed_time = self.sched_time_var.get().strip()
        hhmm = self._normalize_schedule_time(typed_time) if typed_time else None
        if not hhmm:
            try:
                hhmm = f"{int(self.sched_hour_var.get()) % 24:02d}:{int(self.sched_min_var.get()) % 60:02d}"
            except ValueError:
                messagebox.showwarning("Invalid time", "Enter a valid time.")
                return
        times = self.settings.get("scheduled_times", [])
        if hhmm in times:
            messagebox.showinfo("Duplicate", f"{hhmm} is already scheduled.")
            return
        times.append(hhmm)
        self.settings["scheduled_times"] = sorted(times)
        save_settings(self.settings)
        self._save_trigger_settings()
        self._sync_scheduler_state()
        self.sched_time_var.set("")
        self._sched_refresh_list()

    def _sched_remove_time(self):
        sel = self.sched_listbox.curselection() if hasattr(self, "sched_listbox") else ()
        if not sel:
            messagebox.showinfo("Select", "Select a time to remove.")
            return
        times = sorted(self.settings.get("scheduled_times", []))
        for idx in reversed(sel):
            if idx < len(times):
                times.pop(idx)
        self.settings["scheduled_times"] = times
        save_settings(self.settings)
        self._save_trigger_settings()
        self._sync_scheduler_state()
        self._sched_refresh_list()

    def _on_schedule_toggle(self):
        if not self._save_trigger_settings(interactive=True):
            self.sched_enabled_var.set(False)
            return
        self._sync_scheduler_state()
        self._sched_refresh_list()

    def _open_scheduler_window(self):
        if getattr(self, "_sched_window", None) and self._sched_window.winfo_exists():
            self._sched_window.lift()
            return
        win = ctk.CTkToplevel(self.root)
        win.title("Scheduler")
        win.geometry("620x600")
        win.minsize(620, 600)
        self._sched_window = win
        outer = ctk.CTkFrame(win, fg_color="transparent")
        outer.pack(fill=tk.BOTH, expand=True, padx=14, pady=14)
        ctk.CTkCheckBox(outer, text="Turn on daily schedule", variable=self.sched_enabled_var,
                        onvalue=True, offvalue=False,
                        command=self._on_schedule_toggle).pack(anchor=tk.W)
        ctk.CTkLabel(outer, textvariable=self.sched_summary_var, fg_color="transparent", text_color=CARD_SUBTEXT,
                     wraplength=500, justify="left", anchor="w").pack(anchor=tk.W, fill=tk.X, pady=(4, 0))
        self.next_run_lbl_popup = ctk.CTkLabel(outer, textvariable=self.next_run_var, fg_color="transparent",
                     text_color=CARD_SUBTEXT, anchor="w")
        self.next_run_lbl_popup.pack(anchor=tk.W, fill=tk.X, pady=(4, 0))
        ctk.CTkLabel(outer, textvariable=self.sched_last_var, fg_color="transparent", text_color=CARD_SUBTEXT,
                     anchor="w").pack(anchor=tk.W, fill=tk.X)
        batch_row = ctk.CTkFrame(outer, fg_color="transparent")
        batch_row.pack(fill=tk.X, pady=(14, 0))
        ctk.CTkLabel(batch_row, text="Candidates per run", fg_color="transparent",
                     text_color=CARD_TEXT).pack(side=tk.LEFT)
        schedule_batch = ctk.CTkComboBox(
            batch_row,
            variable=self.batch_size_var,
            values=list(_BATCH_SIZE_CHOICES),
            width=100,
            command=lambda _value: self._save_trigger_settings(),
        )
        schedule_batch.pack(side=tk.LEFT, padx=(10, 0))
        schedule_batch.bind(
            "<FocusOut>", lambda _event: self._save_trigger_settings(), add="+"
        )
        schedule_batch.bind(
            "<Return>", lambda _event: self._save_trigger_settings(interactive=True),
            add="+",
        )
        ctk.CTkLabel(
            outer,
            text="Type 9am or use the pickers. Selecting a saved time loads it back into the editor.",
            fg_color="transparent", text_color=CARD_SUBTEXT, wraplength=500, justify="left", anchor="w"
        ).pack(anchor=tk.W, fill=tk.X, pady=(10, 0))
        row = ctk.CTkFrame(outer, fg_color="transparent"); row.pack(fill=tk.X, pady=(16, 2))
        ctk.CTkLabel(row, text="Quick time", fg_color="transparent", text_color=CARD_TEXT
                     ).pack(side=tk.LEFT)
        ctk.CTkEntry(row, textvariable=self.sched_time_var, width=100).pack(side=tk.LEFT, padx=(6, 10))
        ctk.CTkLabel(row, text="Hour", fg_color="transparent", text_color=CARD_TEXT).pack(side=tk.LEFT)
        ttk.Spinbox(row, from_=0, to=23, width=3, textvariable=self.sched_hour_var,
                    format="%02.0f", wrap=True).pack(side=tk.LEFT, padx=(6, 0))
        ctk.CTkLabel(row, text=":", fg_color="transparent", text_color=CARD_TEXT).pack(side=tk.LEFT)
        ttk.Spinbox(row, from_=0, to=59, width=3, textvariable=self.sched_min_var,
                    format="%02.0f", wrap=True).pack(side=tk.LEFT)
        btn_row = ctk.CTkFrame(outer, fg_color="transparent"); btn_row.pack(fill=tk.X, pady=(6, 6))
        ctk.CTkButton(btn_row, text="Add", command=self._sched_add_time).pack(side=tk.LEFT)
        ctk.CTkButton(btn_row, text="Edit Selected", command=self._sched_edit_time).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(btn_row, text="Remove", command=self._sched_remove_time).pack(side=tk.LEFT)
        self.sched_listbox = tk.Listbox(
            outer, height=8, selectmode=tk.EXTENDED, bg=CHART_SURFACE, fg=CARD_TEXT,
            selectbackground=CARD_ALT, relief=tk.FLAT, highlightthickness=0
        )
        self.sched_listbox.pack(fill=tk.X, pady=(6, 12))
        self.sched_listbox.bind("<<ListboxSelect>>", self._sched_prefill_selected)
        if not CLIENT_MODE:
            ctk.CTkCheckBox(outer, text="Score automatically when I open the app",
                            variable=self.auto_start_var, onvalue=True, offvalue=False,
                            command=self._save_trigger_settings).pack(anchor=tk.W)
            ctk.CTkCheckBox(outer, text="Catch up on any times I missed",
                            variable=self.backdate_var, onvalue=True, offvalue=False,
                            command=self._save_trigger_settings).pack(anchor=tk.W, pady=(4, 0))
        ctk.CTkCheckBox(outer, text="Open DriverAI when my computer starts",
                        variable=self.launch_startup_var, onvalue=True, offvalue=False,
                        command=self._save_trigger_settings).pack(anchor=tk.W, pady=(4, 0))

        def close_scheduler():
            if self._save_trigger_settings(interactive=True):
                self._sync_scheduler_state()
                win.destroy()

        ctk.CTkButton(outer, text="Close", command=close_scheduler).pack(
            anchor=tk.E, pady=(14, 0)
        )
        self._sched_refresh_list()

    def _on_run_complete(self, payload):
        self._bot_running = False
        self._set_busy(False)
        if not isinstance(payload, dict):
            payload = {"rc": payload, "secs": 0.0, "mode": "run", "dry": False}
        rc = payload.get("rc", -1)
        now = datetime.datetime.now().strftime("%I:%M %p").lstrip("0")
        parts = []
        if "processed" in payload:
            parts.append(f"{payload['processed']} scored")
            parts.append(f"{payload.get('rejected', 0)} rejected")
            if payload.get("errors"):
                parts.append(f"{payload['errors']} errors")
        detail = f" - {', '.join(parts)}" if parts else ""
        if rc == 0:
            self.sp_last_run.configure(text=f"Last: {now}{detail}", text_color=SUCCESS)
        else:
            self.sp_last_run.configure(text=f"Last: {now} FAILED (code {rc}) - see log", text_color=ERROR)
            if payload.get("mode") == "score-sharepoint" and hasattr(
                    self, "run_progress_title_var"):
                self.run_progress_title_var.set("Run stopped or failed")
                self.run_progress_left_var.set("Check the activity log for the last completed step")
        if not payload.get("dry"):
            entry = {"ts": datetime.datetime.now().isoformat(timespec="seconds"), **payload}
            append_history(entry)
            self._history.append(entry)
        self._refresh_metrics()
        self.root.after(400, self._run_health_check)
        self.root.after(600, self._refresh_insights)
        if payload.get("mode") == "score-sharepoint":
            self.root.after(800, self._refresh_queue)

    def _build_home_tab(self):
        tab = self._scroll_tab("Home")
        self.sched_enabled_var = tk.BooleanVar(value=self.settings.get("schedule_enabled", False))
        self.auto_start_var = tk.BooleanVar(value=self.settings.get("auto_start", False))
        self.backdate_var = tk.BooleanVar(value=self.settings.get("backdate_enabled", True))
        self.launch_startup_var = tk.BooleanVar(value=self.settings.get("launch_on_startup", False))
        self.sp_interval = tk.StringVar(value=str(self.settings.get("trigger_interval_min", 5)))
        self.batch_size_var = tk.StringVar(value=str(self.settings.get("batch_size", 25)))
        self.sched_time_var = tk.StringVar()
        self.sched_hour_var = tk.StringVar(value="09")
        self.sched_min_var = tk.StringVar(value="00")
        shell = ctk.CTkFrame(tab, fg_color="transparent")
        shell.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        top = ctk.CTkFrame(shell, fg_color="transparent")
        top.pack(fill=tk.X)
        left = ctk.CTkFrame(top, fg_color="transparent")
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        right = ctk.CTkFrame(top, fg_color="transparent")
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        hero_outer, hero = self._card(left, "P2 Control",
                                      "",
                                      minheight=252)
        hero_outer.pack(fill=tk.X, pady=(0, 8))
        actions = ctk.CTkFrame(hero, fg_color="transparent"); actions.pack(fill=tk.X, pady=(8, 8))
        self.run_btn = ctk.CTkButton(
            actions, text="Run Live", width=120, command=lambda: self._sp_run(False)
        )
        self.run_btn.pack(side=tk.LEFT)
        ctk.CTkButton(actions, text="Stop", width=90, command=self._sp_stop).pack(side=tk.LEFT, padx=6)
        # Rebuilding the client sheet used to require a whole scoring run (the export only
        # ran as its final step), so a mid-day refresh for the client meant either waiting
        # or running a scoring pass you did not otherwise want.
        ctk.CTkButton(actions, text="Export Results", width=130,
                      command=self._sp_export_results).pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkLabel(actions, text="Batch", fg_color="transparent",
                     text_color=CARD_SUBTEXT).pack(side=tk.LEFT, padx=(14, 6))
        self.batch_size_combo = ctk.CTkComboBox(
            actions,
            variable=self.batch_size_var,
            values=list(_BATCH_SIZE_CHOICES),
            width=90,
            command=lambda _value: self._selected_batch_size(interactive=False),
        )
        self.batch_size_combo.pack(side=tk.LEFT)
        self.batch_size_combo.bind(
            "<Return>",
            lambda _event: self._selected_batch_size(interactive=True),
            add="+",
        )
        self.run_spinner = ttk.Progressbar(actions, mode="indeterminate", length=90)
        self.run_spinner.pack(side=tk.RIGHT)
        status = ctk.CTkFrame(hero, fg_color="transparent"); status.pack(fill=tk.X)
        self.sp_status = ctk.CTkLabel(status, text="Ready", fg_color="transparent", text_color=CARD_SUBTEXT)
        self.sp_status.pack(side=tk.LEFT)
        self.sp_last_run = ctk.CTkLabel(status, text="No runs yet", fg_color="transparent", text_color=CARD_SUBTEXT)
        self.sp_last_run.pack(side=tk.RIGHT)
        self.run_progress_title_var = tk.StringVar(value="Ready to run")
        self.run_progress_detail_var = tk.StringVar(
            value="Current candidate and row will appear here"
        )
        self.run_progress_left_var = tk.StringVar(value="Queue progress is shown during a run")
        ctk.CTkLabel(
            hero, textvariable=self.run_progress_title_var, fg_color="transparent",
            text_color=CARD_TEXT, font=self.fonts["heading"], anchor="w",
            wraplength=470, justify=tk.LEFT,
        ).pack(anchor=tk.W, fill=tk.X, pady=(14, 2))
        ctk.CTkLabel(
            hero, textvariable=self.run_progress_detail_var, fg_color="transparent",
            text_color=CARD_SUBTEXT, anchor="w", wraplength=470, justify=tk.LEFT,
        ).pack(anchor=tk.W, fill=tk.X)
        self.run_progress_bar = ctk.CTkProgressBar(
            hero, height=10, progress_color=ACCENT, fg_color=CARD_EDGE,
        )
        self.run_progress_bar.pack(fill=tk.X, pady=(8, 4))
        self.run_progress_bar.set(0)
        ctk.CTkLabel(
            hero, textvariable=self.run_progress_left_var, fg_color="transparent",
            text_color=CARD_SUBTEXT, font=self.fonts["caption"], anchor="w",
            wraplength=470, justify=tk.LEFT,
        ).pack(anchor=tk.W, fill=tk.X)
        cards = ctk.CTkFrame(left, fg_color="transparent"); cards.pack(fill=tk.X)
        queue_outer, queue_card = self._card(cards, "Queue", minheight=118)
        queue_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        self.queue_big_var = tk.StringVar(value="Checking...")
        ctk.CTkLabel(queue_card, textvariable=self.queue_big_var, fg_color="transparent", text_color=ACCENT,
                     font=self.fonts["display"], anchor="w").pack(anchor=tk.W, fill=tk.X, pady=(10, 2))
        self.queue_lbl = ctk.CTkLabel(queue_card, text="Checking for new candidates...", fg_color="transparent",
                                   text_color=CARD_SUBTEXT, wraplength=260, justify=tk.LEFT, anchor="w")
        self.queue_lbl.pack(anchor=tk.W, fill=tk.X)
        ctk.CTkButton(queue_card, text="Refresh", command=self._refresh_queue).pack(anchor=tk.W, pady=(12, 0))
        sched_outer, sched_card = self._card(cards, "Scheduler", minheight=118)
        sched_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))
        self.sched_summary_var = tk.StringVar(value="Schedule is off")
        self.next_run_var = tk.StringVar(value="No next run")
        self.sched_last_var = tk.StringVar(value="No scheduled run yet")
        ctk.CTkLabel(sched_card, textvariable=self.sched_summary_var, fg_color="transparent", text_color=CARD_TEXT,
                     font=self.fonts["heading"], wraplength=260, justify=tk.LEFT, anchor="w"
                     ).pack(anchor=tk.W, fill=tk.X, pady=(8, 2))
        self.next_run_lbl = ctk.CTkLabel(sched_card, textvariable=self.next_run_var, fg_color="transparent",
                     text_color=CARD_SUBTEXT, anchor="w")
        self.next_run_lbl.pack(anchor=tk.W, fill=tk.X)
        ctk.CTkLabel(sched_card, textvariable=self.sched_last_var, fg_color="transparent", text_color=CARD_SUBTEXT,
                     font=self.fonts["caption"], anchor="w").pack(anchor=tk.W, fill=tk.X, pady=(4, 0))
        ctk.CTkButton(sched_card, text="Schedule", command=self._open_scheduler_window).pack(anchor=tk.W, pady=(12, 0))
        health_outer, health_card = self._card(right, "Status",
                                               "",
                                               minheight=172)
        health_outer.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self.health_summary_var = tk.StringVar(value="Checking services...")
        ctk.CTkLabel(health_card, textvariable=self.health_summary_var, fg_color="transparent", text_color=CARD_TEXT,
                     font=self.fonts["heading"], wraplength=340, justify=tk.LEFT, anchor="w"
                     ).pack(anchor=tk.W, fill=tk.X, pady=(8, 6))
        self.health_sp_var = tk.StringVar(value="SharePoint: checking")
        self.health_ai_var = tk.StringVar(value="AI: checking")
        self.health_jd_var = tk.StringVar(value="Jobs: checking")
        for var in (self.health_sp_var, self.health_ai_var, self.health_jd_var):
            ctk.CTkLabel(health_card, textvariable=var, fg_color="transparent", text_color=CARD_SUBTEXT,
                         wraplength=340, justify=tk.LEFT, anchor="w").pack(anchor=tk.W, fill=tk.X, pady=1)
        activity_outer, activity_card = self._card(right, "Recent Runs",
                                                   "",
                                                   minheight=118)
        activity_outer.pack(fill=tk.BOTH, expand=True)
        self.activity_list = tk.Listbox(activity_card, height=6, bg=CHART_SURFACE, fg=CARD_TEXT,
                                        selectbackground=CARD_ALT, relief=tk.FLAT, highlightthickness=0)
        self.activity_list.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        kpis = ctk.CTkFrame(shell, fg_color="transparent"); kpis.pack(fill=tk.X, pady=(10, 8))
        self._tile_vars = {}
        # "In queue" and "Next run" are deliberately NOT tiles here - they'd just
        # duplicate the Queue Health and Scheduler cards above pixel-for-pixel
        # (same StringVars, see _refresh_metrics/_sched_update_next).
        for key, caption in (("last", "Last run"), ("today", "Runs today"),
                             ("cands", "Scored / Rejected"), ("okfail", "OK / Failed"),
                             ("avg", "Avg Time")):
            outer, body = self._card(kpis, caption, compact=True)
            outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
            var = tk.StringVar(value="-")
            ctk.CTkLabel(body, textvariable=var, fg_color="transparent", text_color=CARD_TEXT,
                         font=self.fonts["heading"], wraplength=170, justify=tk.LEFT, anchor="w"
                         ).pack(anchor=tk.W, fill=tk.X, pady=(6, 0))
            self._tile_vars[key] = var
        analytics = ctk.CTkFrame(shell, fg_color="transparent"); analytics.pack(fill=tk.BOTH, expand=True)
        charts = ctk.CTkFrame(analytics, fg_color="transparent"); charts.pack(fill=tk.X)
        area_outer, area_body = self._card(charts, "Volume")
        area_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        self.area_canvas = tk.Canvas(area_body, height=110, bg=CHART_SURFACE, highlightthickness=0)
        self.area_canvas.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        dur_outer, dur_body = self._card(charts, "Duration")
        dur_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))
        self.dur_canvas = tk.Canvas(dur_body, height=110, bg=CHART_SURFACE, highlightthickness=0)
        self.dur_canvas.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self._area_pts, self._dur_pts = [], []
        self.area_canvas.bind("<Configure>", lambda _e: self._draw_area_chart())
        self.dur_canvas.bind("<Configure>", lambda _e: self._draw_duration_chart())
        self.area_canvas.bind("<Motion>", lambda e: self._chart_hover(self.area_canvas, self._area_pts, e))
        self.dur_canvas.bind("<Motion>", lambda e: self._chart_hover(self.dur_canvas, self._dur_pts, e))
        self.area_canvas.bind("<Leave>", lambda _e: self.area_canvas.delete("hover"))
        self.dur_canvas.bind("<Leave>", lambda _e: self.dur_canvas.delete("hover"))
        info_row = ctk.CTkFrame(analytics, fg_color="transparent"); info_row.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        cat_outer, cat_body = self._card(info_row, "Top Categories", "Most frequent categories in live SharePoint data.")
        cat_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        role_outer, role_body = self._card(info_row, "Top Role Matches", "Most common Suggested Role 1 values in the workbook.")
        role_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 4))
        snap_outer, snap_body = self._card(info_row, "Live Snapshot", "Queue, workbook, rejected, and connection details from SharePoint.")
        snap_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))
        self.category_list = tk.Listbox(cat_body, height=6, bg=CHART_SURFACE, fg=CARD_TEXT, selectbackground=CARD_ALT, relief=tk.FLAT, highlightthickness=0)
        self.category_list.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.role_list = tk.Listbox(role_body, height=6, bg=CHART_SURFACE, fg=CARD_TEXT, selectbackground=CARD_ALT, relief=tk.FLAT, highlightthickness=0)
        self.role_list.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.snapshot_list = tk.Listbox(snap_body, height=6, bg=CHART_SURFACE, fg=CARD_TEXT, selectbackground=CARD_ALT, relief=tk.FLAT, highlightthickness=0)
        self.snapshot_list.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.sp_log = self._make_log(shell, height=9)
        self._sched_refresh_list()
        self._refresh_recent_activity()

    def _jd_view_items(self):
        items = []
        for idx, url in enumerate(self._jd_urls):
            status = self._jd_url_status.get(url, {})
            items.append({
                "kind": "url", "index": idx, "title": url.rstrip("/").split("/")[-1] or "URL posting",
                "source_type": "URL", "url": url, "folder": "",
                "file_name": url.rstrip("/").split("/")[-1],
                "skills": status.get("skills", "(fetched at score time)"),
                "status": status.get("status", "Saved"),
            })
        for idx, desc in enumerate(self._jd_descs):
            from hiring_agent.extraction import _scan_skill_keywords
            skills = _scan_skill_keywords(desc.get("text", ""))
            items.append({
                "kind": "pasted", "index": idx, "title": desc.get("title", "Untitled JD"),
                "source_type": "Pasted", "url": "", "folder": "", "file_name": "",
                "skills": ", ".join(skills[:10]) or "(no skills found)", "status": "Ready",
                "text": desc.get("text", ""),
            })
        return items

    def _jd_check_sources(self):
        urls = list(self._jd_urls)
        if not urls:
            messagebox.showinfo("No URL sources", "Add at least one job posting URL first.")
            return
        self.jd_items_help_var.set("Checking URL sources...")

        def worker():
            from hiring_agent.jd_sources import fetch_jd_role
            statuses = {}
            for url in urls:
                try:
                    role = fetch_jd_role(url)
                    if role:
                        skills = ", ".join(role.get("skills", [])[:10]) or "(no skills found)"
                        statuses[url] = {
                            "status": "Reachable",
                            "skills": skills,
                            "title": role.get("title") or url,
                        }
                    else:
                        statuses[url] = {"status": "No skills found", "skills": "(unparsed)"}
                except Exception as e:
                    statuses[url] = {"status": f"Failed: {str(e)[:48]}", "skills": "(unreachable)"}

            def done():
                self._jd_url_status.update(statuses)
                self._jd_refresh_tree()
                self.jd_items_help_var.set("URL source check complete.")
            self.root.after(0, done)
        threading.Thread(target=worker, daemon=True).start()

    def _jd_copy_selected(self):
        sel = self.jd_tree.selection()
        if not sel:
            return
        item = self._jd_items[int(sel[0])]
        target = item.get("url") or self.jd_sp_preview_url_var.get() or item.get("folder")
        self._copy_text(target)

    def _jd_open_selected(self):
        sel = self.jd_tree.selection()
        if not sel:
            return
        item = self._jd_items[int(sel[0])]
        target = item.get("url") or self.jd_sp_preview_url_var.get()
        self._open_link(target)

    def _jd_url_changed(self):
        url = self.jd_sp_url_var.get().strip()
        if not url:
            self.jd_sp_preview_url_var.set(compose_sharepoint_url(
                self.jd_sp_host_var.get(), self.jd_sp_site_var.get(), self.jd_sp_folder_var.get()))
            return
        p = parse_sharepoint_url(url)
        if p["hostname"]:
            self.jd_sp_host_var.set(p["hostname"])
        if p["site_path"]:
            self.jd_sp_site_var.set(p["site_path"])
        folder = (p["folder"] or "").lstrip("/")
        if p["file"]:
            folder = f"{folder}/{p['file']}".strip("/").rsplit("/", 1)[0]
        self.jd_sp_folder_var.set(folder)
        self.jd_sp_preview_url_var.set(compose_sharepoint_url(
            self.jd_sp_host_var.get(), self.jd_sp_site_var.get(), self.jd_sp_folder_var.get()))

    def _jd_refresh_preview_tree(self):
        if not hasattr(self, "jd_preview_tree"):
            return
        self.jd_preview_tree.delete(*self.jd_preview_tree.get_children())
        for item in self._jd_sp_preview:
            self.jd_preview_tree.insert("", tk.END, values=(
                item.get("name", ""), item.get("title", ""), item.get("folder", ""),
                item.get("file_type", ""), item.get("status", ""),
                ", ".join(item.get("skills", [])[:8]) or "-",
            ))

    def _jd_show_selection(self, *_args):
        if not hasattr(self, "jd_tree"):
            return
        sel = self.jd_tree.selection()
        if not sel:
            return
        item = self._jd_items[int(sel[0])]
        self.jd_detail_title_var.set(item.get("title", ""))
        self.jd_detail_source_var.set(item.get("source_type", ""))
        self.jd_detail_url_var.set(item.get("url", "") or "-")
        self.jd_detail_folder_var.set(item.get("folder", "") or "-")
        self.jd_detail_file_var.set(item.get("file_name", "") or "-")
        self.jd_detail_status_var.set(item.get("status", "") or "-")
        self.jd_detail_skills_var.set(item.get("skills", "") or "-")

    def _jd_load(self):
        from hiring_agent.jd_sources import load_jd_sources
        cfg = load_jd_sources()
        self.jd_enabled_var.set(cfg["enabled"])
        self._jd_urls = list(cfg["urls"])
        self._jd_descs = list(cfg.get("descriptions", []))
        self.jd_sp_host_var.set(cfg.get("sharepoint_jd_hostname", ""))
        self.jd_sp_site_var.set(cfg.get("sharepoint_jd_site", ""))
        self.jd_sp_folder_var.set(cfg.get("sharepoint_jd_folder", ""))
        self.jd_sp_preview_url_var.set(compose_sharepoint_url(
            self.jd_sp_host_var.get(), self.jd_sp_site_var.get(), self.jd_sp_folder_var.get()))
        self._jd_sp_preview = []
        self._jd_refresh_tree()
        self._jd_refresh_preview_tree()

    def _jd_refresh_tree(self):
        self._jd_items = self._jd_view_items()
        self.jd_tree.delete(*self.jd_tree.get_children())
        for idx, item in enumerate(self._jd_items):
            self.jd_tree.insert("", tk.END, iid=str(idx), values=(
                item["title"], item["source_type"], item["url"], item["folder"],
                item["file_name"], item["skills"], item["status"],
            ))
        total = len(self._jd_items)
        status = f"{total} job posting(s)" if total else "Using DriverAI's default roles"
        self.jd_count_label.configure(text=status)
        self.jd_items_help_var.set(f"{status}. SharePoint folder preview appears below after Check Now.")
        self._jd_show_selection()

    def _jd_add_url(self):
        url = self.jd_url_var.get().strip()
        if not url:
            messagebox.showwarning("No link", "Enter a job posting link first.")
            return
        if url in self._jd_urls:
            messagebox.showinfo("Already added", "This link is already in the list.")
            return
        self._jd_urls.append(url)
        self.jd_url_var.set("")
        self._jd_refresh_tree()

    def _jd_add_text(self):
        title = self.jd_title_var.get().strip() or "Untitled posting"
        text = self.jd_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("Nothing pasted", "Paste the job posting text first.")
            return
        self._jd_descs.append({"title": title, "text": text})
        self.jd_title_var.set("")
        self.jd_text.delete("1.0", tk.END)
        self._jd_refresh_tree()

    def _jd_edit_selected(self):
        sel = self.jd_tree.selection()
        if not sel:
            messagebox.showinfo("Select one", "Select a posting to edit first.")
            return
        item = self._jd_items[int(sel[0])]
        if item["kind"] == "url":
            self.jd_url_var.set(item["url"])
            self._jd_urls.pop(item["index"])
        else:
            self.jd_title_var.set(item["title"])
            self.jd_text.delete("1.0", tk.END)
            self.jd_text.insert("1.0", item.get("text", ""))
            self._jd_descs.pop(item["index"])
        self._jd_refresh_tree()

    def _jd_remove(self):
        sel = self.jd_tree.selection()
        if not sel:
            messagebox.showinfo("Select one", "Select a posting to remove first.")
            return
        for sel_id in sorted((int(x) for x in sel), reverse=True):
            item = self._jd_items[sel_id]
            if item["kind"] == "url":
                self._jd_urls.pop(item["index"])
            else:
                self._jd_descs.pop(item["index"])
        self._jd_refresh_tree()

    def _jd_save_sp_folder(self):
        from hiring_agent.jd_sources import load_jd_sources, save_jd_sources
        cfg = load_jd_sources()
        save_jd_sources(
            cfg["enabled"], cfg["urls"], cfg.get("descriptions", []),
            sharepoint_jd_folder=self.jd_sp_folder_var.get().strip(),
            sharepoint_jd_hostname=self.jd_sp_host_var.get().strip(),
            sharepoint_jd_site=self.jd_sp_site_var.get().strip(),
        )
        self.jd_sp_preview_url_var.set(compose_sharepoint_url(
            self.jd_sp_host_var.get(), self.jd_sp_site_var.get(), self.jd_sp_folder_var.get()))
        self.jd_sp_status.configure(text="Folder saved", text_color=SUCCESS)

    def _jd_test_sp_folder(self):
        folder = self.jd_sp_folder_var.get().strip()
        if not folder:
            messagebox.showwarning("No folder yet", "Paste a SharePoint folder link first.")
            return
        self.jd_sp_status.configure(text="Checking...", text_color=CARD_SUBTEXT)
        hostname = self.jd_sp_host_var.get().strip()
        site = self.jd_sp_site_var.get().strip()
        def worker():
            try:
                from hiring_agent.jd_sources import fetch_jd_folder_catalog, build_jd_cache
                from hiring_agent.config import SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH
                catalog = fetch_jd_folder_catalog(hostname or SHAREPOINT_HOSTNAME,
                                                  site or SHAREPOINT_SITE_PATH, folder)
                # Persist the parsed roles so scoring runs read this cache instantly
                # instead of re-downloading every JD each run (the "one-time scan").
                cached = build_jd_cache(catalog=catalog) if catalog else []
                def done():
                    self._jd_sp_preview = catalog
                    self._jd_refresh_preview_tree()
                    self.jd_sp_status.configure(
                        text=(f"Scanned {len(catalog)} JD file(s) - {len(cached)} role(s) cached"
                              if catalog else "No JD files found"),
                        text_color=SUCCESS if catalog else WARNING)
                self.root.after(0, done)
            except Exception as e:
                self.log_q.put((self.sp_log, f"[job postings] Couldn't check: {e}"))
                self.root.after(0, lambda: self.jd_sp_status.configure(
                    text="Couldn't check - see the Home tab log", text_color=ERROR))
        threading.Thread(target=worker, daemon=True).start()

    def _build_jd_tab(self):
        tab = self._scroll_tab("Job Descriptions")
        top = ctk.CTkFrame(tab, fg_color="transparent")
        top.pack(fill=tk.X, pady=(10, 4))
        self.jd_enabled_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(top, text="Match candidates against my own job postings",
                        variable=self.jd_enabled_var, onvalue=True, offvalue=False
                        ).pack(side=tk.LEFT)
        self.jd_count_label = ctk.CTkLabel(top, text="", fg_color="transparent", text_color=ACCENT)
        self.jd_count_label.pack(side=tk.RIGHT, padx=10)
        online_card, online = self._card(tab, "Source Manager")
        online_card.pack(fill=tk.X, pady=4)
        url_row = ctk.CTkFrame(online, fg_color="transparent"); url_row.pack(fill=tk.X, pady=(6, 2))
        ctk.CTkLabel(url_row, text="Job posting URL", fg_color="transparent", text_color=CARD_TEXT
                     ).pack(side=tk.LEFT)
        self.jd_url_var = tk.StringVar()
        ctk.CTkEntry(url_row, textvariable=self.jd_url_var).pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)
        ctk.CTkButton(url_row, text="Add URL", command=self._jd_add_url).pack(side=tk.LEFT)
        self.jd_sp_host_var = tk.StringVar()
        self.jd_sp_site_var = tk.StringVar()
        self.jd_sp_folder_var = tk.StringVar()
        self.jd_sp_preview_url_var = tk.StringVar()
        sp_row = ctk.CTkFrame(online, fg_color="transparent"); sp_row.pack(fill=tk.X, pady=2)
        ctk.CTkLabel(sp_row, text="SharePoint / OneDrive JD folder", fg_color="transparent",
                     text_color=CARD_TEXT).pack(side=tk.LEFT)
        self.jd_sp_url_var = tk.StringVar()
        ctk.CTkEntry(sp_row, textvariable=self.jd_sp_url_var).pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)
        ctk.CTkButton(sp_row, text="Save Folder", command=self._jd_save_sp_folder).pack(side=tk.LEFT)
        ctk.CTkButton(sp_row, text="Check Now", command=self._jd_test_sp_folder).pack(side=tk.LEFT, padx=6)
        self.jd_sp_status = ctk.CTkLabel(sp_row, text="", fg_color="transparent", text_color=CARD_SUBTEXT)
        self.jd_sp_status.pack(side=tk.LEFT, padx=8)
        self.jd_sp_url_var.trace_add("write", lambda *_: self._jd_url_changed())
        ctk.CTkLabel(online, textvariable=self.jd_sp_preview_url_var, fg_color="transparent",
                     text_color=CARD_SUBTEXT, wraplength=900, justify="left", anchor="w"
                     ).pack(anchor=tk.W, fill=tk.X, pady=(0, 6))
        pasted_card, pasted = self._card(tab, "Add Pasted JD")
        pasted_card.pack(fill=tk.X, pady=4)
        paste_row = ctk.CTkFrame(pasted, fg_color="transparent"); paste_row.pack(fill=tk.X, pady=(6, 2))
        ctk.CTkLabel(paste_row, text="Title", fg_color="transparent", text_color=CARD_TEXT
                     ).pack(side=tk.LEFT)
        self.jd_title_var = tk.StringVar()
        ctk.CTkEntry(paste_row, textvariable=self.jd_title_var, width=320).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(paste_row, text="Add Pasted JD", command=self._jd_add_text).pack(side=tk.LEFT, padx=8)
        self.jd_text = tk.Text(pasted, height=4, wrap=tk.WORD, bg=CHART_SURFACE, fg=CARD_TEXT,
                               insertbackground=CARD_TEXT, relief=tk.FLAT, highlightthickness=0)
        self.jd_text.pack(fill=tk.X, pady=(2, 6))
        list_card, list_frame = self._card(tab, "Configured JD Sources")
        list_card.pack(fill=tk.BOTH, expand=True, pady=4)
        cols = ("Title", "Type", "Full URL", "Folder", "File", "Skills", "Status")
        self.jd_tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=7)
        for col in cols:
            self.jd_tree.heading(col, text=col)
        self.jd_tree.column("Title", width=180, stretch=True)
        self.jd_tree.column("Type", width=90, stretch=False)
        self.jd_tree.column("Full URL", width=250, stretch=True)
        self.jd_tree.column("Folder", width=180, stretch=True)
        self.jd_tree.column("File", width=130, stretch=True)
        self.jd_tree.column("Skills", width=260, stretch=True)
        self.jd_tree.column("Status", width=110, stretch=False)
        ysb = ttk.Scrollbar(list_frame, command=self.jd_tree.yview)
        xsb = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.jd_tree.xview)
        self.jd_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        xsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.jd_tree.pack(fill=tk.BOTH, expand=True)
        self.jd_tree.bind("<<TreeviewSelect>>", self._jd_show_selection)
        btn_row = ctk.CTkFrame(tab, fg_color="transparent"); btn_row.pack(fill=tk.X, pady=(0, 6))
        ctk.CTkButton(btn_row, text="Edit Selected", command=self._jd_edit_selected).pack(side=tk.LEFT)
        ctk.CTkButton(btn_row, text="Remove", command=self._jd_remove).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(btn_row, text="Save", command=self._jd_save).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(btn_row, text="Check URL Sources", command=self._jd_check_sources).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(btn_row, text="Open Selected Link", command=self._jd_open_selected).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(btn_row, text="Copy Selected Link", command=self._jd_copy_selected).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(btn_row, text="See DriverAI Default Roles", command=self._jd_show_builtin).pack(side=tk.RIGHT)
        self.jd_items_help_var = tk.StringVar(value="")
        ctk.CTkLabel(tab, textvariable=self.jd_items_help_var, fg_color="transparent",
                     text_color=CARD_SUBTEXT, anchor="w").pack(anchor=tk.W, fill=tk.X)
        detail_card, detail = self._card(tab, "Selected Source Details")
        detail_card.pack(fill=tk.X, pady=4)
        self.jd_detail_title_var = tk.StringVar(value="-")
        self.jd_detail_source_var = tk.StringVar(value="-")
        self.jd_detail_url_var = tk.StringVar(value="-")
        self.jd_detail_folder_var = tk.StringVar(value="-")
        self.jd_detail_file_var = tk.StringVar(value="-")
        self.jd_detail_status_var = tk.StringVar(value="-")
        self.jd_detail_skills_var = tk.StringVar(value="-")
        self._detail_row(detail, "Title", self.jd_detail_title_var)
        self._detail_row(detail, "Source type", self.jd_detail_source_var)
        self._detail_row(detail, "Full URL", self.jd_detail_url_var, copyable=True, linkable=True)
        self._detail_row(detail, "Folder", self.jd_detail_folder_var, copyable=True)
        self._detail_row(detail, "File name", self.jd_detail_file_var, copyable=True)
        self._detail_row(detail, "Status", self.jd_detail_status_var)
        self._detail_row(detail, "Parsed skills", self.jd_detail_skills_var, copyable=True)
        preview_card, preview = self._card(tab, "SharePoint Folder Preview")
        preview_card.pack(fill=tk.BOTH, expand=True, pady=(4, 8))
        pcols = ("File", "Inferred Title", "Folder", "Type", "Status", "Skills")
        self.jd_preview_tree = ttk.Treeview(preview, columns=pcols, show="headings", height=5)
        for col in pcols:
            self.jd_preview_tree.heading(col, text=col)
        py = ttk.Scrollbar(preview, command=self.jd_preview_tree.yview)
        px = ttk.Scrollbar(preview, orient=tk.HORIZONTAL, command=self.jd_preview_tree.xview)
        self.jd_preview_tree.configure(yscrollcommand=py.set, xscrollcommand=px.set)
        py.pack(side=tk.RIGHT, fill=tk.Y)
        px.pack(side=tk.BOTTOM, fill=tk.X)
        self.jd_preview_tree.pack(fill=tk.BOTH, expand=True)
        self._jd_load()

    def _sp_refresh_overview(self):
        resume_url = compose_sharepoint_url(self.sp_hostname_var.get().strip(),
                                            self.sp_site_var.get().strip(),
                                            self.sp_resumes_var.get().strip())
        workbook_url = self._compose_workbook_url()
        self.sp_summary_resume_var.set(resume_url or "-")
        self.sp_summary_workbook_var.set(workbook_url or "-")
        self.sp_summary_file_var.set(self._compose_workbook_path().rsplit("/", 1)[-1])
        site_text = f"{self.sp_hostname_var.get().strip()}{self.sp_site_var.get().strip()}"
        self.sp_summary_site_var.set(site_text or "-")
        self.sp_summary_table_var.set(self.sp_table_var.get().strip() or "-")
        sheets = self._sp_live_meta.get("worksheets") or []
        self.sp_summary_sheets_var.set(", ".join(sheets) if sheets else "(discover after Test Connection)")

    def _apply_sharepoint_metadata(self, meta: dict):
        self._sp_live_meta = meta or {}
        if hasattr(self, "sp_summary_sheets_var"):
            self.sp_summary_sheets_var.set(", ".join(meta.get("worksheets", [])) or "(none found)")
            self.sp_meta_status_var.set(
                f"Connected - {len(meta.get('worksheets', []))} sheet(s), "
                f"{len(meta.get('tables', []))} table(s), {len(meta.get('columns', []))} column(s)"
            )
            self.sp_meta_site_var.set(f"{meta.get('hostname', '')}{meta.get('site_path', '')}" or "-")
            self.sp_meta_workbook_var.set(meta.get("workbook_path", "-"))
            self.sp_meta_tables_var.set(", ".join(meta.get("tables", [])) or "(none)")
            self.sp_meta_columns_var.set(", ".join(meta.get("columns", [])) or "(none)")
        self._sp_refresh_overview()

    def _sp_url_changed(self):
        url = self.sp_url_var.get().strip()
        if url:
            p = parse_sharepoint_url(url)
            if p["hostname"] and p["site_path"]:
                self.sp_hostname_var.set(p["hostname"])
                self.sp_site_var.set(p["site_path"])
                self.sp_resumes_var.set(p["folder"])
        self._sp_refresh_overview()

    def _sp_wburl_changed(self):
        url = self.sp_wb_url_var.get().strip()
        if url:
            p = parse_sharepoint_url(url)
            if p["file"]:
                self.sp_workbook_var.set(f"{p['folder']}/{p['file']}")
        self._sp_refresh_overview()

    def _sp_test_connection(self):
        self.sp_conn_status.configure(text="Checking...", text_color=CARD_SUBTEXT)
        def worker():
            try:
                from sharepoint_client import SharePointClient
                client = SharePointClient()
                client.ensure_workbook()
                meta = client.workbook_metadata()
                self.log_q.put((self.sp_tab_log, "[test] Connected to SharePoint OK"))
                self.root.after(0, lambda: self.sp_conn_status.configure(text="Connected", text_color=SUCCESS))
                self.root.after(0, lambda m=meta: self._apply_sharepoint_metadata(m))
            except Exception as e:
                self.log_q.put((self.sp_tab_log, f"[test] Couldn't connect: {e}"))
                self.root.after(0, lambda: self.sp_conn_status.configure(
                    text="Couldn't connect - see details below", text_color=ERROR))
        threading.Thread(target=worker, daemon=True).start()

    def _sp_show_columns(self):
        self.sp_conn_status.configure(text="Checking...", text_color=CARD_SUBTEXT)
        def worker():
            try:
                from sharepoint_client import SharePointClient
                meta = SharePointClient().workbook_metadata()
                self.log_q.put((self.sp_tab_log, f"[columns] {len(meta['columns'])} columns: {', '.join(meta['columns'])}"))
                self.root.after(0, lambda m=meta: self._apply_sharepoint_metadata(m))
                self.root.after(0, lambda: self.sp_conn_status.configure(
                    text=f"Found {len(meta['columns'])} columns", text_color=SUCCESS))
            except Exception as e:
                self.log_q.put((self.sp_tab_log, f"[columns] Couldn't read: {e}"))
                self.root.after(0, lambda: self.sp_conn_status.configure(
                    text="Couldn't connect - see details below", text_color=ERROR))
        threading.Thread(target=worker, daemon=True).start()

    def _build_sharepoint_tab(self):
        tab = self._scroll_tab("SharePoint")
        try:
            from dotenv import dotenv_values
            env = dotenv_values(APP_DIR / ".env")
        except Exception:
            env = {}
        specs = [
            ("TENANT_ID", "sp_tenant_var"), ("CLIENT_ID", "sp_client_var"),
            ("CLIENT_SECRET", "sp_secret_var"), ("SHAREPOINT_HOSTNAME", "sp_hostname_var"),
            ("SHAREPOINT_SITE_PATH", "sp_site_var"), ("SHAREPOINT_TABLE", "sp_table_var"),
            ("SHAREPOINT_RESUMES_FOLDER", "sp_resumes_var"), ("SHAREPOINT_WORKBOOK", "sp_workbook_var"),
        ]
        self._sp_env_keys = list(specs)
        for key, attr in specs:
            setattr(self, attr, tk.StringVar(value=env.get(key, os.getenv(key, ""))))
        shell = ctk.CTkFrame(tab, fg_color="transparent")
        shell.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        summary_card, summary = self._card(shell, "SharePoint Data Map")
        summary_card.pack(fill=tk.X, pady=(0, 8))
        self.sp_summary_resume_var = tk.StringVar(value="-")
        self.sp_summary_workbook_var = tk.StringVar(value="-")
        self.sp_summary_file_var = tk.StringVar(value="-")
        self.sp_summary_site_var = tk.StringVar(value="-")
        self.sp_summary_table_var = tk.StringVar(value="-")
        self.sp_summary_sheets_var = tk.StringVar(value="(discover after Test Connection)")
        self._detail_row(summary, "Resume input URL", self.sp_summary_resume_var,
                         copyable=True, linkable=True)
        self._detail_row(summary, "Workbook URL", self.sp_summary_workbook_var,
                         copyable=True, linkable=True)
        self._detail_row(summary, "Workbook file", self.sp_summary_file_var, copyable=True)
        self._detail_row(summary, "Site", self.sp_summary_site_var,
                         copyable=True, linkable=True)
        self._detail_row(summary, "Candidate table", self.sp_summary_table_var, copyable=True)
        self._detail_row(summary, "Known sheet names", self.sp_summary_sheets_var, copyable=True)
        info_card, info = self._card(shell, "Readable Locations")
        info_card.pack(fill=tk.X, pady=4)
        # _card()'s title/subtitle are pack()-managed inside `info` - grid() can't be
        # mixed with pack() in the same parent, so this section's grid layout goes in
        # its own nested frame instead of directly into `info`.
        info_grid = ctk.CTkFrame(info, fg_color="transparent")
        info_grid.pack(fill=tk.BOTH, expand=True)
        ctk.CTkLabel(info_grid, text="Candidate list name", fg_color="transparent", text_color=CARD_TEXT
                     ).grid(row=0, column=0, sticky=tk.W, padx=(0, 4), pady=(8, 2))
        ctk.CTkEntry(info_grid, textvariable=self.sp_table_var, width=500
                     ).grid(row=0, column=1, sticky=tk.EW, padx=4, pady=(8, 2))
        ctk.CTkLabel(info_grid, text="Resume folder URL", fg_color="transparent", text_color=CARD_TEXT
                     ).grid(row=1, column=0, sticky=tk.W, padx=(0, 4), pady=2)
        self.sp_url_var = tk.StringVar(value=compose_sharepoint_url(self.sp_hostname_var.get(), self.sp_site_var.get(), self.sp_resumes_var.get()))
        ctk.CTkEntry(info_grid, textvariable=self.sp_url_var, width=500
                     ).grid(row=1, column=1, sticky=tk.EW, padx=4, pady=2)
        ctk.CTkLabel(info_grid, text="Workbook URL (optional)", fg_color="transparent", text_color=CARD_TEXT
                     ).grid(row=2, column=0, sticky=tk.W, padx=(0, 4), pady=(8, 2))
        self.sp_wb_url_var = tk.StringVar(value=self._compose_workbook_url())
        ctk.CTkEntry(info_grid, textvariable=self.sp_wb_url_var, width=500
                     ).grid(row=2, column=1, sticky=tk.EW, padx=4, pady=(8, 2))
        self.sp_url_var.trace_add("write", lambda *_: self._sp_url_changed())
        self.sp_wb_url_var.trace_add("write", lambda *_: self._sp_wburl_changed())
        info_grid.columnconfigure(1, weight=1)
        adv_card, adv = self._card(shell, "Technical Fields and Credentials")
        adv_card.pack(fill=tk.X, pady=4)
        self._sp_adv_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(adv, text="Show exact technical fields", variable=self._sp_adv_var,
                        onvalue=True, offvalue=False,
                        command=lambda: self._sp_adv_frame.pack(fill=tk.X, pady=4)
                        if self._sp_adv_var.get() else self._sp_adv_frame.pack_forget()
                        ).pack(anchor=tk.W, pady=(8, 0))
        self._sp_adv_frame = ctk.CTkFrame(adv, fg_color="transparent")
        rows = [
            ("Tenant ID", self.sp_tenant_var), ("Client ID", self.sp_client_var),
            ("Client Secret", self.sp_secret_var), ("Hostname", self.sp_hostname_var),
            ("Site Path", self.sp_site_var), ("Resumes Folder", self.sp_resumes_var),
            ("Workbook Path", self.sp_workbook_var),
        ]
        for idx, (label, var) in enumerate(rows):
            ctk.CTkLabel(self._sp_adv_frame, text=label, fg_color="transparent", text_color=CARD_TEXT
                         ).grid(row=idx, column=0, sticky=tk.W, padx=(0, 4), pady=2)
            ctk.CTkEntry(self._sp_adv_frame, textvariable=var, width=500,
                      show="*" if label == "Client Secret" else "").grid(row=idx, column=1, sticky=tk.EW, pady=2)
        self._sp_adv_frame.columnconfigure(1, weight=1)
        meta_card, meta = self._card(shell, "Live Workbook Metadata")
        meta_card.pack(fill=tk.X, pady=4)
        self.sp_meta_status_var = tk.StringVar(value="Run Test Connection to load metadata")
        self.sp_meta_workbook_var = tk.StringVar(value="-")
        self.sp_meta_site_var = tk.StringVar(value="-")
        self.sp_meta_tables_var = tk.StringVar(value="-")
        self.sp_meta_columns_var = tk.StringVar(value="-")
        self._detail_row(meta, "Status", self.sp_meta_status_var)
        self._detail_row(meta, "Resolved site", self.sp_meta_site_var,
                         copyable=True, linkable=True)
        self._detail_row(meta, "Workbook path", self.sp_meta_workbook_var, copyable=True)
        self._detail_row(meta, "Table names", self.sp_meta_tables_var, copyable=True)
        self._detail_row(meta, "Columns", self.sp_meta_columns_var, copyable=True)
        btns = ctk.CTkFrame(shell, fg_color="transparent"); btns.pack(fill=tk.X, pady=4)
        ctk.CTkButton(btns, text="Save", command=self._sp_save_env).pack(side=tk.LEFT)
        ctk.CTkButton(btns, text="Test Connection", command=self._sp_test_connection).pack(side=tk.LEFT, padx=6)
        ctk.CTkButton(btns, text="Refresh Metadata", command=self._sp_show_columns).pack(side=tk.LEFT)
        self.sp_conn_status = ctk.CTkLabel(btns, text="", fg_color="transparent", text_color=CARD_SUBTEXT)
        self.sp_conn_status.pack(side=tk.LEFT, padx=8)
        self.sp_tab_log = self._make_log(shell, height=10)
        self._sp_refresh_overview()

    def _run_health_check(self):
        def worker():
            sp_ok = False
            try:
                from hiring_agent.config import SHAREPOINT_CONFIGURED
                sp_ok = SHAREPOINT_CONFIGURED
            except Exception:
                pass
            ollama_ok = False
            try:
                import requests
                host = self.settings.get("ollama_host", "http://localhost:11434")
                r = requests.get(f"{host.rstrip('/')}/api/tags", timeout=3)
                ollama_ok = r.ok
            except Exception:
                pass
            jd_on, n_jd = False, 0
            try:
                from hiring_agent.jd_sources import load_jd_sources
                cfg = load_jd_sources()
                n_jd = len(cfg.get("urls", [])) + len(cfg.get("descriptions", []))
                jd_on = cfg["enabled"] and (n_jd > 0 or cfg.get("sharepoint_jd_folder"))
            except Exception:
                pass
            all_ok = sp_ok and ollama_ok
            summary = ("Ready" if all_ok
                       else "Set up SharePoint" if not sp_ok
                       else "Ready - keyword mode")
            brain = "AI" if ollama_ok else "Keyword"
            jd_note = f"{n_jd} custom JDs" if jd_on else "Default roles"
            detail = f"Scoring: {brain}  |  Jobs: {jd_note}  |  SharePoint: {'connected' if sp_ok else 'not connected'}"
            def done():
                self.status_lbl.configure(text=summary, text_color=SUCCESS if all_ok else WARNING)
                self.status_detail.configure(text=detail)
                if hasattr(self, "health_summary_var"):
                    self.health_summary_var.set(summary)
                    self.health_sp_var.set(f"SharePoint: {'Connected' if sp_ok else 'Not connected'}")
                    self.health_ai_var.set(f"AI: {'Available' if ollama_ok else 'Keyword mode'}")
                    self.health_jd_var.set(f"Jobs: {jd_note}")
            self.root.after(0, done)
        threading.Thread(target=worker, daemon=True).start()

    def _build_status_bar(self):
        bar = ctk.CTkFrame(self.root, fg_color="transparent")
        bar.pack(fill=tk.X, side=tk.BOTTOM, padx=8, pady=(0, 4))
        self.status_lbl = ctk.CTkLabel(bar, text="Getting ready...", fg_color="transparent", text_color=CARD_SUBTEXT)
        self.status_lbl.pack(side=tk.LEFT)
        ctk.CTkLabel(bar, text=f"v{APP_VERSION}", fg_color="transparent", text_color=CARD_SUBTEXT
                     ).pack(side=tk.RIGHT)
        self.status_detail = ctk.CTkLabel(bar, text="", fg_color="transparent", text_color=CARD_SUBTEXT)
        self.status_detail.pack(side=tk.RIGHT, padx=10)

def main():
    if not acquire_single_instance():
        hidden = tk.Tk()
        hidden.withdraw()
        messagebox.showwarning(
            "Already running",
            "DriverAI Hiring Agent is already open.\n"
            "Running two copies would double-trigger the scheduler.")
        hidden.destroy()
        return
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    app = HiringApp(root)
    root.protocol("WM_DELETE_WINDOW", lambda: _on_close(root, app))
    root.mainloop()

def _on_close(root, app):
    try:
        app.settings["window_geometry"] = root.geometry()
        save_settings(app.settings)
    except Exception:
        pass
    if hasattr(app, "_sched_stop_event"):
        app._sched_stop_event.set()
    if app.proc and app.proc.poll() is None:
        app.proc.terminate()
    root.destroy()

if __name__ == "__main__":
    main()
