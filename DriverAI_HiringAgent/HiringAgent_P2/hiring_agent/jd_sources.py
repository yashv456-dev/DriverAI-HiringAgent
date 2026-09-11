"""JD sources: load/save config, fetch JD pages, parse JD text, build role dicts.

Supports three JD source types (all combinable, all optional):
  1. sharepoint_jd_folder — auto-scans a SharePoint/OneDrive folder, downloads every
     PDF/DOCX and treats each file as one JD. Uses the same Graph API credentials
     as the scoring client (TENANT_ID / CLIENT_ID / CLIENT_SECRET in .env).
  2. urls — public web pages (job boards, company career pages). Fetched and parsed.
  3. descriptions — JD text pasted directly into jd_sources.json.

When enabled=false (default) the scorer uses the built-in default_roles from config.yaml.
"""

import json
import os
import re
from html import unescape

from hiring_agent.config import JD_SOURCES_FILE, JD_CACHE_FILE, logger
from hiring_agent.extraction import _scan_skill_keywords, html_to_text
from hiring_agent.jd_identity import stamp_role_identity, stamp_roles
from hiring_agent.scoring import get_open_roles

_GRAPH = "https://graph.microsoft.com/v1.0"


class JDFolderScanError(RuntimeError):
    """A SharePoint JD folder IS configured but yielded nothing.

    Raised (2026-08-21) so a broken folder scan can never again be mistaken for a
    successful one. Every failure inside fetch_jd_folder_catalog - bad token, unresolvable
    site, unresolvable folder path - logs a warning and returns [], which is
    indistinguishable from 'the folder is genuinely empty'. When other JD sources still
    produced roles, the caller saw a non-empty list, wrote a cache built entirely from
    stale/pasted text, and reported success.

    Real incident: `sharepoint_jd_folder` was set to 'Documents/Staffing/PDs' when the true
    path is 'Staffing/PDs' (Staffing sits at the drive root; the 'Documents' folder beside
    it is unrelated). Graph 404'd on every refresh, the scan returned [], and 85 stale
    pasted descriptions carried the cache - so the live folder had not been read at all,
    the 24h auto-refresh was a no-op, and folder_fingerprint stayed null, which silently
    disabled file-change detection too. Nothing surfaced above WARNING in an unattended run.
    """


def load_jd_sources() -> dict:
    """Read the JD config file. Returns a dict with all supported keys."""
    if JD_SOURCES_FILE.exists():
        try:
            data = json.loads(JD_SOURCES_FILE.read_text(encoding="utf-8"))
            return {
                "enabled": bool(data.get("enabled", False)),
                "urls": [u for u in data.get("urls", []) if isinstance(u, str) and u.strip()],
                "descriptions": [
                    d for d in data.get("descriptions", [])
                    if isinstance(d, dict) and d.get("text", "").strip()
                ],
                # SharePoint/OneDrive folder scanner (all three keys needed for personal sites)
                "sharepoint_jd_folder":   (data.get("sharepoint_jd_folder")   or "").strip(),
                "sharepoint_jd_hostname": (data.get("sharepoint_jd_hostname") or "").strip(),
                "sharepoint_jd_site":     (data.get("sharepoint_jd_site")     or "").strip(),
            }
        except Exception as e:
            logger.warning(f"Could not read {JD_SOURCES_FILE.name}: {e}")
    return {
        "enabled": False, "urls": [], "descriptions": [],
        "sharepoint_jd_folder": "", "sharepoint_jd_hostname": "", "sharepoint_jd_site": "",
    }


def save_jd_sources(enabled: bool, urls: list, descriptions: list = None,
                    sharepoint_jd_folder: str = "", sharepoint_jd_hostname: str = "",
                    sharepoint_jd_site: str = "") -> None:
    """Persist the JD config (URLs, pasted descriptions, and optional SP folder path)."""
    clean_urls = [u.strip() for u in urls if isinstance(u, str) and u.strip()]
    clean_descs = []
    for d in (descriptions or []):
        if isinstance(d, dict) and d.get("text", "").strip():
            clean_descs.append({
                "title": (d.get("title") or "").strip() or "Untitled JD",
                "text": d["text"].strip(),
            })
    payload = {
        "enabled": bool(enabled),
        "sharepoint_jd_hostname": sharepoint_jd_hostname.strip(),
        "sharepoint_jd_site":     sharepoint_jd_site.strip(),
        "sharepoint_jd_folder":   sharepoint_jd_folder.strip(),
        "urls": clean_urls,
        "descriptions": clean_descs,
    }
    JD_SOURCES_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")


# ── SharePoint / OneDrive folder scanner ─────────────────────────────────────

def _graph_token() -> str:
    """Acquire an app-only Graph token using the same credentials as the scoring client."""
    import requests as _r
    tenant = (os.getenv("TENANT_ID") or "").strip()
    client_id = (os.getenv("CLIENT_ID") or "").strip()
    client_secret = (os.getenv("CLIENT_SECRET") or "").strip()
    if not all([tenant, client_id, client_secret]):
        raise RuntimeError("TENANT_ID / CLIENT_ID / CLIENT_SECRET not set in .env")
    resp = _r.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    token = data.get("access_token")
    if not token:
        raise RuntimeError(
            f"OAuth token request failed: {data.get('error_description') or data.get('error') or data}"
        )
    return token


def fetch_jd_folder_catalog(hostname: str, site_path: str, folder: str) -> list:
    """Return metadata + parsed-role info for every PDF/DOCX JD in a SharePoint folder.

    hostname  — e.g. 'led1234567-my.sharepoint.com'
    site_path — e.g. '/personal/tracys_driverai_io'
    folder    — e.g. 'Documents/Projects/Automation/Recruiting PDs'

    The app registration needs Files.Read.All (Application) admin-consented, in addition
    to the Sites permissions already needed for scoring.

    Returns a list of dicts with:
      name, title, folder, url, file_type, status, skills, role
    """
    from hiring_agent.extraction import extract_text_from_bytes
    import requests as _r

    try:
        token = _graph_token()
    except Exception as e:
        logger.warning(f"JD folder scan: cannot get Graph token: {e}")
        return []

    headers = {"Authorization": f"Bearer {token}"}
    folder = folder.strip("/")

    # 1. Resolve the SharePoint site ID.
    try:
        site_resp = _r.get(
            f"{_GRAPH}/sites/{hostname}:{site_path}",
            headers=headers, timeout=20,
        )
        site_resp.raise_for_status()
        site_id = site_resp.json()["id"]
    except Exception as e:
        logger.warning(f"JD folder scan: cannot resolve site '{hostname}:{site_path}': {e}")
        return []

    # 2. Resolve the root folder ID.
    try:
        folder_resp = _r.get(
            f"{_GRAPH}/sites/{site_id}/drive/root:/{folder}",
            headers=headers, timeout=20,
        )
        folder_resp.raise_for_status()
        root_folder_id = folder_resp.json()["id"]
    except Exception as e:
        logger.warning(f"JD folder scan: cannot resolve folder path '{folder}': {e}")
        return []

    # 3. Recursively find all PDF/DOCX files in the folder and its subfolders.
    queue = [(root_folder_id, "")]
    items_to_parse = []

    while queue:
        curr_id, rel_path = queue.pop(0)
        try:
            list_resp = _r.get(
                f"{_GRAPH}/sites/{site_id}/drive/items/{curr_id}/children",
                headers=headers, timeout=20,
            )
            list_resp.raise_for_status()
            children = list_resp.json().get("value", [])
        except Exception as e:
            logger.warning(f"JD folder scan: cannot list contents of '{rel_path or folder}': {e}")
            continue

        for child in children:
            name = child.get("name", "")
            child_id = child.get("id", "")
            if "folder" in child:
                child_rel = f"{rel_path}/{name}" if rel_path else name
                queue.append((child_id, child_rel))
            elif name.lower().endswith((".pdf", ".docx")):
                items_to_parse.append((child, rel_path))

    logger.info(f"JD folder scan: found {len(items_to_parse)} file(s) recursively in '{folder}'")

    # 4. Download and parse each PDF/DOCX.
    base_url = f"https://{hostname}{site_path}"
    catalog = []
    for item, rel_path in items_to_parse:
        name = item.get("name", "")
        item_id = item.get("id", "")
        full_subfolder = f"{folder}/{rel_path}" if rel_path else folder
        web_url = item.get("webUrl") or f"{base_url}/{full_subfolder.strip('/')}/{name}"
        try:
            raw = _r.get(
                f"{_GRAPH}/sites/{site_id}/drive/items/{item_id}/content",
                headers=headers, timeout=30,
            ).content
        except Exception as e:
            logger.warning(f"JD folder scan: download failed for '{name}': {e}")
            catalog.append({
                "name": name,
                "title": name.rsplit(".", 1)[0].replace("_", " ").replace("-", " ").strip(),
                "folder": full_subfolder,
                "url": web_url,
                "file_type": name.rsplit(".", 1)[-1].upper(),
                "status": f"Download failed: {e}",
                "skills": [],
                "role": None,
            })
            continue

        text = extract_text_from_bytes(raw, name)
        title = name.rsplit(".", 1)[0].replace("_", " ").replace("-", " ").strip()
        title = _normalized_file_jd_title(title, text)
        role = parse_jd_text(title, text)
        catalog.append({
            "name": name,
            "title": (role or {}).get("title", title),
            "folder": full_subfolder,
            "url": web_url,
            "file_type": name.rsplit(".", 1)[-1].upper(),
            "status": "Ready" if role else "No recognizable skills",
            "skills": (role or {}).get("skills", []),
            "role": role,
        })

    return catalog


def fetch_jd_folder_from_sharepoint(hostname: str, site_path: str, folder: str) -> list:
    """Download every PDF/DOCX in a SharePoint/OneDrive folder and parse them as JDs."""
    catalog = fetch_jd_folder_catalog(hostname, site_path, folder)
    return [item["role"] for item in catalog if item.get("role")]


def _normalized_file_jd_title(default_title: str, text: str) -> str:
    """Replace administrative filenames with the actual position named in the JD."""
    if not re.search(r"(?i)offer\s+letter\s+info", default_title or ""):
        return default_title
    explicit = re.search(
        r"(?im)^\s*(?:position|role)(?:\s+title)?\s*[:\-]\s*(.{3,120})$",
        text or "",
    )
    if explicit:
        return re.sub(r"\s+", " ", explicit.group(1)).strip(" ,-:")[:120]
    seeking = re.search(
        r"(?is)\bDriverAI\s+is\s+seeking\s+(?:an?\s+)?(.{3,120}?)"
        r"(?:,\s+to\b|\s+to\b|\.\s)",
        text or "",
    )
    if seeking:
        return re.sub(r"\s+", " ", seeking.group(1)).strip(" ,-:")[:120]
    return default_title


# ── Public-URL JD fetcher ────────────────────────────────────────────────────

def _title_from_jd(html: str, url: str) -> str:
    """Best-effort role title: the page <title>, else the last URL path segment."""
    m = re.search(r"(?is)<title>\s*(.*?)\s*</title>", html or "")
    if m:
        title = re.sub(r"\s+", " ", unescape(m.group(1))).strip()
        if title:
            return title[:80]
    slug = url.rstrip("/").split("/")[-1] or url
    return slug.replace("-", " ").replace("_", " ").strip()[:80] or url


def fetch_jd_role(url: str) -> dict | None:
    """Fetch one public JD URL and turn it into a role dict {title, skills, url}."""
    try:
        import requests
        resp = requests.get(url, timeout=20, headers={"User-Agent": "DriverAI-HiringAgent"})
        resp.raise_for_status()
    except Exception as e:
        logger.warning(f"   JD fetch failed for {url}: {e}")
        return None
    text = html_to_text(resp.text)
    skills = _scan_skill_keywords(text)
    if not skills:
        logger.warning(f"   JD had no recognizable skills: {url}")
        return None
    role = stamp_role_identity(
        {"title": _title_from_jd(resp.text, url), "skills": skills, "url": url})
    logger.info(f"   JD loaded: {role['title']} <- {', '.join(skills[:10])}")
    return role


# Section headings that begin the ABOUT-THE-COMPANY blurb, and the ones that end it by
# starting the actual role content. Used by _strip_company_boilerplate below.
_COMPANY_SECTION_START = re.compile(
    r"^\s*(?:about\s+(?:driverai|us|the\s+company)|company\s+(?:overview|profile|background))\s*:?\s*$",
    re.IGNORECASE | re.MULTILINE,
)
_ROLE_SECTION_START = re.compile(
    r"^\s*(?:about\s+the\s+(?:role|position|job|opportunity)|(?:position|role|job)\s+(?:summary|overview|description)"
    r"|what\s+you\s+(?:will|'ll)\s+do|responsibilities|key\s+responsibilities|duties"
    r"|qualifications|requirements|who\s+you\s+are|what\s+we\s+(?:are\s+)?look)",
    re.IGNORECASE | re.MULTILINE,
)


def _strip_company_boilerplate(text: str) -> str:
    """Drop the 'About DriverAI' blurb from a JD before its skills are scanned.

    Added 2026-08-21. That paragraph describes the COMPANY, not the role - DriverAI's runs
    "...combine artificial intelligence, computer vision, precision navigation, cloud
    platforms..." - so every JD carrying it inherited AI/CV skills no matter what the job
    actually was. It put 'computer vision' on 11 postings across Marketing, Executive,
    Finance and Business Analytics, which both advertised skills those roles never wanted
    and let unrelated candidates score against them.

    The proof it was boilerplate and not the role: for one opening, the Job Announcement
    (which carries the blurb) yielded 'computer vision' while the Position Description for
    that same opening (which does not) did not.

    Only the span between an about-the-COMPANY heading and the next role heading is removed;
    if either marker is missing the text is returned untouched, so a JD written without
    those headings is never truncated.
    """
    if not text:
        return text
    out, cursor = [], 0
    for start in _COMPANY_SECTION_START.finditer(text):
        if start.start() < cursor:
            continue
        end_match = _ROLE_SECTION_START.search(text, start.end())
        if not end_match:
            continue                      # no role heading after it - leave the text alone
        out.append(text[cursor:start.start()])
        cursor = end_match.start()
    out.append(text[cursor:])
    return "".join(out)


def parse_jd_text(title: str, text: str) -> dict | None:
    """Turn a pasted (or extracted) JD description into a role dict {title, skills}."""
    title = _normalized_file_jd_title(title, text)
    # Title resolution above still sees the FULL text (the offer-letter/normalisation rules
    # read the opening paragraphs); only the skill scan works on the de-boilerplated copy.
    skills = _scan_skill_keywords(_strip_company_boilerplate(text))
    if not skills:
        logger.warning(f"   JD text '{title}' had no recognizable skills.")
        return None
    role = stamp_role_identity({"title": title or "Untitled JD", "skills": skills})
    logger.info(f"   JD parsed: {role['title']} <- {', '.join(skills[:10])} "
                f"[opening: {role['opening_id']}"
                f"{', ' + role['doc_type'] if role['doc_type'] else ''}]")
    return role


# ── Active role resolver ──────────────────────────────────────────────────────

def _fetch_active_roles_live(cfg: dict) -> list:
    """Fetch + parse every configured JD source from scratch (the slow ~45s path).

    Three JD source types are checked in order and combined:
      1. sharepoint_jd_folder  — scans the configured SP/OneDrive folder
      2. urls                  — fetches public web pages
      3. descriptions          — uses pasted JD text from jd_sources.json
    """
    roles = []

    # 1. SharePoint / OneDrive folder
    sp_folder = cfg["sharepoint_jd_folder"]
    if sp_folder:
        from hiring_agent.config import SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH
        hostname  = cfg["sharepoint_jd_hostname"] or SHAREPOINT_HOSTNAME
        site_path = cfg["sharepoint_jd_site"]     or SHAREPOINT_SITE_PATH
        folder_roles = fetch_jd_folder_from_sharepoint(hostname, site_path, sp_folder)
        # A configured folder that yields nothing is a FAILURE, not an empty folder (see
        # JDFolderScanError). Raise instead of quietly continuing on whatever the pasted
        # descriptions happen to hold - the caller decides whether to keep the last known
        # good cache, but it must never overwrite it with a folder-less rebuild.
        if not folder_roles:
            raise JDFolderScanError(
                f"JD folder '{sp_folder}' is configured but returned no JDs "
                f"(host={hostname!r}, site={site_path!r}). Check the path, the app's "
                f"Files.Read.All consent, and network access - see the warning logged above."
            )
        roles.extend(folder_roles)

    # 2. Public URL JDs
    for url in cfg["urls"]:
        role = fetch_jd_role(url)
        if role:
            roles.append(role)

    # 3. Pasted JD descriptions
    for desc in cfg.get("descriptions", []):
        role = parse_jd_text(desc.get("title", ""), desc.get("text", ""))
        if role:
            roles.append(role)

    return roles


def _source_signature(cfg: dict) -> str:
    """Fingerprint of the JD *configuration* (which sources, not their contents).

    Adding/removing a JD file inside the same folder does NOT change this — the cache is
    rebuilt only when you re-scan via the app (or `bot.py --refresh-jd`), or when the
    configured sources themselves change. That is what keeps scoring runs from silently
    re-downloading every JD each run.
    """
    import hashlib
    payload = json.dumps({
        "enabled": cfg.get("enabled"),
        "folder": cfg.get("sharepoint_jd_folder", ""),
        "host": cfg.get("sharepoint_jd_hostname", ""),
        "site": cfg.get("sharepoint_jd_site", ""),
        "urls": sorted(cfg.get("urls", [])),
        "descriptions": sorted(d.get("text", "") for d in cfg.get("descriptions", [])),
    }, sort_keys=True)
    return hashlib.sha256(payload.encode()).hexdigest()[:16]


def _jd_folder_fingerprint(cfg: dict) -> str | None:
    """Lightweight fingerprint of the SharePoint JD folder's CURRENT contents — the file
    names + lastModified + size of every PDF/DOCX in it. Metadata ONLY (recursive folder
    listing, no file downloads). Returns None when there is no folder source configured or
    it can't be listed (no token / network / permission).

    This complements _source_signature (which fingerprints WHICH sources are configured, not
    their contents): folding a folder-contents fingerprint into the cache lets a JD file
    added/removed/edited IN the folder be detected on the very next run, instead of waiting
    up to 24h for the TTL. None is treated by callers as 'can't tell' -> skip the check (never
    forces a spurious refresh when the folder is momentarily unreachable)."""
    host = (cfg.get("sharepoint_jd_hostname") or "").strip()
    site = (cfg.get("sharepoint_jd_site") or "").strip()
    folder = (cfg.get("sharepoint_jd_folder") or "").strip().strip("/")
    if not (host and site and folder):
        return None
    import hashlib
    import requests as _r
    try:
        headers = {"Authorization": f"Bearer {_graph_token()}"}
        site_resp = _r.get(f"{_GRAPH}/sites/{host}:{site}", headers=headers, timeout=20)
        site_resp.raise_for_status()
        site_id = site_resp.json()["id"]

        folder_resp = _r.get(
            f"{_GRAPH}/sites/{site_id}/drive/root:/{folder}",
            headers=headers, timeout=20,
        )
        folder_resp.raise_for_status()
        root_folder_id = folder_resp.json()["id"]

        queue = [root_folder_id]
        all_metadata = []

        while queue:
            curr_id = queue.pop(0)
            list_resp = _r.get(
                f"{_GRAPH}/sites/{site_id}/drive/items/{curr_id}/children"
                "?$select=id,name,size,lastModifiedDateTime,folder&$top=500",
                headers=headers, timeout=20,
            )
            list_resp.raise_for_status()
            children = list_resp.json().get("value", [])
            for child in children:
                if "folder" in child:
                    queue.append(child.get("id"))
                elif child.get("name", "").lower().endswith((".pdf", ".docx")):
                    all_metadata.append((
                        child.get("name", ""),
                        child.get("lastModifiedDateTime", ""),
                        child.get("size", 0)
                    ))
    except Exception as e:
        logger.debug(f"JD folder fingerprint unavailable: {e}")
        return None
    sig = sorted(all_metadata)
    return hashlib.sha256(json.dumps(sig).encode()).hexdigest()[:16]


def _write_role_cache(roles: list, cfg: dict) -> None:
    """Persist parsed roles + the config signature so the next run reads them instantly.

    Also stores a fingerprint of the JD folder's current contents (see
    _jd_folder_fingerprint) so a file added to the folder is detected on the next run."""
    import datetime
    fingerprint = _jd_folder_fingerprint(cfg)
    # A null fingerprint alongside a configured folder means file-change detection is OFF:
    # get_active_roles skips the comparison when either side is None, so an edited/added JD
    # would sit unnoticed until the 24h TTL. That combination went unnoticed for days in the
    # 2026-08-21 incident (see JDFolderScanError), so name it explicitly instead.
    if cfg.get("sharepoint_jd_folder") and fingerprint is None:
        logger.warning(
            "JD folder fingerprint could not be computed - per-file change detection is "
            "DISABLED for this cache; JD edits will only be picked up by the 24h refresh."
        )
    try:
        JD_CACHE_FILE.write_text(json.dumps({
            "built_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "signature": _source_signature(cfg),
            "folder_fingerprint": fingerprint,
            "roles": roles,
        }, indent=2), encoding="utf-8")
    except OSError as e:
        logger.warning(f"Could not write JD cache {JD_CACHE_FILE.name}: {e}")


def _read_role_cache() -> dict | None:
    if not JD_CACHE_FILE.exists():
        return None
    try:
        cached = json.loads(JD_CACHE_FILE.read_text(encoding="utf-8"))
    except Exception as e:
        logger.warning(f"Could not read JD cache {JD_CACHE_FILE.name}: {e}")
        return None
    # A cache written before 2026-08-03 has no opening identity on its roles. Backfill it on
    # read (rather than forcing a full rescan, which needs Graph credentials and ~45s) so the
    # scorer can rely on the field being present no matter how old the cache on disk is.
    stamp_roles(cached.get("roles") if isinstance(cached, dict) else None)
    return cached


def build_jd_cache(catalog: list | None = None) -> list:
    """Scan every configured JD source ONCE and cache the parsed roles to disk.

    This is the "one-time" refresh the app's Job Descriptions tab (and
    `bot.py --refresh-jd`) call. After it runs, scoring reads the cache instantly instead
    of re-downloading all JDs each run. Pass a `catalog` already fetched by
    fetch_jd_folder_catalog() to reuse the app's scan instead of downloading twice.
    Returns the roles written (built-in roles if JD sources are disabled).
    """
    cfg = load_jd_sources()
    if not cfg["enabled"]:
        try:
            JD_CACHE_FILE.unlink(missing_ok=True)   # disabled -> built-in roles, no cache
        except OSError:
            pass
        return get_open_roles()

    if catalog is not None:
        roles = [item["role"] for item in catalog if item.get("role")]
        for url in cfg["urls"]:                     # folder catalog lacks URL/pasted JDs
            r = fetch_jd_role(url)
            if r:
                roles.append(r)
        for desc in cfg.get("descriptions", []):
            r = parse_jd_text(desc.get("title", ""), desc.get("text", ""))
            if r:
                roles.append(r)
    else:
        try:
            roles = _fetch_active_roles_live(cfg)
        except JDFolderScanError as e:
            # Keep the last known good cache rather than replacing it with a rebuild that
            # never saw the JD folder. Logged at ERROR (not WARNING) because this runs
            # unattended - a warning here is what let the broken path hide for days.
            logger.error(f"JD REFRESH ABORTED: {e}")
            cached = _read_role_cache()
            if cached and cached.get("roles"):
                logger.error(f"   Keeping the existing cache of {len(cached['roles'])} role(s) "
                             f"(built {cached.get('built_at', '?')}) - it was NOT overwritten. "
                             f"Scoring continues on it, but it is now STALE until this is fixed.")
                return cached["roles"]
            logger.error("   No usable cache on disk either - falling back to built-in roles.")
            return get_open_roles()

    if not roles:
        logger.warning("JD refresh found no roles — leaving the built-in roles in effect.")
        return get_open_roles()
    _write_role_cache(roles, cfg)
    logger.info(f"JD cache built: {len(roles)} role(s) -> {JD_CACHE_FILE.name}")
    return roles


def get_active_roles(refresh: bool = False) -> list:
    """The role list the scorer matches against — served from the on-disk cache when valid.

    Priority: on-disk JD cache (instant) → live fetch of JD sources → built-in roles.
    The cache is built once by the app / `bot.py --refresh-jd`; a normal scoring run never
    re-downloads all JDs unless the cache is missing, the JD config changed, or
    refresh=True. On a cache miss it fetches live and writes the cache so the next run is
    instant (in the cloud, commit jd_roles_cache.json so fresh checkouts start warm).
    """
    cfg = load_jd_sources()
    if not cfg["enabled"]:
        return get_open_roles()

    if not refresh:
        cached = _read_role_cache()
        if cached and cached.get("signature") == _source_signature(cfg) and cached.get("roles"):
            import datetime
            try:
                built_at_str = cached.get("built_at")
                if built_at_str:
                    built_at = datetime.datetime.fromisoformat(built_at_str)
                    age = datetime.datetime.now() - built_at
                    if age > datetime.timedelta(hours=24):
                        logger.info(f"JD cache is older than 24 hours ({age.total_seconds() / 3600:.1f} hours old) — triggering automatic background refresh.")
                        refresh = True
            except Exception as ex:
                logger.warning(f"Error checking JD cache age: {ex}")

            # Detect a JD FILE added/removed/edited in the SharePoint folder (the config
            # signature above only covers WHICH sources are configured, not their contents).
            # A metadata-only folder listing is compared to the fingerprint stored in the
            # cache; a mismatch rebuilds immediately instead of waiting for the 24h TTL. If
            # the folder can't be listed (None), the check is skipped — no spurious refresh.
            if not refresh:
                current_fp = _jd_folder_fingerprint(cfg)
                cached_fp = cached.get("folder_fingerprint")
                if current_fp is not None and cached_fp is not None and current_fp != cached_fp:
                    logger.info("JD folder contents changed since last scan "
                                "(file added/removed/edited) — rebuilding role cache now.")
                    refresh = True

            if not refresh:
                logger.info(f"Matching against {len(cached['roles'])} JD role(s) "
                            f"(cached {cached.get('built_at', '?')}; use the app's Refresh or "
                            f"`bot.py --refresh-jd` to rescan).")
                return cached["roles"]

    try:
        roles = _fetch_active_roles_live(cfg)
    except JDFolderScanError as e:
        # Same rule as build_jd_cache: a scoring run must never silently downgrade the
        # cache. Serve the existing roles if there are any, and say loudly that they are
        # stale, rather than scoring this candidate against a folder-less rebuild.
        logger.error(f"JD FOLDER SCAN FAILED: {e}")
        cached = _read_role_cache()
        if cached and cached.get("roles"):
            logger.error(f"   Scoring on the existing STALE cache of {len(cached['roles'])} "
                         f"role(s) (built {cached.get('built_at', '?')}); cache NOT overwritten.")
            return cached["roles"]
        logger.error("   No usable cache on disk either - falling back to built-in roles.")
        return get_open_roles()

    if not roles:
        logger.warning("JD sources enabled but none loaded — falling back to built-in roles.")
        return get_open_roles()
    logger.info(f"Matching against {len(roles)} JD role(s) (freshly fetched).")
    _write_role_cache(roles, cfg)   # warm the cache so the next run is instant
    return roles
