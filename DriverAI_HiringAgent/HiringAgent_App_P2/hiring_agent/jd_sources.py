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
from hiring_agent.scoring import get_open_roles

_GRAPH = "https://graph.microsoft.com/v1.0"


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
    role = {"title": _title_from_jd(resp.text, url), "skills": skills, "url": url}
    logger.info(f"   JD loaded: {role['title']} <- {', '.join(skills[:10])}")
    return role


def parse_jd_text(title: str, text: str) -> dict | None:
    """Turn a pasted (or extracted) JD description into a role dict {title, skills}."""
    title = _normalized_file_jd_title(title, text)
    skills = _scan_skill_keywords(text)
    if not skills:
        logger.warning(f"   JD text '{title}' had no recognizable skills.")
        return None
    role = {"title": title or "Untitled JD", "skills": skills}
    logger.info(f"   JD parsed: {role['title']} <- {', '.join(skills[:10])}")
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
        roles.extend(fetch_jd_folder_from_sharepoint(hostname, site_path, sp_folder))

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
    try:
        JD_CACHE_FILE.write_text(json.dumps({
            "built_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "signature": _source_signature(cfg),
            "folder_fingerprint": _jd_folder_fingerprint(cfg),
            "roles": roles,
        }, indent=2), encoding="utf-8")
    except OSError as e:
        logger.warning(f"Could not write JD cache {JD_CACHE_FILE.name}: {e}")


def _read_role_cache() -> dict | None:
    if not JD_CACHE_FILE.exists():
        return None
    try:
        return json.loads(JD_CACHE_FILE.read_text(encoding="utf-8"))
    except Exception as e:
        logger.warning(f"Could not read JD cache {JD_CACHE_FILE.name}: {e}")
        return None


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
        roles = _fetch_active_roles_live(cfg)

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

    roles = _fetch_active_roles_live(cfg)
    if not roles:
        logger.warning("JD sources enabled but none loaded — falling back to built-in roles.")
        return get_open_roles()
    logger.info(f"Matching against {len(roles)} JD role(s) (freshly fetched).")
    _write_role_cache(roles, cfg)   # warm the cache so the next run is instant
    return roles
