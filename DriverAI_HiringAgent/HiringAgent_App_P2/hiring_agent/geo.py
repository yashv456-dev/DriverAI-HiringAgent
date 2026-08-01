"""Location / geo-filter logic (USA-only hiring).

Classifies candidates as confirmed USA, confirmed non-USA, or unknown/conflicting.
Offline rules run first (state abbrevs, full state names, territories, ZIP, known US
cities), followed by optional online Nominatim confirmation for ambiguous cases.
Unknown evidence is never treated as a confirmed residence or an automatic rejection.
"""

import re as _re
import time as _time
from enum import Enum

from hiring_agent.config import (
    KNOWN_US_CITIES, FOREIGN_COUNTRIES, FOREIGN_CITIES, FOREIGN_REGIONS, FOREIGN_DEMONYMS,
    US_STATE_ABBREVS, US_STATE_NAMES, US_TERRITORY_NAMES, US_COUNTRY_TERMS,
    logger,
)


# Any of these (case/punctuation-insensitive) is the United States — canonicalize to one label.
_US_CANONICAL = "United States"


class GeoDecision(str, Enum):
    """Location decision used by P2's scoring and recovery workflows."""

    CONFIRMED_US = "confirmed_us"
    CONFIRMED_NON_US = "confirmed_non_us"
    UNKNOWN = "unknown"


def _country_key(s: str) -> str:
    """Fold a country string to a comparison key: lowercase, drop dots/hyphens, collapse spaces."""
    k = str(s or "").lower().replace(".", " ").replace("-", " ")
    return " ".join(k.split())


# US_COUNTRY_TERMS ('usa', 'u.s.', ...) plus common extras, all folded to the same key shape.
_US_KEYS = {_country_key(t) for t in US_COUNTRY_TERMS} | {
    _country_key(x) for x in
    ("america", "united states of america", "the united states", "u s a", "u s")
} | {"usa", "us"}

_ADMIN_REGION_TO_COUNTRY = {
    # Indian states / union territories that extraction can occasionally place in Country.
    "andhra pradesh": "India", "arunachal pradesh": "India", "assam": "India",
    "bihar": "India", "chhattisgarh": "India", "goa": "India", "gujarat": "India",
    "haryana": "India", "himachal pradesh": "India", "jharkhand": "India",
    "karnataka": "India", "kerala": "India", "madhya pradesh": "India",
    "maharashtra": "India", "manipur": "India", "meghalaya": "India",
    "mizoram": "India", "nagaland": "India", "odisha": "India",
    "orissa": "India", "punjab": "India", "rajasthan": "India",
    "sikkim": "India", "tamil nadu": "India", "telangana": "India",
    "tripura": "India", "uttar pradesh": "India", "uttarakhand": "India",
    "west bengal": "India", "delhi": "India", "jammu and kashmir": "India",
    "ladakh": "India",
    # Common Canadian provinces.
    "ontario": "Canada", "on": "Canada", "british columbia": "Canada",
    "alberta": "Canada", "quebec": "Canada",
}


def normalize_country(country: str) -> str:
    """Canonicalize any US spelling ('USA', 'U.S.', 'US', 'America', ...) to 'United States'.

    Non-US countries are returned unchanged (case preserved), and gaps/placeholders pass
    through untouched so 'Not extracted' or '' never becomes 'United States'. Deterministic
    and side-effect-free — used everywhere P2 finalizes the Country column so the sheet
    always reads one consistent label.
    """
    raw = str(country or "").strip()
    if not raw:
        return raw
    key = _country_key(raw)
    if key in _US_KEYS or key.replace(" ", "") in {"usa", "us"}:
        return _US_CANONICAL
    if key in _ADMIN_REGION_TO_COUNTRY:
        return _ADMIN_REGION_TO_COUNTRY[key]
    return raw


def reconcile_us_country(country: str, location: str, is_usa: bool) -> str:
    """Finalize the Country label for a candidate we are keeping as US-based.

    When `is_usa` is True AND the location carries a CONCRETE US signal (state / territory /
    ZIP via is_strong_usa, or a known US city via is_us_city), the Country cell must read the
    canonical 'United States' — extraction routinely leaves a US STATE name here (e.g.
    'California' off 'San Francisco, CA') or a FOREIGN country pulled from a past overseas
    degree (e.g. 'India' for someone now at UChicago in 'Chicago, IL'). Returns 'United
    States' in that case, otherwise the country unchanged (so a bare benefit-of-doubt keep,
    with no US location signal, never fabricates 'United States' over a real/unknown value).
    """
    if is_usa and location and (is_strong_usa(location) or is_us_city(location)):
        return _US_CANONICAL
    return country


def _word(term: str, loc: str) -> bool:
    """Whole-word (boundary) match of `term` inside lowercased `loc`."""
    return bool(_re.search(r"(?<![a-z])" + _re.escape(term) + r"(?![a-z])", loc))


def is_strong_usa(location: str) -> bool:
    """STRONG USA signal: country term, state (abbrev/name), territory, or US ZIP.

    Deliberately excludes bare city names (a weak signal) so a clear FOREIGN country
    can override an ambiguous US-city substring (e.g. 'Ontario, Canada'). See is_us_city.
    """
    raw = (location or "").strip()
    loc = raw.lower()
    if not loc or loc == "not extracted":
        return False

    for term in US_COUNTRY_TERMS:
        if term in loc:
            return True

    single_admin = loc.strip().rstrip(".")
    if single_admin in US_STATE_ABBREVS or single_admin in US_STATE_NAMES or single_admin in US_TERRITORY_NAMES:
        return True

    # County + state (e.g. "Cook County, IL" or "Cook County, Illinois") - a bare county
    # name with no state at all still falls through to the online/benefit-of-doubt path.
    if "county" in loc:
        for abbr in US_STATE_ABBREVS:
            if _word(abbr, loc):
                return True
        for name in US_STATE_NAMES:
            if _word(name, loc):
                return True

    m = _re.search(r",\s*([A-Za-z])\.?([A-Za-z])\.?\s*(\d{5})?\s*$", raw)
    if m and (m.group(1) + m.group(2)).lower() in US_STATE_ABBREVS:
        return True

    tail = _re.split(r",\s*", loc)
    if len(tail) >= 2:
        last_seg = tail[-1].strip().rstrip(".")
        if (last_seg in US_STATE_NAMES or last_seg in US_TERRITORY_NAMES
                or last_seg in US_STATE_ABBREVS or last_seg == "us"):
            return True

    if _re.search(r"\b\d{5}(?:-\d{4})?\b", loc):
        return True

    for terr in US_TERRITORY_NAMES:
        if _word(terr, loc):
            return True

    return False


def _has_us_admin_signal(location: str) -> bool:
    """Concrete US-place evidence excluding bare country words.

    Used when a location also contains a foreign city/country. A string like
    'Mumbai, United States' should not be kept just because it contains the words
    United States; a concrete state/territory/ZIP signal like 'Paris, Texas' can keep.
    """
    raw = (location or "").strip()
    loc = raw.lower()
    if not loc or loc == "not extracted":
        return False

    single_admin = loc.strip().rstrip(".")
    if single_admin in US_STATE_ABBREVS or single_admin in US_STATE_NAMES or single_admin in US_TERRITORY_NAMES:
        return True

    if "county" in loc:
        for abbr in US_STATE_ABBREVS:
            if _word(abbr, loc):
                return True
        for name in US_STATE_NAMES:
            if _word(name, loc):
                return True

    m = _re.search(r",\s*([A-Za-z])\.?([A-Za-z])\.?\s*(\d{5})?\s*$", raw)
    if m and (m.group(1) + m.group(2)).lower() in US_STATE_ABBREVS:
        return True

    tail = _re.split(r",\s*", loc)
    if len(tail) >= 2:
        last_seg = tail[-1].strip().rstrip(".")
        if (last_seg in US_STATE_NAMES or last_seg in US_TERRITORY_NAMES
                or last_seg in US_STATE_ABBREVS):
            return True

    if _re.search(r"\b\d{5}(?:-\d{4})?\b", loc):
        return True

    for terr in US_TERRITORY_NAMES:
        if _word(terr, loc):
            return True

    return False


def is_us_city(location: str) -> bool:
    """WEAK USA signal: a bare unambiguous US city name (no state needed)."""
    loc = (location or "").strip().lower()
    if not loc or loc == "not extracted":
        return False
    # Resume writers commonly spell Washington, DC as "Washington D.C.".
    # Normalize punctuation for city matching while preserving the original
    # boundary-aware match for every other configured city.
    normalized = _re.sub(r"[^a-z0-9]+", " ", loc).strip()
    normalized = _re.sub(r"\bd\s+c\b", "dc", normalized)
    return any(
        _word(city, loc)
        or _word(_re.sub(r"[^a-z0-9]+", " ", city).strip(), normalized)
        for city in KNOWN_US_CITIES
    )


# Backwards-compatible alias: any USA signal (strong OR a known US city).
def is_usa_location(location: str) -> bool:
    return is_strong_usa(location) or is_us_city(location)


def is_foreign_location(location: str) -> bool:
    """True if the location matches a non-US country, region/bloc, demonym, or foreign city."""
    loc = (location or "").strip().lower()
    if not loc or loc == "not extracted":
        return False
    for country in FOREIGN_COUNTRIES:
        if _word(country, loc):
            return True
    for region in FOREIGN_REGIONS:
        if _word(region, loc):
            return True
    for dem in FOREIGN_DEMONYMS:
        if _word(dem, loc):
            return True
    for city in FOREIGN_CITIES:
        if _word(city, loc):
            return True
    return False


def verify_location_online(location: str) -> bool | None:
    """Query OpenStreetMap Nominatim to confirm whether `location` is in the USA.

    Checks the top 3 candidates (not just the single best match) and returns True if
    ANY of them resolves to the US - closes the "wrong top-ranked match" failure mode,
    where a same-named foreign town outranks the real US match and would otherwise
    silently cause a false reject. Returns True (USA), False (non-USA, none of the top
    3 matched), or None (lookup failed / no results). Free, no API key, rate-limited to
    1 req/sec by the caller.
    """
    loc = (location or "").strip()
    if not loc or loc.lower() == "not extracted":
        return None
    try:
        import requests as _requests
        resp = _requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={"q": loc, "format": "json", "limit": 3, "addressdetails": 1},
            headers={"User-Agent": "DriverAI-HiringAgent/1.0"},
            timeout=10,
        )
        resp.raise_for_status()
        results = resp.json()
        if not results:
            return None
        for r in results:
            cc = r.get("address", {}).get("country_code", "").lower()
            if cc == "us":
                logger.debug(f"   online match for '{loc}': {r.get('display_name', '')} (US)")
                return True
        top = results[0]
        logger.debug(f"   online: '{loc}' resolved to {top.get('display_name', '')} "
                     f"(country_code={top.get('address', {}).get('country_code', '?')}) - "
                     f"none of top {len(results)} matched US")
        return False
    except Exception as e:
        logger.debug(f"   online location lookup failed for '{loc}': {e}")
        return None
    finally:
        _time.sleep(1.1)  # Nominatim public API: max 1 req/sec


_USA_CONTEXT_PROMPT = (
    "Look ONLY at this resume's university/education names and employer/work-history "
    "names below - ignore any stated home address or contact location entirely. Based "
    "only on those institution and company names, does this person's education or work "
    "history read as US-based? Answer true only if you recognize specific US "
    "universities or US-headquartered employers in the text. Answer false only if you "
    "recognize specific non-US institutions or employers. Answer null if the text gives "
    "no clear signal either way (schools/employers you don't recognize, or none "
    "mentioned) - never force a guess. "
    "Return ONLY a JSON object: {\"usa\": true, false, or null}"
)


def verify_usa_from_context(resume_text: str) -> bool | None:
    """Second opinion for an ambiguous location: does the resume's education/employment
    history read as US-based? A signal Nominatim can't provide at all - it only knows
    geography, not which universities/employers are American.

    Same class of task as infer_looking_for_role/infer_missing_portfolios: a narrow,
    single-purpose local Ollama call, gated on OLLAMA_ENABLED, wrapped in try/except,
    never raises. Returns True/False only when the model recognizes a specific
    institution/employer; None otherwise (Ollama off/unreachable, or the resume gives no
    signal) - callers must treat None as "no corroboration available", not as a foreign
    signal. Honest limitation: a local model like llama3.2 will recognize well-known
    universities/employers easily but may miss an obscure regional one - a corroborating
    signal, not an infallible verifier.
    """
    from hiring_agent.config import OLLAMA_ENABLED, OLLAMA_MODEL, OLLAMA_HOST, OLLAMA_TIMEOUT, AI_TEXT_LIMIT
    if not OLLAMA_ENABLED or not (resume_text or "").strip():
        return None
    try:
        import json as _json
        import requests as _requests
        resp = _requests.post(
            f"{OLLAMA_HOST}/api/chat",
            json={
                "model": OLLAMA_MODEL,
                "format": "json",
                "stream": False,
                "options": {"temperature": 0},
                "messages": [
                    {"role": "system", "content": _USA_CONTEXT_PROMPT},
                    {"role": "user", "content": resume_text[:AI_TEXT_LIMIT]},
                ],
            },
            timeout=OLLAMA_TIMEOUT,
        )
        resp.raise_for_status()
        data = _json.loads(resp.json()["message"]["content"])
        usa = data.get("usa", None)
        return usa if isinstance(usa, bool) else None
    except Exception as e:
        logger.debug(f"   resume-context USA check unavailable: {e}")
        return None


#: Canadian NANP area codes. The North American Numbering Plan covers both the US and
#: Canada with the same +1/10-digit shape, so a NANP-shaped number alone can't tell them
#: apart - excluding known-Canadian area codes here catches the common case. Added
#: 2026-07-24 for the Mukesh Kumar Sharma case: 'London, ON' + Country='Canada' + phone
#: '+1 519-859-0693' (519 = Ontario) was reading as a "US phone" and contradicting a
#: correctly-foreign geo verdict. List is stable (new area codes are added by overlay,
#: existing ones don't change), so a static set is fine to maintain by hand.
_CANADIAN_AREA_CODES = {
    "204", "226", "236", "249", "250", "289", "306", "343", "354", "365", "367", "368",
    "382", "403", "416", "418", "428", "431", "437", "438", "450", "468", "474", "506",
    "514", "519", "548", "579", "581", "584", "587", "604", "613", "639", "647", "672",
    "683", "705", "709", "742", "753", "778", "780", "782", "807", "819", "825", "867",
    "873", "879", "902", "905",
}


def _looks_like_us_phone(phone: str) -> bool:
    """True if the phone reads as a US (NANP) number: a bare 10-digit number, or 11 digits
    starting with country code 1 — never an international +NN number (e.g. +91 India), and
    never a NANP number whose area code is Canadian (see _CANADIAN_AREA_CODES). Mirrors
    extraction.format_phone's US-number rule. A US phone is a moderately strong 'this person is
    in the US' signal, used only to CONTRADICT a mis-extracted foreign country/location (a
    résumé's hometown or past study) so a US-based applicant is not wrongly rejected. Added
    2026-07-12 for the Manasa Puli case: current job 'Georgia (Remote)', US phone 865-…, but the
    parser read her hometown 'Hyderabad'/'India' and rejected her."""
    raw = str(phone or "").strip()
    if not raw or "*" in raw:                       # blank or deliberately masked
        return False
    d = "".join(c for c in raw if c.isdigit())
    if len(d) == 11 and d.startswith("1"):
        return d[1:4] not in _CANADIAN_AREA_CODES
    if len(d) == 10 and not raw.lstrip().startswith("+"):
        # A bare, unformatted 10-digit string (no separators at all) is indistinguishable
        # from a foreign mobile number written without its country code - fixed 2026-08-01,
        # live case (Divy Parmar / APP-20260720-1013-09AC): his real Indian number
        # "7405465204", written with zero separators, coincidentally matched the NANP
        # area-code shape and was trusted as "looks like a US phone" strongly enough to
        # override his resume's own explicit "Khokhra, Ahmedabad" location, leaving a
        # clearly non-US candidate sitting in Needs Review instead of Rejected. Every
        # protected US-phone regression case already carries a visual separator (dash,
        # space, dot, or parens) or a '+' country code, so requiring one here costs nothing.
        if not any(c in raw for c in " -.()"):
            return False
        # NANP area code and exchange code cannot start with 0 or 1, and N11 (911, etc.) is not an area code
        if d[1:3] == "11":
            return False
        if d[0] not in "01" and d[3] not in "01":
            return d[0:3] not in _CANADIAN_AREA_CODES
        return False
    return False


def _education_reads_usa(text: str) -> bool:
    """Loose scan for a US state/country/territory name inside free-form Education (or
    resume-context) text, e.g. 'MS Computer Science, Arizona State University'.

    Unlike is_strong_usa (built for a location-shaped 'City, ST' string), this text isn't
    shaped like an address, so it deliberately checks full names as word-boundary
    substrings anywhere in the text. It deliberately does NOT check bare 2-letter state
    abbreviations here (unlike is_strong_usa's tightly-scoped abbreviation checks) - across
    free-form prose, 'IN'/'OR'/'ME'/'HI' etc. collide constantly with common English words
    ('in', 'or', 'me', 'hi'), which would make this fire on nearly any resume.
    """
    t = (text or "").strip().lower()
    if not t or t == "not extracted":
        return False
    for term in US_COUNTRY_TERMS:
        if term in t:
            return True
    for name in US_STATE_NAMES:
        if _word(name, t):
            return True
    for terr in US_TERRITORY_NAMES:
        if _word(terr, t):
            return True
    return False


def _corroborates_non_usa(phone: str = "", education: str = "", resume_text: str = "") -> str:
    """Does Phone/Education corroborate — or contradict — a foreign Location match?

    Returns "contradicts_usa" if either signal reads as US-based (a US-formatted phone,
    or an Education entry with a US state/territory/country term per _education_reads_usa,
    or - when Education itself is blank - the resume's education/employment context per
    verify_usa_from_context reading USA), "no_signal" if neither Phone nor Education
    carries any signal at all (nothing to
    corroborate a reject with), else "corroborates" (at least one signal present and
    neither contradicts). Added 2026-07-24: a Location match alone (which extraction can
    pick up from anywhere in the resume text, including a stale hometown/old-degree
    mention) is no longer sufficient proof of current whereabouts — see check_location_usa
    step 2.
    """
    phone_str = str(phone or "").strip()
    phone_present = bool(phone_str) and phone_str.lower() not in ("not extracted", "n/a")
    phone_contradicts = phone_present and _looks_like_us_phone(phone_str)

    edu_str = str(education or "").strip()
    edu_present = bool(edu_str) and edu_str.lower() not in ("not extracted", "n/a")
    edu_contradicts = False
    if edu_present:
        edu_contradicts = _education_reads_usa(edu_str)
    else:
        context = verify_usa_from_context(resume_text) if resume_text else None
        if context is not None:
            edu_present = True
            edu_contradicts = context is True

    if phone_contradicts or edu_contradicts:
        return "contradicts_usa"
    if not phone_present and not edu_present:
        return "no_signal"
    return "corroborates"


def classify_location_usa(location: str, country: str = "", resume_text: str = "",
                          phone: str = "", education: str = "") -> tuple[GeoDecision, str]:
    """Classify current-location evidence without turning uncertainty into a verdict.

    Confirmed US evidence may be scored, confirmed non-US evidence may be rejected,
    and unknown/conflicting evidence must remain on the main sheet for clarification.
    A US phone, university, or employer is supporting context only: it can prevent an
    unsafe rejection, but it does not independently prove current US residence.

    `resume_text`, if given, lets step 4 (below) corroborate an ambiguous/failed online
    lookup against the candidate's education/work history before ever rejecting them -
    see verify_usa_from_context. Optional and defaults to empty so any caller without
    resume text in scope keeps working unchanged.

    `education` and `phone`, if given, are also required to corroborate a clear FOREIGN
    location match (step 2) before it is trusted enough to reject — see
    _corroborates_non_usa. Both optional and default to empty so any caller without that
    data in scope keeps working unchanged (a foreign match with neither signal available
    is treated as insufficient evidence to reject, not as proof of the reject).
    """
    loc = (location or "").strip()
    loc_present = bool(loc) and loc.lower() != "not extracted"

    # 1) STRONG US location (state / territory / ZIP / 'United States') wins over EVERYTHING,
    #    including a contradicting Country field. Checked FIRST (moved above the country field
    #    on 2026-07-12) because the Country field is often a MIS-EXTRACTION: a parser reads
    #    'India' off a candidate's PAST foreign degree or a finished overseas internship even
    #    though they currently live/study in the US. Physical-presence evidence (a US state/ZIP
    #    in the location) must beat that. Real case: Sadaf Khan — current UChicago master's in
    #    'Chicago, IL', earlier bachelor's in 'Hyderabad, IND' -> Country was parsed as 'India'
    #    and the old country-first order wrongly rejected a US-based candidate (and would have
    #    emailed her a decline). So 'Paris, Texas' and 'City, CA 90210' are kept too.
    if loc_present and is_strong_usa(loc):
        return (
            GeoDecision.CONFIRMED_US,
            f"offline: '{loc}' matched as USA (state/territory/ZIP) - overrides country field",
        )

    # 2) A clear FOREIGN current-location signal is necessary but no longer sufficient on
    #    its own to reject — it must also be corroborated by Phone or Education (see
    #    _corroborates_non_usa) before being trusted enough to reject. Location extraction
    #    can pick up a stale signal (an old degree, a hometown mention, a past internship)
    #    from anywhere in the resume text when it can't find a clean current-address line,
    #    so a location-text match alone is no longer treated as proof of current whereabouts.
    #    Changed 2026-07-24: previously this rejected immediately with zero cross-check,
    #    even though the corroboration primitives (_looks_like_us_phone,
    #    verify_usa_from_context) already existed and were used elsewhere (steps 4 and 6).
    #
    #    - If Phone reads as a US number, or Education/resume context reads as US-based,
    #      that CONTRADICTS the location match -> keep for clarification instead of reject.
    #    - If neither Phone nor Education carries any signal at all, there is not enough
    #      evidence to reject either -> keep (benefit of the doubt).
    #    - Only rejected when the foreign location match has at least one non-contradicting
    #      corroborating signal from Phone or Education.
    if loc_present and is_foreign_location(loc):
        corroboration = _corroborates_non_usa(phone, education, resume_text)
        if corroboration == "contradicts_usa":
            return (
                GeoDecision.UNKNOWN,
                f"offline: '{loc}' matched as non-USA, but phone/education "
                f"indicates USA - current location needs clarification",
            )
        if corroboration == "no_signal":
            return (
                GeoDecision.UNKNOWN,
                f"offline: '{loc}' matched as non-USA, but no independent signal "
                f"corroborates current residence",
            )
        return GeoDecision.CONFIRMED_NON_US, f"offline: '{loc}' matched as non-USA"

    # 3) Bare unambiguous US city (weak signal) -> keep, e.g. 'Austin'. This still runs before
    #    Country so a current US city can override a foreign country pulled from old education.
    if loc_present and is_us_city(loc):
        return GeoDecision.CONFIRMED_US, f"offline: '{loc}' matched as USA (city)"

    # 4) Explicit country field. US country keeps only when the location is blank or ambiguous;
    #    clear foreign locations have already been rejected above. Foreign country rejects
    #    unless current US location was already found above.
    country_low = (country or "").strip().lower()
    if country_low and country_low != "not extracted":
        for term in US_COUNTRY_TERMS:
            if term == country_low:
                return (
                    GeoDecision.UNKNOWN,
                    f"country field '{country}' reads USA, but no confirmed current "
                    f"city/state was extracted",
                )
        for fc in FOREIGN_COUNTRIES:
            if _word(fc, country_low):
                return (
                    GeoDecision.UNKNOWN,
                    f"country field '{country}' reads non-USA, but no confirmed current "
                    f"location was extracted",
                )

    if not loc_present:
        if country_low:
            return GeoDecision.UNKNOWN, f"country '{country}' unresolved as current residence"
        return GeoDecision.UNKNOWN, "no current location extracted"

    # 5) Ambiguous -> ask Nominatim (checks the top 3 candidates, see verify_location_online).
    online = verify_location_online(loc)
    if online is True:
        return GeoDecision.CONFIRMED_US, f"online: '{loc}' confirmed USA"

    # Nominatim didn't confirm USA (False or inconclusive) - get a second, independent
    # opinion from the resume's substance before trusting one geocoding guess enough to
    # reject someone.
    context = verify_usa_from_context(resume_text) if resume_text else None
    if context is True:
        return (
            GeoDecision.UNKNOWN,
            "resume education/employment suggests USA, but current location remains unconfirmed",
        )
    if online is False:
        return (
            GeoDecision.UNKNOWN,
            f"online: '{loc}' resolved outside the USA, but the extracted text was not "
            f"confirmed as the candidate's current location",
        )

    return GeoDecision.UNKNOWN, f"'{loc}' unresolved as current residence"


def check_location_usa(location: str, country: str = "", resume_text: str = "",
                       phone: str = "", education: str = "") -> tuple[bool, str]:
    """Backward-compatible keep/reject view of :func:`classify_location_usa`.

    Existing read-only audits and local helpers expect a boolean. Unknown candidates
    return ``True`` here so they are never rejected by legacy callers; live scoring and
    recovery use the tri-state classifier directly and keep them in location review.
    """
    decision, reason = classify_location_usa(
        location, country=country, resume_text=resume_text,
        phone=phone, education=education,
    )
    return decision != GeoDecision.CONFIRMED_NON_US, reason
