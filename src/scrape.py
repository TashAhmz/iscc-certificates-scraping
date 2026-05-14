import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
import re
from mappings import *
from thefuzz import fuzz, process
import numpy as np
from collections import defaultdict
import unicodedata
import json
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# URLs
API_URL = "https://iscc-system.org/wp-json/api/certificates"
LIST_PAGE = "https://iscc-system.org/certification/all-certificates/"

# GSTs of Geo filepath
GST_GEO = pd.read_excel("C:/Users/tashif.ahmed/OneDrive - Shell/T&S LCF - Analytics, Digital, and Economics - Shared Documents/00. LCF Data Lakehouse/GSTs/GST Geographies/LCF GST of Geographies.xlsx", sheet_name="GS_LCF_Geographies")

# GSTs of Assets filepath
GST_ASSETS = pd.read_excel(r"C:/Users/tashif.ahmed/OneDrive - Shell/T&S LCF - Analytics, Digital, and Economics - Shared Documents/00. LCF Data Lakehouse/GSTs/GST Assets/00. Golden Source File of Asset Capacities.xlsm", sheet_name="GoldenSource")

countries = set(c.lower() for c in ALL_COUNTRIES)
stopwords = set(s.lower() for s in STOPWORDS)   
city_stopwords = set(s.lower() for s in CITY_STOPWORDS)

# Headers

DEFAULT_HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
    "Origin": "https://iscc-system.org",
    "Referer": LIST_PAGE,
    "User-Agent": "Mozilla/5.0",
}

STATUS_TYPES = ["valid", "suspended", "expired", "terminated", "withdrawn"]


STATUS_PRIORITY = {s: i for i, s in enumerate(["withdrawn", "terminated", "suspended", "expired", "valid"])}


# Column names (from table)
COLUMNS = [
    "cert_status", "cert_number","cert_owner","cert_scope","cert_processingunittype","cert_in_put","cert_add_on",
    "cert_products","cert_valid_from","cert_valid_until","cert_suspended_date",
    "cert_issuer","cert_map","cert_file","cert_audit"
]


def _strip_accents(s: str) -> str:
    s = "" if s is None else str(s)
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(ch)
    )

def _safe(text):
    return "" if text is None else str(text).strip()

def _asset_identifier_join(company_name, city):
    company = _safe(company_name)
    city = _safe(city)
    return f"{company} {city}".strip()

def _normalize_for_match(s: str) -> str:
    s = _safe(s).lower()

    # accents: "mède" -> "mede"
    s = _strip_accents(s)

    # html ampersand
    s = s.replace("&amp;", " and ").replace("&", " and ")

    # punctuation including dash -> space
    s = re.sub(r"[.,;:/\-\(\)\[\]]", " ", s)

    # collapse spaces
    s = re.sub(r"\s+", " ", s).strip()
    return s

def add_asset_identifier_and_match(
    df_iscc: pd.DataFrame,
    gst_df: pd.DataFrame
) -> pd.DataFrame:

    # --- Column checks ---
    required_iscc = {"Company_Name", "City", "Country"}
    missing_iscc = required_iscc - set(df_iscc.columns)
    if missing_iscc:
        raise KeyError(f"Missing ISCC columns: {missing_iscc}")

    GST_ID_COL = "Asset Identifier"
    GST_COUNTRY_COL = "Country"
    GST_TERR_COL = "Territory"

    required_gst = {GST_ID_COL, GST_COUNTRY_COL, GST_TERR_COL}
    missing_gst = required_gst - set(gst_df.columns)
    if missing_gst:
        raise KeyError(f"Missing GST columns: {missing_gst}")

    df_iscc = df_iscc.copy()

    # --- 1) Build ISCC Company_City ---
    df_iscc["Company_City"] = [
        _asset_identifier_join(cn, city)
        for cn, city in zip(df_iscc["Company_Name"], df_iscc["City"])
    ]

    # --- 2) Tokenization helper ---
    def _tokens(s: str) -> set:
        """
        Tokenization for matching:
        - keeps len>=2 tokens (bp, sk, etc.)
        - drops pure numeric tokens
        - collapses dotted abbreviations: 'b.v.' -> 'bv', 's.a.' -> 'sa'
        - drops legal suffixes + stopwords
        """
        s = _normalize_for_match(s)

        # collapse dotted abbreviations: a.s. -> as, s.r.o. -> sro, etc.
        s = re.sub(r"\b(?:[a-z]\.){2,}", lambda m: m.group(0).replace(".", ""), s)

        raw = s.split()
        out = set()
        for t in raw:
            t = t.strip()
            if not t:
                continue
            if len(t) < 2:
                continue
            if t.isdigit():
                continue
            if t in LEGAL_SUFFIXES:
                continue
            if t in stopwords:
                continue
            if t in countries:
                continue
            out.add(t)
        return out

    def _disjoint_tokens(comp_tokens: set, city_tokens: set):
        overlap = comp_tokens & city_tokens
        if overlap:
            comp_tokens = comp_tokens - overlap
            city_tokens = city_tokens - overlap
        return comp_tokens, city_tokens

    def _middle_from_cert_holder(cert_holder: str) -> str:
        """
        "Neste Components B.V, Botlek, Rotterdam, Netherlands" -> "Botlek Rotterdam"
        Returns "" if not parseable.
        """
        if not cert_holder:
            return ""
        parts = [p.strip() for p in str(cert_holder).split(",") if p.strip()]
        if len(parts) < 3:
            return ""
        return " ".join(parts[1:-1]).strip()

    # --- 3) Prepare GST lists (raw + normalized for exact match) ---
    gst_raw_series = gst_df[GST_ID_COL].fillna("").astype(str).str.strip()
    gst_country_series = gst_df[GST_COUNTRY_COL].fillna("").astype(str).str.strip()
    gst_terr_series = gst_df[GST_TERR_COL].fillna("").astype(str).str.strip()

    # Keep only rows with a non-empty Asset Identifier
    keep_mask = gst_raw_series.ne("")
    gst_raw_list = gst_raw_series[keep_mask].tolist()
    gst_country_list = gst_country_series[keep_mask].map(_normalize_for_match).tolist()
    gst_terr_list = gst_terr_series[keep_mask].tolist()

    # Exact match prep
    gst_norm_list = [_normalize_for_match(x) for x in gst_raw_list]
    gst_norm_set = set(gst_norm_list)

    # norm -> raw (first one wins if duplicates)
    norm_to_raw = {}
    for n, r in zip(gst_norm_list, gst_raw_list):
        norm_to_raw.setdefault(n, r)

    # --- 4) Build GST token sets INCLUDING Territory tokens ---
    gst_token_sets = []
    for asset_id, terr in zip(gst_raw_list, gst_terr_list):
        toks = _tokens(asset_id)

        # ✅ add territory tokens
        if terr:
            toks |= _tokens(terr)

        gst_token_sets.append(toks)

    # --- 5) Build inverted index for fast candidate retrieval ---
    token_to_gst_idxs = defaultdict(set)
    for i, toks in enumerate(gst_token_sets):
        for tok in toks:
            token_to_gst_idxs[tok].add(i)

    # --- 6) Country gating index ---
    country_to_idxs = defaultdict(list)
    for i, c in enumerate(gst_country_list):
        if c:
            country_to_idxs[c].append(i)

    # --- 7) Main loop ---
    match_flags = []
    overwritten_asset_ids = []

    for row in df_iscc.itertuples(index=False):
        company_name = getattr(row, "Company_Name", None)
        city_name = getattr(row, "City", None)
        iscc_country = getattr(row, "Country", None)
        cert_holder = getattr(row, "Certificate_Holder", None)

        asset_id = getattr(row, "Company_City", "")
        original_display = asset_id
        norm = _normalize_for_match(asset_id)

        # Empty
        if not norm:
            match_flags.append(0)
            overwritten_asset_ids.append(original_display)
            continue

        # Exact match
        if norm in gst_norm_set:
            match_flags.append(1)
            overwritten_asset_ids.append(norm_to_raw[norm])
            continue

        # --- Build ISCC tokens ---
        comp_tokens = _tokens(company_name)

        # City reference: prefer middle of Certificate_Holder, else fallback to row.City
        mid_city = _safe(city_name)  # ✅ always defined
        holder_mid = _middle_from_cert_holder(_safe(cert_holder))
        if holder_mid:
            mid_city = holder_mid

        city_tokens = _tokens(mid_city)

        # enforce disjointness
        comp_tokens, city_tokens = _disjoint_tokens(comp_tokens, city_tokens)

        # ✅ NO city-only matches: require some company tokens to proceed
        if not comp_tokens:
            match_flags.append(0)
            overwritten_asset_ids.append(original_display)
            continue

        # --- Country gating ---
        iscc_country_norm = _normalize_for_match(iscc_country)

        if iscc_country_norm and iscc_country_norm in country_to_idxs:
            gated_idxs = set(country_to_idxs[iscc_country_norm])
        else:
            gated_idxs = None  # means "no gate"

        # --- Candidate selection using inverted index (fast) ---
        candidate_idxs = set()
        for t in comp_tokens:
            candidate_idxs |= token_to_gst_idxs.get(t, set())

        # optional: city tokens help ranking (but not allowed to match alone)
        for t in city_tokens:
            candidate_idxs |= token_to_gst_idxs.get(t, set())

        # apply country gate if available
        if gated_idxs is not None:
            candidate_idxs &= gated_idxs

        best_idx = None
        best_score = -1

        for i in candidate_idxs:
            gst_toks = gst_token_sets[i]
            comp_hit = len(comp_tokens & gst_toks)
            city_hit = len(city_tokens & gst_toks) if city_tokens else 0

            # ✅ Acceptance rules (no city-only):
            combined_ok = (comp_hit >= 1 and city_hit >= 1)     # company + city
            company_only_ok = (comp_hit >= 2 and city_hit == 0) # only company allowed if strong

            # If there was NO country gate, tighten company-only to reduce false positives
            if gated_idxs is None:
                company_only_ok = False

            if combined_ok or company_only_ok:
                # scoring: prefer combined matches
                score = (100 + comp_hit + city_hit) if combined_ok else (comp_hit + city_hit)

                if score > best_score:
                    best_score = score
                    best_idx = i

        if best_idx is not None:
            match_flags.append(1)
            overwritten_asset_ids.append(gst_raw_list[best_idx])
        else:
            match_flags.append(0)
            overwritten_asset_ids.append(original_display)

    df_iscc["Company_City"] = overwritten_asset_ids
    df_iscc["Match_Found"] = match_flags
    return df_iscc


def _normalize(text: str) -> str:
    """Light normalization + stopword removal to improve fuzzy company matches."""
    if not isinstance(text, str):
        return ""

    text = text.lower()

    removals = [
        " inc", " llc", " l.l.c", " lp", " l.p.", " bv", " b.v.", " ltd",
        " co", "co.", " company", " limited",
        ".", ",", "&", "&amp;", "&amp;amp;",
        "ltd.", "pte.", "gmbh", " plc", " s.p.a"
    ]

    for w in removals:
        text = text.replace(w, " ")

    # ✅ remove legal suffixes as whole tokens only (prevents 'cargo' being mangled by 'ag')
    if LEGAL_SUFFIXES:
        pattern = r"\b(?:%s)\b" % "|".join(map(re.escape, LEGAL_SUFFIXES))
        text = re.sub(pattern, " ", text)

    # collapse whitespace
    text = " ".join(text.split())

    # ✅ remove stopwords at token level
    if stopwords:
        tokens = [t for t in text.split() if t not in stopwords]
        text = " ".join(tokens)
    if countries:
        tokens = [t for t in text.split() if t not in countries]
        text = " ".join(tokens)

    return text


def _build_lookup_exact_columns(gst_df: pd.DataFrame, stopwords: set[str]):
    CP_COL = "Company/Producer"
    CPSN_COL = "Company/Producer Short Name"

    for col in (CP_COL, CPSN_COL):
        if col not in gst_df.columns:
            raise KeyError(f"Column '{col}' not found in GST assets DataFrame.")

    tmp = gst_df[[CP_COL, CPSN_COL]].copy()

    tmp["__norm_cp__"]   = tmp[CP_COL].apply(_normalize)
    tmp["__norm_cpsn__"] = tmp[CPSN_COL].apply(_normalize)

    # Universe: unique + non-empty
    universe = pd.unique(pd.concat([tmp["__norm_cp__"], tmp["__norm_cpsn__"]], ignore_index=True)).tolist()
    universe = [u for u in universe if isinstance(u, str) and u.strip()]

    # Map normalized -> original short name (only if short name is not blank)
    to_short = {}
    for _, r in tmp.iterrows():
        short = r[CPSN_COL]
        short = "" if pd.isna(short) else str(short).strip()
        if not short:
            continue

        if r["__norm_cp__"]:
            to_short[r["__norm_cp__"]] = short
        if r["__norm_cpsn__"]:
            to_short[r["__norm_cpsn__"]] = short

    return universe, to_short


def overwrite_company_with_gst_shortname_exact(
    iscc_df: pd.DataFrame,
    gst_df: pd.DataFrame,
    score_threshold
) -> pd.DataFrame:

    if "Company_Name" not in iscc_df.columns:
        raise KeyError("Expected column 'Company_Name' not found in ISCC DataFrame.")

    universe, to_short = _build_lookup_exact_columns(gst_df, STOPWORDS)

    def _as_is(value):
        return "" if pd.isna(value) else str(value)

    if not universe:
        iscc_df["Company_Name"] = iscc_df["Company_Name"].astype(str)
        return iscc_df

    new_values = []
    for original in iscc_df["Company_Name"]:
        original_safe = _as_is(original)
        norm = _normalize(original_safe)

        if not norm.strip():
            new_values.append(original_safe)
            continue

        match, score = process.extractOne(norm, universe, scorer=fuzz.ratio) if universe else (None, 0)

        if match and score >= score_threshold:
            candidate = to_short.get(match, "")
            candidate = candidate if isinstance(candidate, str) else _as_is(candidate)
            candidate = candidate.strip()

            if candidate:
                new_values.append(candidate)
            else:
                new_values.append(original_safe)
        else:
            new_values.append(original_safe)

    iscc_df["Company_Name"] = new_values
    return iscc_df


# Define a function to determine the facility grouping based on Scope* codes
    # It checks each abbreviation and returns the matching group(s)
def determine_facility_grouping(scope_text):
    if not isinstance(scope_text, str):
        return ""
    abbreviations = [abbr.strip() for abbr in scope_text.split(",")]
    groupings = set()
    for abbr in abbreviations:
        group = FACILITY_GROUPING_MAP.get(abbr)
        if group:
            groupings.add(group)
    return ", ".join(sorted(groupings)) if groupings else "Unclassified"

def get_country_name(c):
    exempt_words = ["of", "the", "and"]
    return " ".join([w.capitalize() if w not in exempt_words else w.lower() for w in c.split()])

# old city function now using one that Tom developed

def every_word_has_digit(tok: str) -> bool:
    words = tok.split()
    return bool(words) and all(any(ch.isdigit() for ch in w) for w in words)

def get_city_name(cert_owner):

    if not isinstance(cert_owner, str) or not cert_owner.strip():
            return None

    exempt_words = ["ltd.", "ltd", "s.i.u",
                    "s.a.", "s.a", "s.r.o.",
                    "s.r.o", "s.i.", "s.i",
                    "s.p.a", "s.p.a.", "s.l.u",
                    "s.l.u", "a.s", "a.s.",
                    "s.l", "s.l.", "inc.", "inc",
                    ". ltd", "-", "oils", "l.p.",
                    "llc", "l.l.c.", "llc.", "lp",
                    "inc..", "city", ".ltd.", "ltd .", "/", "-"]
    
    parts = [p.strip().lower() for p in cert_owner.split(",") if p.strip()][1:-1]

    tokens = [
        tok
        for tok in parts
        if tok
        and tok not in exempt_words
        and not any(w in city_stopwords for w in tok.split())
        and not every_word_has_digit(tok)
    ]
    
    country = get_country_name(cert_owner.split(",")[-1].strip().lower() if parts else "").lower()

    # testing to see if this logic works to remove street names coming into the city column by mistake
    if len(tokens) == 1:
        return " ".join([w for w in tokens[0].split() if not any(ch.isdigit() for ch in w)]).title()
    elif len(tokens) >= 2:
        if country in ("united states", "china", "republic of", "brazil", "indonesia", "australia", "japan", "canada"):
            return " ".join([w for w in tokens[-2].split() if not any(ch.isdigit() for ch in w)]).title()
        else:
            return " ".join([w for w in tokens[-1].split() if not any(ch.isdigit() for ch in w)]).title()
        
    return None

def get_lat_lon(link):
    if not isinstance(link, str) or "maps/place/" not in link:
        return None, None
    coords = link.split("maps/place/")[-1].split("+")
    # Filter out empty strings
    coords = [c.strip() for c in coords if c.strip() and c.strip() != "0.000000"]
    if len(coords) >= 2:
        return coords[0], coords[1]
    else:
        return None, None

def get_latitude(link):
    lat, lon = get_lat_lon(link)
    return lat if lat else "Unknown"

def get_longitude(link):
    lat, lon = get_lat_lon(link)
    return lon if lon else "Unknown"


def map_certificate_type(cert_id):
    try:
        parts = [p.strip() for p in cert_id.split("-")]
        id = " ".join(parts[0:2]).upper()
    except (ValueError, TypeError):
        return ""
    if id == "CORSIA ISCC":
        return "Aviation"
    elif id == "DE B":
        return "Legacy"
    elif id == "EU ISCC":
        return "Mandated"
    else:
        return CERTIFICATE_TYPE_MAP.get(id, "Undefined")

def map_certificate_class(cert_type):
    for key, value in CERTIFICATE_TYPE_MAP.items():
        if value == cert_type:
            return key
    return "Unknown"

def map_region(country):
    r_map = GST_GEO[["Country", "LCF SnD region 2"]]
    r_map_dict = dict(zip(r_map["Country"], r_map["LCF SnD region 2"]))
    return r_map_dict.get(country, "Unknown")

def map_subregion(country):
    r_map = GST_GEO[["Country", "LCF SnD region 1"]]
    r_map_dict = dict(zip(r_map["Country"], r_map["LCF SnD region 1"]))
    return r_map_dict.get(country, "Unknown")

def clean_excel_string(x):
    """
    Cleans strings coming from Excel/HTML/PDF by removing XML-illegal controls,
    normalising whitespace, and stripping invisible characters commonly found
    in certificates and scraped data.
    """
    # XML-disallowed control characters (except \t, \n, \r which we handle explicitly)
    _ILLEGAL_CTRL = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")
    if x is None:
        return ""
    s = str(x)
    s = _ILLEGAL_CTRL.sub("", s)
    s = (
        s.replace("\r", "")           # carriage return
         .replace("\t", " ")          # tabs -> space
         .replace("\u00A0", " ")      # NBSP (unicode)
         .replace("\xa0", " ")        # NBSP (python literal)
         .replace("&nbsp;", " ")      # HTML entity NBSP
         .replace("\u200b", "")       # zero-width space
         .replace("\u200c", "")       # zero-width non-joiner
         .replace("\u200d", "")       # zero-width joiner
         .replace("\ufeff", "")       # zero-width no-break space / BOM
         .replace("\u00ad", "")       # soft hyphen
         .replace("\n", " ") 
         .replace("\"", "")         # newline -> space
         .strip()
    )
    s = re.sub(r"\s+", " ", s)
    return s

####################################################################
# Scraping Logic
####################################################################

def _try_extract_nonce_from_html(html: str) -> str | None:
    # Common WordPress pattern: a JS object with "nonce":"..."
    m = re.search(r'"nonce"\s*:\s*"([^"]+)"', html)
    if m:
        return m.group(1).strip()

    soup = BeautifulSoup(html, "html.parser")
    meta = soup.find("meta", attrs={"name": re.compile(r"x-wp-nonce", re.I)})
    if meta and meta.get("content"):
        return meta["content"].strip()

    return None

def bootstrap_session() -> tuple[requests.Session, str | None]:
    """
    Creates a session and visits the list page first (cookies).
    Returns (session, nonce_if_found).
    The REST API nonce is usually sent via X-WP-Nonce header. [2](https://developer.wordpress.org/rest-api/using-the-rest-api/authentication/)
    """
    s = requests.Session()
    r = s.get(LIST_PAGE, headers={"User-Agent": DEFAULT_HEADERS["User-Agent"]}, timeout=60, verify=False)
    r.raise_for_status()

    nonce = r.headers.get("X-WP-Nonce") or r.headers.get("x-wp-nonce")
    if nonce:
        return s, nonce.strip()

    nonce = _try_extract_nonce_from_html(r.text)
    return s, nonce

def fetch_certificates_page(
    session: requests.Session,
    page: int,
    count: int = 100,
    search: str = "",
    valid_from: str = "",
    valid_until: str = "",
    nonce: str | None = None,
    status_filter: str | None = None,
) -> tuple[str, int, int]:
    """
    Returns (html, totalCount, maxPages)
    """
    payload = {
        "valid_from": valid_from,
        "valid_until": valid_until,
        "search": search,
        "count": str(count),
        "page": int(page),
    }

    # Add status filter when provided
    if status_filter:
        payload["filters"] = {"status": [status_filter]}

    headers = dict(DEFAULT_HEADERS)
    if nonce:
        headers["X-WP-Nonce"] = nonce

    r = session.post(API_URL, headers=headers, data=json.dumps(payload), timeout=60, verify=False)

    # If nonce missing/expired, refresh once by revisiting LIST_PAGE and retry
    if r.status_code in (401, 403):
        rp = session.get(LIST_PAGE, headers={"User-Agent": DEFAULT_HEADERS["User-Agent"]}, timeout=60, verify=False)
        rp.raise_for_status()
        new_nonce = rp.headers.get("X-WP-Nonce") or rp.headers.get("x-wp-nonce") or _try_extract_nonce_from_html(rp.text)
        if new_nonce:
            headers["X-WP-Nonce"] = new_nonce
            r = session.post(API_URL, headers=headers, data=json.dumps(payload), timeout=60, verify=False)

    r.raise_for_status()

    try:
        js = r.json()
    except requests.exceptions.JSONDecodeError:
        js = json.loads(r.content.decode("utf-8-sig"))
    block = js.get("data", {}).get("data", {})
    html = block.get("html", "")
    total = int(block.get("totalCount", 0))
    max_pages = int(block.get("maxPages", 0))

    return html, total, max_pages


def _text(el):
    return el.get_text(" ", strip=True) if el else ""

def parse_certificates_html(html: str, status_value: str = "") -> list:
    if not html:
        return []

    soup = BeautifulSoup(html, "lxml")
    cards = soup.select("div.is-certificate")
    out = []

    for card in cards:
        cert_id = _text(card.select_one(".tag"))

        # Validity range
        date_text = _text(card.select_one(".date"))
        valid_from = ""
        valid_until = ""
        if date_text:
            parts = [p.strip() for p in date_text.replace("–", "-").split("-") if p.strip()]
            if len(parts) >= 2:
                valid_from, valid_until = parts[0], parts[1]
            elif len(parts) == 1:
                valid_from = parts[0]

        # Certificate holder (tooltip title has full string)
        holder_span = card.select_one("h3 span.has-tip")
        holder_full = holder_span.get("title", "").strip() if holder_span else ""
        holder_display = _text(holder_span)

        # Suspended period (only appears for suspended certs)
        suspended_period = ""
        suspend_el = card.select_one("p.suspend-date")
        if suspend_el:
            # This will collapse whitespace and treat <br> as a space
            s_text = suspend_el.get_text(" ", strip=True)
            # Example becomes "21.04.26 – 01.06.26" or "21.04.26 - 01.06.26"
            s_text = s_text.replace("–", "-")
            s_parts = [p.strip() for p in s_text.split("-") if p.strip()]
            if len(s_parts) >= 2:
                suspended_period = f"{s_parts[0]} – {s_parts[1]}"
            elif len(s_parts) == 1:
                suspended_period = s_parts[0]

        scope = ""
        processing_unit_type = ""
        raw_material = ""
        products = ""
        add_ons = ""
        issuing_cb = ""

        fold_items = card.select(".is-certificate-fold .is-certificate-fold-item")
        for item in fold_items:
            title = _text(item.select_one(".title")).lower()
            value = _text(item.select_one("p:not(.title)"))

            if title == "scope":
                scope = value
            elif title == "processing unit type":
                processing_unit_type = value
            elif title == "raw material":
                raw_material = value
            elif title == "products":
                products = value
            elif "add-ons" in title or "add-ons/cts" in title:
                add_ons = value
            elif title == "issuing cb":
                issuing_cb = value

        map_link = ""
        audit_link = ""
        cert_link = ""

        for a in card.select("a.custom-button"):
            label = _text(a).lower()
            href = (a.get("href") or "").strip()
            if not href:
                continue
            if "geolocation" in label:
                map_link = href
            elif "audit" in label:
                audit_link = href
            elif "certificate" in label:
                cert_link = href

        out.append({
            "cert_status": status_value,
            "cert_number": cert_id,
            "cert_owner": holder_full or holder_display,
            "cert_scope": scope,
            "cert_processingunittype": processing_unit_type,
            "cert_in_put": raw_material,
            "cert_add_on": add_ons,
            "cert_products": products,
            "cert_valid_from": valid_from,
            "cert_valid_until": valid_until,
            "cert_suspended_date": suspended_period if status_value == "suspended" else (suspended_period or ""),
            "cert_issuer": issuing_cb,
            "cert_map": map_link,
            "cert_file": cert_link,
            "cert_audit": audit_link,
        })

    return out


def split_cert_owner(value):
    """Split 'Company, City, Country' into 3 separate columns safely."""
    if not value or not isinstance(value, str):
        return "", "", ""

    parts = [p.strip() for p in value.split(",") if p.strip()]

    # Handle names with internal commas
    if len(parts) >= 3:
        company = parts[0]
        city = parts[1]
        country = parts[-1]
        return company, city, country

    if len(parts) == 2:
        return parts[0], parts[1], ""

    if len(parts) == 1:
        return parts[0], "", ""

    return "", "", ""


def scrape_all(output_file, page_size=200, delay=0, search="", valid_from="", valid_until=""):
    session, nonce = bootstrap_session()
    print("Initial nonce:", nonce)

    all_rows = []

    for status in STATUS_TYPES:
        print(f"\n--- Scraping status bucket: {status} ---")

        html, total_records, max_pages = fetch_certificates_page(
            session=session,
            page=1,
            count=page_size,
            search=search,
            valid_from=valid_from,
            valid_until=valid_until,
            nonce=nonce,
            status_filter=status,
        )

        print(f"{status}: Total certificates: {total_records}, max pages: {max_pages}")

        rows = parse_certificates_html(html, status_value=status)
        all_rows.extend(rows)

        for page in range(2, max_pages + 1):
            if page % 50 == 0:
                print(f"{status}: Fetching page {page} of {max_pages} ...")

            try:
                html, _, _ = fetch_certificates_page(
                    session=session,
                    page=page,
                    count=page_size,
                    search=search,
                    valid_from=valid_from,
                    valid_until=valid_until,
                    nonce=nonce,
                    status_filter=status,
                )

                rows = parse_certificates_html(html, status_value=status)
                if not rows:
                    print(f"{status}: No rows returned on page {page}, stopping this bucket.")
                    break

                all_rows.extend(rows)
                if delay:
                    time.sleep(delay)

            except Exception as e:
                print(f"Error on status {status}, page {page}: {e}")
                break

    # Build DataFrame
    df = pd.DataFrame(all_rows, columns=COLUMNS)

    # Optional: Deduplicate by cert_number with status priority
    if not df.empty:
        df["_status_rank"] = df["cert_status"].map(lambda x: STATUS_PRIORITY.get(x, 999))
        df = df.sort_values(["cert_number", "_status_rank"]).drop_duplicates(subset=["cert_number"], keep="first")
        df = df.drop(columns=["_status_rank"])

    # Extract new cert_owner fields
    company_series, city_series, country_series = zip(*df["cert_owner"].apply(split_cert_owner))

    # Add the manual country overrides to the countries list
    country_series = [MANUAL_COUNTRY_OVERRIDES.get(get_country_name(c), get_country_name(c)) for c in country_series]

    # Insert company, city, country directly after cert_owner
    owner_index = df.columns.get_loc("cert_owner") + 1
    df.insert(owner_index, "Company_Name", company_series)
    df.insert(owner_index + 1, "City", [c.capitalize() for c in city_series])
    df.insert(owner_index + 2, "Country", country_series)

    df["City"] = df["cert_owner"].apply(get_city_name)

    df.insert(df.columns.get_loc("cert_number") + 1, "Certificate_Type", df["cert_number"].apply(map_certificate_type))
    df.insert(df.columns.get_loc("Country") + 1, "Region", df["Country"].apply(map_region))
    df.insert(df.columns.get_loc("Country") + 2, "Sub_Region", df["Country"].apply(map_subregion))

    # Fill Status column from cert_status (string), not numeric map_status
    df["cert_status"] = df["cert_status"].astype(str).str.capitalize()

    df.insert(df.columns.get_loc("cert_number") + 2, "Certificate_Class", df["Certificate_Type"].apply(map_certificate_class))
    df.insert(df.columns.get_loc("cert_map") + 1, "Latitude", df["cert_map"].apply(get_latitude))
    df.insert(df.columns.get_loc("cert_map") + 2, "Longitude", df["cert_map"].apply(get_longitude))

    df = df.rename(columns=COLUMN_MAP)

    # Add the facility grouping column
    df.insert(
        df.columns.get_loc("Scope") + 1,
        "Facility_Grouping",
        df["Scope"].apply(determine_facility_grouping)
    )

    # Normalise to remove whitespaces and invisible characters
    df = df.map(clean_excel_string)

    df = overwrite_company_with_gst_shortname_exact(df, GST_ASSETS, score_threshold=75)
    df = add_asset_identifier_and_match(df, GST_ASSETS)

    df["Asset_Identifier"] = np.where(
        df["Match_Found"] == 1,
        df["Company_City"],
        None
    )

    df = df.replace(r"^\s*nan\s*$", "", regex=True)

    exclude = {"Scope_Description", "Processing_Unit_Type_Description",
               "Map", "Certificate", "Audit_Report", "Products", "Add-ons** /CTS"}

    text_cols = [c for c in df.select_dtypes(include="object").columns if c not in exclude]
    df[text_cols] = df[text_cols].replace(r'[\\/\"\'„“»«]', "", regex=True)

    df.to_excel(output_file, index=False, engine="openpyxl", sheet_name="Certificate Database")
    print(f"\nScraping complete! Saved {len(df)} rows to {output_file}")

# TODO: clean up this file from a commenting POV
# TODO: Setup the the correct SSL verify flag in the requests calls. For now, we are ignoring SSL warnings and setting verify=False in the requests calls to avoid SSL errors. This is not recommended for production use, but it allows us to proceed with scraping without SSL issues. We should investigate the root cause of the SSL errors and fix them properly in the future.


