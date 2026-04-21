#This is the old scrape file before ISCC changed their website  and broke the old logic.




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

# URLs
BASE_URL = "https://www.iscc-system.org/wp-admin/admin-ajax.php?action=get_wdtable&table_id=2"
MAIN_PAGE = "https://www.iscc-system.org/certification/certificate-database/all-certificates/"

# GSTs of Geo filepath
GST_GEO = pd.read_excel("C:/Users/tashif.ahmed/OneDrive - Shell/T&S LCF - Analytics, Digital, and Economics - Shared Documents/00. LCF Data Lakehouse/GSTs/GST Geographies/LCF GST of Geographies.xlsx", sheet_name="GS_LCF_Geographies")

# GSTs of Assets filepath
GST_ASSETS = pd.read_excel(r"C:/Users/tashif.ahmed/OneDrive - Shell/T&S LCF - Analytics, Digital, and Economics - Shared Documents/00. LCF Data Lakehouse/GSTs/GST Assets/00. Golden Source File of Asset Capacities.xlsm", sheet_name="GoldenSource")

countries = set(c.lower() for c in ALL_COUNTRIES)
stopwords = set(s.lower() for s in STOPWORDS)   
city_stopwords = set(s.lower() for s in CITY_STOPWORDS)

# Headers
HEADERS = {
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest"
}

# Column names (from table)
COLUMNS = [
    "cert_ikon","cert_number","cert_owner","cert_scope","cert_processingunittype","cert_in_put","cert_add_on",
    "cert_products","cert_valid_from","cert_valid_until","cert_suspended_date",
    "cert_issuer","cert_map","cert_file","cert_audit","cert_status"
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


import re


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
    if not isinstance(link, str) or "maps?q=" not in link:
        return None, None
    coords = link.split("maps?q=")[-1].split(",")
    # Filter out empty strings
    coords = [c.strip() for c in coords if c.strip()]
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

def map_status(code):
    try:
        code = int(code)
    except (ValueError, TypeError):
        return ""
    return STATUS_MAP.get(code, "Unknown")

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

def get_fresh_nonce():
    """Fetch the main page and extract the current wdtNonce"""
    response = requests.get(MAIN_PAGE, verify=False)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    input_tag = soup.find("input", {"id": "wdtNonceFrontendEdit_2"})
    if input_tag and input_tag.has_attr("value"):
        return input_tag["value"]
    else:
        raise ValueError("Could not find wdtNonce on the page")

def fetch_page(start: int, length: int = 10000, nonce: str = None):
    """Fetch a page of certificates from the server"""
    if nonce is None:
        nonce = get_fresh_nonce()

    form_data = {
        "draw": "5",
        "order[0][column]": "4",
        "order[0][dir]": "desc",
        "start": str(start),
        "length": str(length),
        "search[value]": "",
        "search[regex]": "false",
        "wdtNonce": nonce,
        "sRangeSeparator": "|"
    }

    # Add columns for server-side processing
    for i, name in enumerate(COLUMNS):
        form_data[f"columns[{i}][data]"] = str(i)
        form_data[f"columns[{i}][name]"] = name
        form_data[f"columns[{i}][searchable]"] = "true"
        form_data[f"columns[{i}][orderable]"] = "true"
        form_data[f"columns[{i}][search][value]"] = ""
        form_data[f"columns[{i}][search][regex]"] = "false"

    response = requests.post(BASE_URL, headers=HEADERS, data=form_data, verify=False)
    response.raise_for_status()

    js = response.json()
    return js["data"], int(js["recordsTotal"])

def parse_rows(rows):
    """Clean HTML in each cell and extract links (PDFs, maps) safely"""
    clean_rows = []
    for row in rows:
        clean_row = []
        for cell in row:
            if cell is None:
                clean_row.append("")
                continue
            soup = BeautifulSoup(str(cell), "html.parser")
            
            # Check if there is an <a> tag and check if the cell contains a tooltip and extract that
            link = soup.find("a", href=True)
            tooltip = soup.find("span", class_="has-tip top", tabindex=2)
            if link:
                # Extract the href
                clean_row.append(link["href"].strip())
            elif tooltip:
                 clean_row.append(tooltip["title"].strip())
            else:
                # Otherwise, just text
                clean_row.append(soup.get_text(strip=True))    
        clean_rows.append(clean_row)
    return clean_rows

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

def map_multiple_scopes(scope_value):
    if not scope_value:
        return "Unknown"
    codes = [code.strip() for code in scope_value.split(",")]
    descriptions = [SCOPE_DESCRIPTIONS.get(code, "No Mapping") for code in codes]
    return ", ".join(descriptions)

def scrape_all(output_file, page_size, delay):
    """Scrape all certificates and save to CSV"""
    nonce = get_fresh_nonce()
    print("Using nonce:", nonce)

    # First page to get total records
    rows, total_records = fetch_page(start=0, length=page_size, nonce=nonce)
    print(f"Total certificates: {total_records}")

    all_rows = parse_rows(rows)

    for start in range(page_size, total_records, page_size):
        print(f"Fetching rows {start} to {start+page_size}...")
        try:
            rows, _ = fetch_page(start=start, length=page_size, nonce=nonce)
            if not rows:
                print("No more rows returned, stopping.")
                break
            all_rows.extend(parse_rows(rows))
            time.sleep(delay) # polite delay
        except Exception as e:
            print(f"Error fetching page starting at {start}: {e}")
            break
    
    # Save to XLSX
    df = pd.DataFrame(all_rows, columns=COLUMNS)

    # Insert "scope_description" after "scope"
    scope_index = df.columns.get_loc("cert_scope") + 1
    df.insert(scope_index, "Scope_Description", df["cert_scope"].apply(map_multiple_scopes))

    # Insert "Processing_Unit_Type_Description"
    df.insert(scope_index + 2, "Processing_Unit_Type_Description", df["cert_processingunittype"].apply(map_multiple_scopes))

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

    # Add the facility grouping column
    df.insert(
        df.columns.get_loc("Scope_Description") + 1,
        "Facility_Grouping",
        df["cert_scope"].apply(determine_facility_grouping)
    )

    columns_to_remove = ["cert_ikon"]  # Add more if needed
    df = df.drop(columns=columns_to_remove)

    df.insert(df.columns.get_loc("cert_number") + 1, "Certificate_Type", df["cert_number"].apply(map_certificate_type))
    df.insert(df.columns.get_loc("Country") + 1, "Region", df["Country"].apply(map_region))
    df.insert(df.columns.get_loc("Country") + 2, "Sub_Region", df["Country"].apply(map_subregion))
    df.insert(0, "Status", df["cert_status"].apply(map_status))
    df.insert(df.columns.get_loc("cert_number") + 2, "Certificate_Class", df["Certificate_Type"].apply(map_certificate_class))
    df.insert(df.columns.get_loc("cert_map") + 1, "Latitude", df["cert_map"].apply(get_latitude))
    df.insert(df.columns.get_loc("cert_map") + 2, "Longitude", df["cert_map"].apply(get_longitude))

    df = df.rename(columns=COLUMN_MAP)

    # Normalise to remove whitespaces and invisible characters that could break further logic
    df = df.map(clean_excel_string)

    df = overwrite_company_with_gst_shortname_exact(df, GST_ASSETS, score_threshold=75)

    df = add_asset_identifier_and_match(df, GST_ASSETS)

    # Add new column called Asset_Identifier for ones where match found, otherwise keep blank
    df["Asset_Identifier"] = np.where(
        df["Match_Found"] == 1,
        df["Company_City"],
        None
    )

    # extra cleaning
    df = df.replace(r"^\s*nan\s*$", "", regex=True)
    
    exclude = {"Scope_Description", "Processing_Unit_Type_Description",
                "Map", "Certificate", "Audit_Report", "Products", "Add-ons** /CTS"} 

    text_cols = [c for c in df.select_dtypes(include="object").columns if c not in exclude]

    df[text_cols] = df[text_cols].replace(r'[\\/\"\'„“»«]', "", regex=True)


    # Save and add styles
    df.to_excel(output_file, index=False, engine="openpyxl", sheet_name="Certificate Database")

    print(f"Scraping complete! Saved {len(df)} rows to {output_file}")

# TODO: clean up this file from a commenting POV


