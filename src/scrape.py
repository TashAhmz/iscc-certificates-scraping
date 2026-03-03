import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
import re
from mappings import *
from thefuzz import fuzz, process
import numpy as np

# URLs
BASE_URL = "https://www.iscc-system.org/wp-admin/admin-ajax.php?action=get_wdtable&table_id=2"
MAIN_PAGE = "https://www.iscc-system.org/certification/certificate-database/all-certificates/"

# GSTs of Geo filepath
GST_GEO = pd.read_excel("C:/Users/tashif.ahmed/OneDrive - Shell/T&S LCF - Analytics, Digital, and Economics - Shared Documents/00. LCF Data Lakehouse/GSTs/GST Geographies/LCF GST of Geographies.xlsx", sheet_name="GS_LCF_Geographies")

# GSTs of Assets filepath
GST_ASSETS = pd.read_excel(r"C:/Users/tashif.ahmed/OneDrive - Shell/T&S LCF - Analytics, Digital, and Economics - Shared Documents/00. LCF Data Lakehouse/GSTs/GST Assets/00. Golden Source File of Asset Capacities.xlsm", sheet_name="GoldenSource")

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

def _safe(text):
    return "" if text is None else str(text).strip()

def _asset_identifier_join(company_name, city):
    company = _safe(company_name)
    city = _safe(city)
    return f"{company} {city}".strip()

def _normalize_for_match(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = s.lower()
    s = re.sub(r"[\\.,;:/\\-\\(\\)\\[\\]&]", " ", s)  # remove punctuation
    s = re.sub(r"\\s+", " ", s).strip()               # collapse spaces
    return s


def add_asset_identifier_and_match(
    df_iscc: pd.DataFrame,
    gst_df: pd.DataFrame,
    fuzzy_threshold: int  # kept in signature to avoid breaking callers; unused now
) -> pd.DataFrame:

    # 1) Build initial ISCC Asset_Identifier
    if "Company_Name" not in df_iscc.columns or "City" not in df_iscc.columns:
        raise KeyError("Expected columns 'Company_Name' and 'City' not found in ISCC DataFrame.")

    df_iscc["Asset_Identifier"] = [
        _asset_identifier_join(cn, city)
        for cn, city in zip(df_iscc["Company_Name"], df_iscc["City"])
    ]

    # 2) Prepare GST lists and mappings
    GST_COL = "Asset Identifier"
    if GST_COL not in gst_df.columns:
        raise KeyError(f"Column '{GST_COL}' not found in GST assets DataFrame.")

    gst_raw_list = gst_df[GST_COL].astype(str).tolist()
    gst_norm_list = [_normalize_for_match(x) for x in gst_raw_list]
    gst_norm_set = set(gst_norm_list)
    norm_to_raw = {n: r for n, r in zip(gst_norm_list, gst_raw_list)}
        
    def _tokens(s: str):
        """
        Simple tokenization for keyword matching:
        - keeps short but meaningful tokens like 'bp' (len>=2)
        - drops pure numeric tokens like '1', '3044'
        - collapses dotted abbreviations: 'a.s.' -> 'as', 's.a.' -> 'sa', 'b.v.' -> 'bv'
        - drops legal suffix tokens using LEGAL_SUFFIXES
        """
        s = _normalize_for_match(s).lower()

        # collapse dotted abbreviations: a.s. -> as, s.r.o. -> sro, etc.
        s = re.sub(r"\b(?:[a-z]\.){2,}", lambda m: m.group(0).replace(".", ""), s)

        # split on whitespace after normalization
        raw = s.split()

        out = set()
        for t in raw:
            t = t.strip()
            if not t:
                continue
            if len(t) < 2:          # kills 'a' and 's'
                continue
            if t.isdigit():         # kills '1' '3044'
                continue
            if t in LEGAL_SUFFIXES:
                continue
            if t in STOPWORDS:
                continue
            out.add(t)

        return out

    def _safe_str(x):
        return "" if x is None else str(x)

    # make company tokens & city tokens disjoint
    def _disjoint_tokens(comp_tokens: set, city_tokens: set):
        overlap = comp_tokens & city_tokens
        if overlap:
            comp_tokens = comp_tokens - overlap
            city_tokens = city_tokens - overlap
        return comp_tokens, city_tokens

    def _middle_from_cert_holder(cert_holder: str) -> str:
        """
        Optional enhancement:
        "Neste Components B.V, Botlek, Rotterdam, Netherlands"
           -> "Botlek Rotterdam"
        Returns "" if not parseable.
        """
        if not cert_holder:
            return ""
        parts = [p.strip() for p in str(cert_holder).split(",") if p.strip()]
        if len(parts) < 3:
            return ""
        return " ".join(parts[1:-1]).strip()

    gst_token_sets = [_tokens(x) for x in gst_raw_list]

    match_flags = []
    overwritten_asset_ids = []

    # --- Main processing loop ---
    for row in df_iscc.itertuples(index=False):

        company_name = row.Company_Name
        city_name = row.City
        cert_holder = getattr(row, "Certificate_Holder", None)
        asset_id = row.Asset_Identifier

        original_display = asset_id
        norm = _normalize_for_match(asset_id)

        # Empty case
        if norm == "":
            match_flags.append(0)
            overwritten_asset_ids.append(original_display)
            continue

        # Exact match
        if norm in gst_norm_set:
            match_flags.append(1)
            overwritten_asset_ids.append(norm_to_raw[norm])
            continue

        # --- Keyword Inclusion Attempt ---
        comp_tokens = _tokens(_safe_str(company_name))

        # City reference: prefer middle-of-cert-holder (Company, ..., Country)
        # fallback to existing get_city_name(cert_holder) if middle extraction fails.
        cert_holder_str = _safe_str(cert_holder)
        mid_from_holder = _middle_from_cert_holder(cert_holder_str)
        if mid_from_holder:
            mid_city = mid_from_holder
        #else:
            #mid_city = get_city_name(cert_holder_str)

        city_tokens = _tokens(mid_city)

        # enforce company and city tokens CANNOT be the same
        comp_tokens, city_tokens = _disjoint_tokens(comp_tokens, city_tokens)

        keyword_matched_index = None

        # CHANGED: allow matching even if only one side exists
        if comp_tokens or city_tokens:
            best_score = -1
            best_idx = -1

            for i, gst_toks in enumerate(gst_token_sets):
                comp_hit = len(comp_tokens & gst_toks) if comp_tokens else 0
                city_hit = len(city_tokens & gst_toks) if city_tokens else 0

                # CHANGED acceptance rules:
                # A) combined: at least 1 hit on both
                # B) company-only: >= 2 unique company hits AND 0 city hits
                # C) city-only: >= 2 unique city hits AND 0 company hits
                combined_ok = (comp_hit > 0 and city_hit > 0)
                company_only_ok = (comp_hit >= 2 and city_hit == 0)

                if combined_ok or company_only_ok:
                    # Score strategy:
                    # Prefer combined matches if available, otherwise use hit count.
                    if combined_ok:
                        score = 100 + comp_hit + city_hit
                    else:
                        score = comp_hit + city_hit  # will be >=2 in solo cases

                    if score > best_score:
                        best_score = score
                        best_idx = i

            if best_idx >= 0:
                keyword_matched_index = best_idx

        if keyword_matched_index is not None:
            best_match = gst_raw_list[keyword_matched_index]
            match_flags.append(1)
            overwritten_asset_ids.append(best_match)
            continue

        # CHANGED: removed fuzzy fallback completely
        match_flags.append(0)
        overwritten_asset_ids.append(original_display)

    df_iscc["Asset_Identifier"] = overwritten_asset_ids
    df_iscc["Match_Found"] = match_flags
    return df_iscc


def _normalize(text: str) -> str:
    """Light normalization to improve fuzzy company matches."""
    if not isinstance(text, str):
        return ""
    text = text.lower()
    removals = [
        " inc", " llc", " l.l.c", " lp", " l.p.", " bv", " b.v.", " ltd", " co", " company",
        " limited", ".", ",", "&", "ltd.", "pte.", "gmbh", "ag", "plc", "s.p.a"
    ]
    # pad with space to avoid partial token issues (e.g., 'lp' in 'help')
    for w in removals:
        text = text.replace(w, " ")
    return " ".join(text.split())

def _build_lookup_exact_columns(gst_df: pd.DataFrame):
    """
    Build the fuzzy universe and a map of normalized name -> 'Company/Producer Short Name'
    using the exact columns from your GST of Assets file.
    """
    # Use exact headers as shown in your screenshot
    CP_COL = "Company/Producer"
    CPSN_COL = "Company/Producer Short Name"

    # Guard rails
    for col in (CP_COL, CPSN_COL):
        if col not in gst_df.columns:
            raise KeyError(f"Column '{col}' not found in GST assets DataFrame.")

    tmp = gst_df[[CP_COL, CPSN_COL]].copy()
    tmp["__norm_cp__"]   = tmp[CP_COL].apply(_normalize)
    tmp["__norm_cpsn__"] = tmp[CPSN_COL].apply(_normalize)

    # Universe of strings to match against (both full and short)
    universe = list(tmp["__norm_cp__"]) + list(tmp["__norm_cpsn__"])

    # Map normalized representation -> short name (original casing)
    to_short = {}
    for _, r in tmp.iterrows():
        to_short[r["__norm_cp__"]]   = r[CPSN_COL]
        to_short[r["__norm_cpsn__"]] = r[CPSN_COL]

    return universe, to_short


def overwrite_company_with_gst_shortname_exact(
    iscc_df: pd.DataFrame,
    gst_df: pd.DataFrame,
    score_threshold
) -> pd.DataFrame:
    """
    Overwrites iscc_df['Company_Name'] with GST 'Company/Producer Short Name'
    when fuzzy match >= score_threshold AND the GST short name is non-empty.
    Otherwise, leaves the original value unchanged.

    Guarantees: never replaces a non-empty Company_Name with a blank.
    """

    if "Company_Name" not in iscc_df.columns:
        raise KeyError("Expected column 'Company_Name' not found in ISCC DataFrame.")

    # Build lookup using your exact columns (as previously aligned)
    # This function must return:
    #   universe: List[str] of normalized GST names (producer + short)
    #   to_short: Dict[normalized_name] -> short_name (original casing)
    universe, to_short = _build_lookup_exact_columns(gst_df)

    # Helper to coerce to safe string for "leave as-is"
    def _as_is(value):
        # If original is NaN/None, return "" (or choose a placeholder if you prefer)
        return "" if (value is None or (isinstance(value, float) and np.isnan(value))) else str(value)

    # If the GST universe is empty, just return untouched
    if not universe:
        # Still normalize dtype to string to avoid later Excel issues
        iscc_df["Company_Name"] = iscc_df["Company_Name"].astype(str)
        return iscc_df

    new_values = []
    for original in iscc_df["Company_Name"]:
        original_safe = _as_is(original)

        # If original is empty, there's nothing sensible to overwrite with "as-is"
        # but we still try to find a match; otherwise keep it empty.
        norm = _normalize(original_safe)

        # If normalization yields empty string, we skip fuzzy and keep original
        if not norm.strip():
            new_values.append(original_safe)
            continue

        match, score = process.extractOne(norm, universe, scorer=fuzz.ratio) if universe else (None, 0)

        # Only overwrite if:
        #  1) We found a match
        #  2) Score >= threshold
        #  3) The mapped short name is a non-empty string
        if match and score >= score_threshold:
            candidate = to_short.get(match, "")
            candidate = candidate if isinstance(candidate, str) else _as_is(candidate)
            candidate = candidate.strip()

            if candidate:  # ensure not blank
                new_values.append(candidate)
            else:
                # Leave as original if GST short name is empty or missing
                new_values.append(original_safe)
        else:
            # No acceptable match → leave as-is
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
                    "llc", "l.l.c.", "llc.", "lp", "inc.."]
    
    #tokens = [t.strip().lower() for t in cert_owner.split(",") if t.strip() and t.strip().lower() not in exempt_words][1:-1] # only including valid cities possibilties
   
    parts = [p.strip().lower() for p in cert_owner.split(",") if p.strip()][1:-1]

    tokens = [
        tok
        for tok in parts
        if tok
        and tok not in exempt_words
        and not any(w in CITY_STOPWORDS for w in tok.split())
        and not every_word_has_digit(tok)
    ]

    country = get_country_name(cert_owner).lower()

    # testing to see if this logic works to remove street names coming into the city column by mistake
    if len(tokens) == 1:
        return " ".join([w for w in tokens[0].split() if not w.isnumeric()]).title()
    elif len(tokens) >= 2:
        if country in ("united states", "china"):
            return " ".join([w for w in tokens[-2].split() if not w.isnumeric()]).title()
        else:
            return " ".join([w for w in tokens[-1].split() if not w.isnumeric()]).title()
        
    return None
"""
# Tom's axtract city function
def extract_city(cert_holder):
        if not isinstance(cert_holder, str):
            return ""
        
        # Split the string by commas and remove empty parts
        parts = [p.strip() for p in cert_holder.split(",") if p.strip()]
        
        if len(parts) >= 3:
            second_last = parts[-2]
            last = parts[-1]  # This line defines 'last'

            # Special case: if second-last is "Korea", use third-last as city
            if second_last.lower() == "korea":
                return parts[-3]
            
            # Special case: if country is Hong Kong, city is also Hong Kong # make logic for Singapore
            if last.lower() == "hong kong":
                return "Hong Kong"

            if last.lower() == "china":
                return ""
            
            
            # If second-last is a 2-letter code or a known country, use third-last. Specifically for USA States 
            if (
                (len(second_last) == 2 and second_last.isupper()) or
                (second_last.lower() in [c.lower() for c in column_a_values])
            ):
                return parts[-3]
            
            # Otherwise, use second-last as city
            return parts[-2]
        
        elif len(parts) == 2:
            return parts[-2]
        
        return ""
"""

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
    df.insert(scope_index + 2, "Processing_Unit_Type_Description", df["cert_scope"].apply(map_multiple_scopes))

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

    df = overwrite_company_with_gst_shortname_exact(df, GST_ASSETS, score_threshold=70)

    df = add_asset_identifier_and_match(df, GST_ASSETS, fuzzy_threshold=85)

    # Save and add styles
    df.to_excel(output_file, index=False, engine="openpyxl", sheet_name="Certificate Database")

    print(f"Scraping complete! Saved {len(df)} rows to {output_file}")

# TODO: clean up this file from a commenting POV


