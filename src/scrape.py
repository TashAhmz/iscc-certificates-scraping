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
from asset_matching import match_assets_to_gst

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
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9,en-GB;q=0.8",
    "Content-Type": "application/json",
    "Origin": "https://iscc-system.org",
    "Referer": "https://iscc-system.org/certification/all-certificates/",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/149.0.0.0 Safari/537.36 Edg/149.0.0.0"
    ),
    "sec-ch-ua": '"Microsoft Edge";v="149", "Chromium";v="149", "Not)A;Brand";v="24"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
}


def bootstrap_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(DEFAULT_HEADERS)

    # Visit page first to collect normal public cookies.
    r = s.get(LIST_PAGE, timeout=60, verify=False)
    r.raise_for_status()

    print("Cookies:", s.cookies.get_dict())

    return s
STATUS_TYPES = ["valid"] #["valid", "suspended", "expired", "terminated", "withdrawn"]



STATUS_PRIORITY = {s: i for i, s in enumerate(["withdrawn", "terminated", "suspended", "expired", "valid"])}


# Column names (from table)
COLUMNS = [
    "cert_status", "cert_number","cert_owner","cert_scope","cert_processingunittype","cert_in_put","cert_add_on",
    "cert_products","cert_valid_from","cert_valid_until","cert_suspended_date",
    "cert_issuer","cert_map","cert_file","cert_audit"
]

# Define a function to determine the facility grouping based on Scope* codes
    # It checks each abbreviation and returns the matching group(s)
def determine_facility_grouping(row):

    groupings = set()

    scope_text = row.get("Scope", "")
    processing_unit_type = row.get("Processing_Unit_Type", "")

    # Process Scope values
    if isinstance(scope_text, str):

        scope_values = [
            value.strip()
            for value in scope_text.split(",")
            if value.strip()
        ]

        for value in scope_values:

            # Skip Processing Unit since we'll use Processing_Unit_Type instead
            if value == "Processing Unit":
                continue

            group = FACILITY_GROUPING_MAP.get(value)
            if group:
                groupings.add(group)

    # If Processing Unit is present in Scope, also process Processing_Unit_Type
    if (
        isinstance(scope_text, str)
        and "Processing Unit" in scope_text
        and isinstance(processing_unit_type, str)
    ):

        pu_values = [
            value.strip()
            for value in processing_unit_type.split(",")
            if value.strip()
        ]

        for value in pu_values:
            group = FACILITY_GROUPING_MAP.get(value)
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


def fetch_certificates_page(
    session: requests.Session,
    page: int,
    count: int = 100,
    search: str = "",
    valid_from: str = "",
    valid_until: str = "",
    status_filter: str | None = None,
) -> tuple[str, int, int]:

    payload = {
        "valid_from": valid_from,
        "valid_until": valid_until,
        "search": search,
        "count": str(count),
        "page": int(page),
    }

    if status_filter:
        payload["filters"] = {"status": [status_filter]}

    r = session.post(API_URL, json=payload, timeout=60, verify=False)

    if r.status_code != 200:
        print("Status:", r.status_code)
        print("Response body:", r.text[:3000])

    r.raise_for_status()

    js = r.json()
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
    html = re.sub(r"<\\?xml.*?\\?>", "", html, flags=re.DOTALL)
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
    session = bootstrap_session()

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

    df["Processing_Unit_Type"] = df["Processing_Unit_Type"].str.replace(" | ", ", ", regex=False)

    # Add the facility grouping column
    df.insert(
        df.columns.get_loc("Scope") + 1,
        "Facility_Grouping",
        df.apply(determine_facility_grouping, axis=1)
    )

    # Normalise to remove whitespaces and invisible characters
    df = df.map(clean_excel_string)

   
    # Match ISCC certificates against GST assets
    df = match_assets_to_gst(
        iscc_df=df,
        gst_df=GST_ASSETS,
        include_match_diagnostics = False
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


