"""
ISCC PDF material-flow extractor

This script reads an ISCC certificate export, downloads each certificate PDF, and
extracts the annex table headed with:

    Input material | Output material | GHG option | Criteria | Add-ons

It writes an Excel workbook with:
    1. PDF Material Flows       - one row per annex table row
    2. PDF Certificate Summary  - one row per certificate
    3. PDF Extraction Failures  - download/parse issues

Dependencies:
    pip install pandas requests pdfplumber pymupdf opencv-python pytesseract pillow openpyxl tqdm

For scanned/image PDFs, Tesseract must also be installed on the machine.
On Windows, install it and pass --tesseract-cmd if needed.
"""

from __future__ import annotations

import argparse
import concurrent.futures as futures
import hashlib
import os
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlparse

import cv2
import numpy as np
import pandas as pd
import pdfplumber
import requests

try:
    import fitz  # PyMuPDF
except Exception as exc:  # pragma: no cover
    raise RuntimeError("PyMuPDF is required. Install with: pip install pymupdf") from exc

try:
    import pytesseract
except Exception:
    pytesseract = None


DEFAULT_HEADERS = {
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9,en-GB;q=0.8",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/126.0.0.0 Safari/537.36"
    ),
}

CERT_NUMBER_COL_CANDIDATES = [
    "Certificate_Number",
    "Certificate Number",
    "cert_number",
    "Certificate No",
    "Certificate_No",
    "Certificate_ID",
]

CERT_URL_COL_CANDIDATES = [
    "Certificate",
    "cert_file",
    "Certificate_File",
    "Certificate File",
    "certificate_file",
    "Certificate_URL",
    "cert_url",
]

STATUS_COL_CANDIDATES = [
    "Status",
    "cert_status",
    "Certificate_Status",
    "Certificate Status",
]

WEBSITE_RAW_COL_CANDIDATES = [
    "Raw_Material",
    "Raw Material",
    "cert_in_put",
    "Input",
    "Inputs",
]

WEBSITE_PRODUCT_COL_CANDIDATES = [
    "Products",
    "Product",
    "cert_products",
]

FULL_FLOW_COLUMNS = [
    "Certificate_Number",
    "Certificate_File",
    "PDF_Page",
    "Table_Index",
    "PDF_Row_Number",
    "Input_Material",
    "Input_Scope",
    "Output_Material",
    "Output_Scope",
    "GHG_Option",
    "Raw_Material_Certification_Criteria",
    "Add_Ons",
    "Input_Material_Base",
    "Input_Material_Qualifier",
    "Output_Material_Base",
    "Output_Material_Qualifier",
    "Output_Product_Family",
    "Input_Is_Intermediate",
    "Extraction_Method",
]

SLIM_FLOW_COLUMNS = [
    "Certificate_Number",
    "Certificate_File",
    "PDF_Page",
    "Table_Index",
    "PDF_Row_Number",
    "Input_Material",
    "Input_Scope",
    "Output_Material",
    "Output_Scope",
    "GHG_Option",
    "Raw_Material_Certification_Criteria",
    "Add_Ons",
]

FAILURE_COLUMNS = [
    "Certificate_Number",
    "Certificate_File",
    "Failure_Stage",
    "Failure_Message",
]


@dataclass
class ExtractResult:
    flows: list[dict[str, Any]] = field(default_factory=list)
    failures: list[dict[str, Any]] = field(default_factory=list)


def clean_text(value: Any) -> str:
    """Clean Excel/PDF/OCR text into a single-line string."""
    if value is None:
        return ""
    s = str(value)
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = s.replace("\u00a0", " ").replace("\ufeff", " ").replace("\u00ad", "")
    s = s.replace("|", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_ocr_text(value: Any) -> str:
    """Light OCR repair for common ISCC table errors."""
    s = clean_text(value)
    if not s:
        return ""

    # Common OCR fixes seen on ISCC scanned certificates.
    replacements = {
        "Com / maize": "Corn / maize",
        "Com/ maize": "Corn / maize",
        "Com /maize": "Corn / maize",
        "Com maize": "Corn / maize",
        "Cor / maize": "Corn / maize",
        "Cor/ maize": "Corn / maize",
        "Cor /maize": "Corn / maize",
        "Biogas pliant": "Biogas plant",
        "Biogas piant": "Biogas plant",
        "Bliogas plant": "Biogas plant",
    }
    for bad, good in replacements.items():
        s = s.replace(bad, good)

    s = re.sub(r"(?<![A-Za-z])N\.?\s*A\.?(?![A-Za-z])", "N.A.", s, flags=re.IGNORECASE)

    # Normalize quote-like or bullet artifacts.
    s = s.strip(" .") if s.strip() not in {"N.A.", "n.a."} else s.strip()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_for_detection(value: Any) -> str:
    s = clean_text(value).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def parse_material(material: Any) -> tuple[str, str]:
    """
    Split a material such as 'Electricity (Corn / maize)' into:
        ('Electricity', 'Corn / maize')
    If no qualifier is present, returns ('material', '').
    """
    s = clean_text(material)
    m = re.match(r"^(.*?)\s*\((.*?)\)\s*$", s)
    if not m:
        return s, ""
    base = clean_text(m.group(1))
    qualifier = clean_text(m.group(2))
    return base, qualifier


def first_existing_column(df: pd.DataFrame, candidates: Iterable[str], required: bool = True) -> str | None:
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in lookup:
            return lookup[key]
    if required:
        raise KeyError(f"Could not find any of these columns: {list(candidates)}")
    return None


def read_input_table(input_path: Path, sheet_name: str | None = None) -> pd.DataFrame:
    suffix = input_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(input_path, sheet_name=sheet_name or 0)
    if suffix == ".csv":
        return pd.read_csv(input_path)
    raise ValueError("Input must be .xlsx, .xlsm, .xls, or .csv")


def safe_pdf_filename(cert_number: str, url: str) -> str:
    base = re.sub(r"[^A-Za-z0-9_.-]+", "_", clean_text(cert_number))[:120].strip("_")
    if not base:
        base = hashlib.sha1(url.encode("utf-8", errors="ignore")).hexdigest()[:16]
    return f"{base}.pdf"


def download_pdf(
    url: str,
    cert_number: str,
    cache_dir: Path,
    session: requests.Session,
    timeout: int = 90,
    verify_ssl: bool = True,
    force: bool = False,
) -> Path:
    if not url or not str(url).startswith(("http://", "https://")):
        raise ValueError("Certificate PDF URL is blank or invalid")

    cache_dir.mkdir(parents=True, exist_ok=True)
    out_path = cache_dir / safe_pdf_filename(cert_number, url)
    if out_path.exists() and out_path.stat().st_size > 1000 and not force:
        return out_path

    with session.get(url, stream=True, timeout=timeout, verify=verify_ssl) as resp:
        resp.raise_for_status()
        content_type = resp.headers.get("content-type", "").lower()
        if "pdf" not in content_type and not urlparse(url).path.lower().endswith(".pdf"):
            # Some ISCC handlers may not set content-type perfectly, so this is a warning-level check.
            pass
        tmp_path = out_path.with_suffix(".pdf.tmp")
        with tmp_path.open("wb") as f:
            for chunk in resp.iter_content(chunk_size=1024 * 128):
                if chunk:
                    f.write(chunk)
        tmp_path.replace(out_path)
    return out_path


def is_header_or_note_row(cells: list[str]) -> bool:
    joined = normalize_for_detection(" ".join(cells))
    if not joined:
        return True
    header_markers = [
        "input material",
        "output material",
        "material scope",
        "ghg option",
        "criteria of raw material",
        "add ons",
    ]
    if any(marker in joined for marker in header_markers):
        return True
    note_markers = [
        "please indicate",
        "default value",
        "actual value",
        "nuts2",
        "iscc eu add ons",
        "the raw material meets",
        "the raw material complies",
    ]
    if any(marker in joined for marker in note_markers):
        return True
    if cells and re.match(r"^(\*|1\)|2\)|3\))", cells[0].strip()):
        return True
    return False


def valid_flow_row(cells: list[str]) -> bool:
    if len(cells) < 4:
        return False
    if is_header_or_note_row(cells):
        return False
    input_material, input_scope, output_material, output_scope = cells[:4]
    if not input_material or not output_material:
        return False
    # Avoid rows that are just page/footer text.
    joined = normalize_for_detection(" ".join(cells))
    bad_markers = ["page ", "stamp", "signature", "certificate", "issuing certification body"]
    return not any(marker.strip() in joined for marker in bad_markers)


def row_to_flow(
    cells: list[str],
    cert_number: str,
    cert_url: str,
    page_number: int,
    table_index: int,
    row_number: int,
    method: str,
) -> dict[str, Any]:
    padded = (cells + [""] * 7)[:7]
    input_material, input_scope, output_material, output_scope, ghg, criteria, add_ons = padded
    input_base, input_qualifier = parse_material(input_material)
    output_base, output_qualifier = parse_material(output_material)
    return {
        "Certificate_Number": cert_number,
        "Certificate_File": cert_url,
        "PDF_Page": page_number,
        "Table_Index": table_index,
        "PDF_Row_Number": row_number,
        "Input_Material": input_material,
        "Input_Scope": input_scope,
        "Output_Material": output_material,
        "Output_Scope": output_scope,
        "GHG_Option": ghg,
        "Raw_Material_Certification_Criteria": criteria,
        "Add_Ons": add_ons,
        "Input_Material_Base": input_base,
        "Input_Material_Qualifier": input_qualifier,
        "Output_Material_Base": output_base,
        "Output_Material_Qualifier": output_qualifier,
        "Output_Product_Family": output_base,
        "Input_Is_Intermediate": False,  # filled later at certificate level
        "Extraction_Method": method,
    }


def extract_tables_pdfplumber(pdf_path: Path, cert_number: str, cert_url: str) -> list[dict[str, Any]]:
    """Extract table rows from digitally-readable PDFs."""
    flows: list[dict[str, Any]] = []
    settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 4,
        "join_tolerance": 4,
        "intersection_tolerance": 6,
        "edge_min_length": 20,
    }

    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page_idx, page in enumerate(pdf.pages, start=1):
                try:
                    tables = page.extract_tables(table_settings=settings) or []
                except Exception:
                    tables = []

                for table_idx, table in enumerate(tables, start=1):
                    joined_table = normalize_for_detection(" ".join(clean_text(c) for row in table for c in row))
                    if "input material" not in joined_table or "output material" not in joined_table:
                        continue

                    row_counter = 0
                    for raw_row in table:
                        cells = [clean_text(c) for c in raw_row]
                        cells = [c for c in cells if c != ""] if len(cells) > 7 else cells
                        if len(cells) < 7:
                            cells = (cells + [""] * 7)[:7]
                        else:
                            cells = cells[:7]
                        if not valid_flow_row(cells):
                            continue
                        row_counter += 1
                        flows.append(
                            row_to_flow(
                                cells=cells,
                                cert_number=cert_number,
                                cert_url=cert_url,
                                page_number=page_idx,
                                table_index=table_idx,
                                row_number=row_counter,
                                method="pdfplumber",
                            )
                        )
    except Exception:
        return []

    return flows


def render_pdf_page(page: fitz.Page, dpi: int) -> np.ndarray:
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        arr = cv2.cvtColor(arr, cv2.COLOR_RGBA2RGB)
    return cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)


def cluster_positions(values: list[int], tolerance: int = 8) -> list[int]:
    if not values:
        return []
    values = sorted(values)
    clusters: list[list[int]] = []
    for value in values:
        if not clusters or abs(clusters[-1][-1] - value) > tolerance:
            clusters.append([value])
        else:
            clusters[-1].append(value)
    return [int(round(sum(c) / len(c))) for c in clusters]


def detect_table_regions(image_bgr: np.ndarray) -> list[tuple[int, int, int, int]]:
    """Find large grid/table boxes in a rendered PDF page."""
    gray = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
    binary = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)[1]
    height, width = binary.shape

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(30, width // 40), 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(30, height // 40)))
    horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel, iterations=1)
    vertical = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel, iterations=1)
    lines = cv2.bitwise_or(horizontal, vertical)

    contours, _ = cv2.findContours(lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    regions: list[tuple[int, int, int, int]] = []
    page_area = width * height
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        area = w * h
        if w > width * 0.45 and h > height * 0.08 and area > page_area * 0.03:
            regions.append((x, y, w, h))

    # Deduplicate/keep largest first.
    regions = sorted(regions, key=lambda r: r[2] * r[3], reverse=True)
    unique: list[tuple[int, int, int, int]] = []
    for r in regions:
        x, y, w, h = r
        overlaps = False
        for ux, uy, uw, uh in unique:
            ix0 = max(x, ux)
            iy0 = max(y, uy)
            ix1 = min(x + w, ux + uw)
            iy1 = min(y + h, uy + uh)
            if ix1 > ix0 and iy1 > iy0:
                inter = (ix1 - ix0) * (iy1 - iy0)
                if inter / min(w * h, uw * uh) > 0.8:
                    overlaps = True
                    break
        if not overlaps:
            unique.append(r)
    return unique


def get_grid_lines(image_bgr: np.ndarray, region: tuple[int, int, int, int]) -> tuple[list[int], list[int], np.ndarray]:
    """Return x grid lines, y grid lines, and detected line mask."""
    gray = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
    binary = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)[1]
    height, width = binary.shape
    table_x, table_y, table_w, table_h = region

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(30, width // 40), 1))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(30, height // 40)))
    horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel, iterations=1)
    vertical = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vertical_kernel, iterations=1)
    line_mask = cv2.bitwise_or(horizontal, vertical)

    h_contours, _ = cv2.findContours(horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    y_values: list[int] = []
    for contour in h_contours:
        x, y, w, h = cv2.boundingRect(contour)
        if (
            y >= table_y - 15
            and y <= table_y + table_h + 15
            and x >= table_x - 30
            and x + w <= table_x + table_w + 30
            and w > table_w * 0.50
        ):
            y_values.append(int(round(y + h / 2)))

    y_lines = cluster_positions(y_values, tolerance=10)

    # For x lines, require that the vertical segment overlaps the material-flow table body.
    # Standard ISCC annex tables have two header rows, so the data grid starts at y_lines[2].
    body_start = y_lines[2] if len(y_lines) >= 3 else table_y
    body_end = y_lines[-1] if y_lines else table_y + table_h

    v_contours, _ = cv2.findContours(vertical, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    x_values: list[int] = []
    for contour in v_contours:
        x, y, w, h = cv2.boundingRect(contour)
        overlaps_body = y <= body_start + 25 and y + h >= body_start + 200
        in_region = x >= table_x - 20 and x <= table_x + table_w + 20
        tall_enough = h > max(150, table_h * 0.20)
        if in_region and overlaps_body and tall_enough:
            x_values.append(int(round(x + w / 2)))

    x_lines = cluster_positions(x_values, tolerance=10)

    # Relaxed fallback for shallow/single-row annex tables.
    # Some ISCC annexes have only one material-flow row followed immediately by footnotes.
    # In those layouts the vertical cell borders are short, so the strict body_start + 200
    # condition above can miss all x boundaries. If that happens, look only at the top
    # material-flow grid and accept shorter vertical borders.
    if len(x_lines) < 8 and len(y_lines) >= 3:
        relaxed_values: list[int] = []
        top_grid_start = y_lines[0]
        body_start_relaxed = y_lines[2]
        top_grid_end = y_lines[3] if len(y_lines) >= 4 else min(table_y + table_h, body_start_relaxed + 160)
        min_relaxed_height = max(35, int((top_grid_end - top_grid_start) * 0.35))

        for contour in v_contours:
            x, y, w, h = cv2.boundingRect(contour)
            in_region = x >= table_x - 25 and x <= table_x + table_w + 25
            overlaps_material_grid = y <= body_start_relaxed + 35 and y + h >= body_start_relaxed + 25
            within_top_grid = y < top_grid_end + 45 and y + h > top_grid_start - 45
            tall_enough = h >= min_relaxed_height
            if in_region and overlaps_material_grid and within_top_grid and tall_enough:
                relaxed_values.append(int(round(x + w / 2)))

        relaxed_x_lines = cluster_positions(relaxed_values, tolerance=12)
        if len(relaxed_x_lines) >= len(x_lines):
            x_lines = relaxed_x_lines

    # Add table edges if they were missed by line detection. This helps scans where the
    # left/right borders are faint but the internal column dividers are visible.
    if 6 <= len(x_lines) < 8:
        edge_candidates = [table_x, *x_lines, table_x + table_w]
        edge_candidates = cluster_positions(edge_candidates, tolerance=15)
        if len(edge_candidates) >= 8:
            x_lines = edge_candidates

    # Keep the 7-column material-flow grid when extra footnote columns are detected.
    # Expected x lines are 8 boundaries: input material, input scope, output material,
    # output scope, ghg, criteria, add-ons.
    if len(x_lines) > 8:
        # Prefer the 8-line span that covers most of the detected table width.
        best = x_lines[:8]
        best_score = -1
        for i in range(0, len(x_lines) - 7):
            candidate = x_lines[i : i + 8]
            span = candidate[-1] - candidate[0]
            left_penalty = abs(candidate[0] - table_x)
            right_penalty = abs(candidate[-1] - (table_x + table_w))
            score = span - 0.25 * (left_penalty + right_penalty)
            if score > best_score:
                best = candidate
                best_score = score
        x_lines = best

    return x_lines, y_lines, line_mask


def ocr_cell(image_bgr: np.ndarray, line_mask: np.ndarray, box: tuple[int, int, int, int]) -> str:
    if pytesseract is None:
        raise RuntimeError("pytesseract is not installed. Install it or use digitally-readable PDFs only.")

    x1, y1, x2, y2 = box
    pad_x = max(4, int((x2 - x1) * 0.015))
    pad_y = max(4, int((y2 - y1) * 0.05))
    x1 = max(0, x1 + pad_x)
    y1 = max(0, y1 + pad_y)
    x2 = min(image_bgr.shape[1], x2 - pad_x)
    y2 = min(image_bgr.shape[0], y2 - pad_y)
    if x2 <= x1 or y2 <= y1:
        return ""

    clean = image_bgr.copy()
    expanded_mask = cv2.dilate(line_mask, cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3)), iterations=1)
    clean[expanded_mask > 0] = (255, 255, 255)

    crop = clean[y1:y2, x1:x2]
    crop = cv2.copyMakeBorder(crop, 10, 10, 10, 10, cv2.BORDER_CONSTANT, value=(255, 255, 255))
    crop = cv2.resize(crop, None, fx=1.7, fy=1.7, interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
    thresholded = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    text = pytesseract.image_to_string(thresholded, config="--oem 3 --psm 6")
    return clean_ocr_text(text)


def extract_tables_ocr(
    pdf_path: Path,
    cert_number: str,
    cert_url: str,
    dpi: int = 300,
    max_pages: int | None = None,
) -> list[dict[str, Any]]:
    """Extract table rows from scanned/image PDFs using grid detection + OCR."""
    if pytesseract is None:
        return []

    flows: list[dict[str, Any]] = []
    doc = fitz.open(str(pdf_path))
    try:
        page_count = len(doc) if max_pages is None else min(len(doc), max_pages)
        for page_idx in range(page_count):
            page = doc.load_page(page_idx)
            image = render_pdf_page(page, dpi=dpi)
            regions = detect_table_regions(image)
            for table_idx, region in enumerate(regions, start=1):
                x_lines, y_lines, line_mask = get_grid_lines(image, region)
                if len(x_lines) < 8 or len(y_lines) < 4:
                    continue

                # Keep exactly 8 x boundaries if a table has extra columns from notes/footers.
                x_lines = x_lines[:8]

                row_counter = 0
                # Standard ISCC annex table has two header intervals:
                #   y0-y1: Input material / Output material grouping
                #   y1-y2: Material / Scope headers
                for row_idx in range(2, len(y_lines) - 1):
                    y1, y2 = y_lines[row_idx], y_lines[row_idx + 1]
                    if y2 - y1 < 20:
                        continue

                    cells: list[str] = []
                    for col_idx in range(7):
                        x1, x2 = x_lines[col_idx], x_lines[col_idx + 1]
                        cells.append(ocr_cell(image, line_mask, (x1, y1, x2, y2)))

                    joined = " ".join(cells)
                    if is_header_or_note_row(cells):
                        # Once footnotes start, the material table is over.
                        if re.search(r"please indicate|default value|actual value|nuts2|iscc eu", normalize_for_detection(joined)):
                            break
                        continue
                    if not valid_flow_row(cells):
                        continue

                    row_counter += 1
                    flows.append(
                        row_to_flow(
                            cells=cells,
                            cert_number=cert_number,
                            cert_url=cert_url,
                            page_number=page_idx + 1,
                            table_index=table_idx,
                            row_number=row_counter,
                            method="ocr_grid",
                        )
                    )
    finally:
        doc.close()
    return flows


def mark_intermediate_inputs(flows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Flag inputs that are likely intermediate products created elsewhere in the same certificate."""
    by_cert: dict[str, list[dict[str, Any]]] = {}
    for row in flows:
        by_cert.setdefault(str(row.get("Certificate_Number", "")), []).append(row)

    for cert_rows in by_cert.values():
        output_bases = {normalize_for_detection(r.get("Output_Material_Base", "")) for r in cert_rows}
        output_bases.discard("")
        for row in cert_rows:
            input_base = normalize_for_detection(row.get("Input_Material_Base", ""))
            row["Input_Is_Intermediate"] = bool(input_base and input_base in output_bases)
    return flows


def extract_one_certificate(
    cert_number: str,
    cert_url: str,
    cache_dir: Path,
    timeout: int,
    verify_ssl: bool,
    force_download: bool,
    dpi: int,
    max_pages: int | None,
    prefer_ocr: bool,
) -> ExtractResult:
    result = ExtractResult()
    session = requests.Session()
    session.headers.update(DEFAULT_HEADERS)

    try:
        pdf_path = download_pdf(
            url=cert_url,
            cert_number=cert_number,
            cache_dir=cache_dir,
            session=session,
            timeout=timeout,
            verify_ssl=verify_ssl,
            force=force_download,
        )
    except Exception as exc:
        result.failures.append(
            {
                "Certificate_Number": cert_number,
                "Certificate_File": cert_url,
                "Failure_Stage": "download",
                "Failure_Message": str(exc),
            }
        )
        return result

    try:
        flows: list[dict[str, Any]] = []
        if not prefer_ocr:
            flows = extract_tables_pdfplumber(pdf_path, cert_number, cert_url)
        if not flows:
            flows = extract_tables_ocr(pdf_path, cert_number, cert_url, dpi=dpi, max_pages=max_pages)

        if not flows:
            result.failures.append(
                {
                    "Certificate_Number": cert_number,
                    "Certificate_File": cert_url,
                    "Failure_Stage": "parse",
                    "Failure_Message": "No annex material-flow table detected or extracted",
                }
            )
        else:
            result.flows.extend(flows)
    except Exception as exc:
        result.failures.append(
            {
                "Certificate_Number": cert_number,
                "Certificate_File": cert_url,
                "Failure_Stage": "parse",
                "Failure_Message": str(exc),
            }
        )
    return result


def unique_join(values: Iterable[Any]) -> str:
    seen: set[str] = set()
    out: list[str] = []
    for value in values:
        s = clean_text(value)
        key = normalize_for_detection(s)
        if s and key not in seen:
            seen.add(key)
            out.append(s)
    return "; ".join(out)


def build_summary(flows_df: pd.DataFrame, source_df: pd.DataFrame, cert_col: str, url_col: str) -> pd.DataFrame:
    if flows_df.empty:
        return pd.DataFrame(
            columns=[
                "Certificate_Number",
                "Certificate_File",
                "PDF_Row_Count",
                "PDF_Input_Materials_All",
                "PDF_Raw_Materials_Likely",
                "PDF_Output_Materials_All",
                "PDF_Output_Product_Families",
                "Extraction_Methods",
            ]
        )

    rows: list[dict[str, Any]] = []
    for cert_number, group in flows_df.groupby("Certificate_Number", dropna=False):
        raw_likely = group.loc[~group["Input_Is_Intermediate"].astype(bool), "Input_Material"]
        rows.append(
            {
                "Certificate_Number": cert_number,
                "Certificate_File": unique_join(group["Certificate_File"]),
                "PDF_Row_Count": len(group),
                "PDF_Input_Materials_All": unique_join(group["Input_Material"]),
                "PDF_Raw_Materials_Likely": unique_join(raw_likely),
                "PDF_Output_Materials_All": unique_join(group["Output_Material"]),
                "PDF_Output_Product_Families": unique_join(group["Output_Product_Family"]),
                "Extraction_Methods": unique_join(group["Extraction_Method"]),
            }
        )
    summary = pd.DataFrame(rows)

    raw_col = first_existing_column(source_df, WEBSITE_RAW_COL_CANDIDATES, required=False)
    product_col = first_existing_column(source_df, WEBSITE_PRODUCT_COL_CANDIDATES, required=False)
    cols_to_merge = [cert_col]
    rename = {cert_col: "Certificate_Number"}
    if url_col not in cols_to_merge:
        cols_to_merge.append(url_col)
        rename[url_col] = "Certificate_File_Source"
    if raw_col:
        cols_to_merge.append(raw_col)
        rename[raw_col] = "Raw_Material_Website"
    if product_col:
        cols_to_merge.append(product_col)
        rename[product_col] = "Products_Website"

    source_small = source_df[cols_to_merge].copy().rename(columns=rename)
    source_small["Certificate_Number"] = source_small["Certificate_Number"].astype(str).map(clean_text)
    summary["Certificate_Number"] = summary["Certificate_Number"].astype(str).map(clean_text)
    summary = source_small.drop_duplicates("Certificate_Number").merge(summary, on="Certificate_Number", how="right")
    return summary


def write_output(
    output_path: Path,
    flows: list[dict[str, Any]],
    failures: list[dict[str, Any]],
    source_df: pd.DataFrame,
    cert_col: str,
    url_col: str,
    include_derived_cols: bool = False,
) -> None:
    """Write output workbook.

    The detailed derived columns are still calculated internally for the summary sheet,
    but the PDF Material Flows sheet is slim by default. Use --include-derived-cols
    if you want the old expanded columns back.
    """
    flows = mark_intermediate_inputs(flows)
    flows_full_df = pd.DataFrame(flows)
    if flows_full_df.empty:
        flows_full_df = pd.DataFrame(columns=FULL_FLOW_COLUMNS)
    else:
        for col in FULL_FLOW_COLUMNS:
            if col not in flows_full_df.columns:
                flows_full_df[col] = ""
        flows_full_df = flows_full_df[FULL_FLOW_COLUMNS]

    # Build summary from the full internal frame, because it needs the derived
    # Output_Product_Family and Input_Is_Intermediate columns.
    summary_df = build_summary(flows_full_df, source_df, cert_col, url_col)

    output_flow_columns = FULL_FLOW_COLUMNS if include_derived_cols else SLIM_FLOW_COLUMNS
    flows_output_df = flows_full_df[output_flow_columns].copy()

    failures_df = pd.DataFrame(failures)
    if failures_df.empty:
        failures_df = pd.DataFrame(columns=FAILURE_COLUMNS)
    else:
        for col in FAILURE_COLUMNS:
            if col not in failures_df.columns:
                failures_df[col] = ""
        failures_df = failures_df[FAILURE_COLUMNS]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        flows_output_df.to_excel(writer, index=False, sheet_name="PDF Material Flows")
        summary_df.to_excel(writer, index=False, sheet_name="PDF Certificate Summary")
        failures_df.to_excel(writer, index=False, sheet_name="PDF Extraction Failures")


def parse_bool(value: str) -> bool:
    value = str(value).strip().lower()
    if value in {"1", "true", "yes", "y"}:
        return True
    if value in {"0", "false", "no", "n"}:
        return False
    raise argparse.ArgumentTypeError("Expected true/false")


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract ISCC PDF annex material-flow tables.")
    parser.add_argument("--input", required=True, help="Input ISCC certificate export: .xlsx, .xlsm, .xls, or .csv")
    parser.add_argument("--output", required=True, help="Output .xlsx workbook")
    parser.add_argument("--sheet", default=None, help="Excel sheet name/index. Defaults to first sheet.")
    parser.add_argument("--cert-col", default=None, help="Certificate number column. Auto-detected by default.")
    parser.add_argument("--url-col", default=None, help="Certificate PDF URL column. Auto-detected by default.")
    parser.add_argument("--status", default=None, help="Optional status filter, e.g. Valid. Case-insensitive.")
    parser.add_argument("--status-col", default=None, help="Status column. Auto-detected by default when --status is used.")
    parser.add_argument("--certificate", default=None, help="Optional exact Certificate_ID filter for testing one certificate. Case-insensitive.")
    parser.add_argument("--cache-dir", default="pdf_cache", help="Folder for downloaded PDFs")
    parser.add_argument("--workers", type=int, default=2, help="Parallel downloads/extractions. Use 1-4 to be polite.")
    parser.add_argument("--delay", type=float, default=0.0, help="Delay between submissions in seconds")
    parser.add_argument("--timeout", type=int, default=90, help="HTTP timeout in seconds")
    parser.add_argument("--verify-ssl", type=parse_bool, default=True, help="true/false. Use false only if your environment has SSL issues.")
    parser.add_argument("--force-download", action="store_true", help="Re-download PDFs even if cached")
    parser.add_argument("--dpi", type=int, default=300, help="OCR render DPI for scanned PDFs")
    parser.add_argument("--max-pages", type=int, default=None, help="Optional max pages per PDF")
    parser.add_argument("--limit", type=int, default=None, help="Optional number of certificates to process")
    parser.add_argument("--prefer-ocr", action="store_true", help="Skip pdfplumber and run OCR/grid extraction first")
    parser.add_argument("--tesseract-cmd", default=None, help="Path to tesseract executable on Windows, if needed")
    parser.add_argument("--include-derived-cols", action="store_true", help="Include base/qualifier/intermediate/extraction-method columns in PDF Material Flows. By default these are hidden.")
    args = parser.parse_args()

    if args.tesseract_cmd:
        if pytesseract is None:
            raise RuntimeError("pytesseract is not installed")
        pytesseract.pytesseract.tesseract_cmd = args.tesseract_cmd

    input_path = Path(args.input)
    output_path = Path(args.output)
    cache_dir = Path(args.cache_dir)

    df = read_input_table(input_path, sheet_name=args.sheet)
    cert_col = args.cert_col or first_existing_column(df, CERT_NUMBER_COL_CANDIDATES, required=True)
    url_col = args.url_col or first_existing_column(df, CERT_URL_COL_CANDIDATES, required=True)

    if args.status:
        status_col = args.status_col or first_existing_column(df, STATUS_COL_CANDIDATES, required=True)
        before_status_filter = len(df)
        wanted_status = clean_text(args.status).lower()
        df = df[df[status_col].astype(str).map(clean_text).str.lower() == wanted_status].copy()
        print(f"Status filter: {status_col} == {args.status} | rows kept: {len(df)}/{before_status_filter}")

    if args.certificate:
        before_cert_filter = len(df)
        wanted_cert = clean_text(args.certificate).lower()
        df = df[df[cert_col].astype(str).map(clean_text).str.lower() == wanted_cert].copy()
        print(f"Certificate filter: {cert_col} == {args.certificate} | rows kept: {len(df)}/{before_cert_filter}")

    work = df[[cert_col, url_col]].copy()
    work[cert_col] = work[cert_col].astype(str).map(clean_text)
    work[url_col] = work[url_col].astype(str).map(clean_text)
    work = work[work[url_col].str.startswith(("http://", "https://"), na=False)]
    work = work.drop_duplicates(subset=[cert_col, url_col])
    if args.limit:
        work = work.head(args.limit)

    print(f"Input rows: {len(df)}")
    print(f"Certificates with PDF links to process: {len(work)}")
    print(f"Certificate column: {cert_col}")
    print(f"PDF URL column: {url_col}")

    all_flows: list[dict[str, Any]] = []
    all_failures: list[dict[str, Any]] = []

    tasks = []
    with futures.ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        for _, row in work.iterrows():
            cert_number = clean_text(row[cert_col])
            cert_url = clean_text(row[url_col])
            task = executor.submit(
                extract_one_certificate,
                cert_number,
                cert_url,
                cache_dir,
                args.timeout,
                args.verify_ssl,
                args.force_download,
                args.dpi,
                args.max_pages,
                args.prefer_ocr,
            )
            tasks.append(task)
            if args.delay:
                time.sleep(args.delay)

        done = 0
        for task in futures.as_completed(tasks):
            done += 1
            result = task.result()
            all_flows.extend(result.flows)
            all_failures.extend(result.failures)
            if done % 10 == 0 or done == len(tasks):
                print(f"Processed {done}/{len(tasks)} | rows extracted: {len(all_flows)} | failures: {len(all_failures)}")

    write_output(output_path, all_flows, all_failures, df, cert_col, url_col, include_derived_cols=args.include_derived_cols)
    print(f"Saved output: {output_path}")


if __name__ == "__main__":
    main()
