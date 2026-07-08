from __future__ import annotations

import argparse
from copy import copy
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup


URL = "https://iscc-system.org/certification/excluded-system-users/"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,en-GB;q=0.8",
    "Referer": "https://iscc-system.org/certification/list-of-non-compliant-points-of-origin-poo/",
}

EXPECTED_COLUMNS = [
    "Company name",
    "Address",
    "Excluded from",
    "Excluded until",
]


def default_output_filename() -> str:
    """
    Creates a filename like:
    ISCC_Certificates_01.07.2026_13.46.xlsx
    """
    return f"ISCC_Certificates_{datetime.now():%d.%m.%Y_%H.%M}.xlsx"


def clean_text(value: str) -> str:
    """
    Removes extra spaces, line breaks, and non-breaking spaces.
    """
    return " ".join(value.replace("\xa0", " ").split())


def fetch_html(url: str = URL, timeout: int = 30) -> str:
    response = requests.get(url, headers=HEADERS, timeout=timeout, verify=False)
    response.raise_for_status()
    return response.text


def find_excluded_table(soup: BeautifulSoup):
    """
    Finds the table containing the expected ISCC excluded-user headers.
    """
    for table in soup.find_all("table"):
        cells = table.find_all(["th", "td"])
        cell_text = [clean_text(cell.get_text(" ")) for cell in cells[:20]]
        joined_text = " | ".join(cell_text).lower()

        if all(column.lower() in joined_text for column in EXPECTED_COLUMNS):
            return table

    raise RuntimeError("Could not find the excluded system users table on the page.")


def parse_rows(table) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []

    for tr in table.find_all("tr"):
        cells = [clean_text(cell.get_text(" ")) for cell in tr.find_all(["td", "th"])]

        if not cells:
            continue

        # Skip header row
        if cells[:4] == EXPECTED_COLUMNS:
            continue

        # Skip incomplete rows
        if len(cells) < 4:
            continue

        rows.append(
            {
                "Company name": cells[0],
                "Address": cells[1],
                "Excluded from": cells[2],
                "Excluded until": cells[3],
            }
        )

    return rows


def scrape_excluded_system_users(html: str) -> pd.DataFrame:
    soup = BeautifulSoup(html, "lxml")
    table = find_excluded_table(soup)
    rows = parse_rows(table)

    if not rows:
        raise RuntimeError("The table was found, but no data rows were parsed.")

    return pd.DataFrame(rows, columns=EXPECTED_COLUMNS)


def save_excel(df: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Excluded System Users")

        worksheet = writer.book["Excluded System Users"]

        # Freeze header row
        worksheet.freeze_panes = "A2"

        # Basic column widths
        column_widths = {
            "A": 42,
            "B": 90,
            "C": 16,
            "D": 16,
        }

        for column, width in column_widths.items():
            worksheet.column_dimensions[column].width = width

        # Wrap text so addresses are readable
        for row in worksheet.iter_rows():
            for cell in row:
                alignment = copy(cell.alignment)
                alignment.wrap_text = True
                alignment.vertical = "top"
                cell.alignment = alignment

        # Add filter to the header row
        worksheet.auto_filter.ref = worksheet.dimensions


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Scrape ISCC excluded system users into Excel."
    )

    parser.add_argument(
        "--url",
        default=URL,
        help="ISCC excluded system users URL",
    )

    parser.add_argument(
        "--output",
        default=None,
        help=(
            "Optional Excel output filename. "
            "If omitted, the script creates a file like "
            "ISCC_Certificates_01.07.2026_13.46.xlsx"
        ),
    )

    parser.add_argument(
        "--html-file",
        default=None,
        help="Optional local HTML file to parse instead of downloading the page",
    )

    args = parser.parse_args()

    if args.html_file:
        html = Path(args.html_file).read_text(encoding="utf-8")
    else:
        html = fetch_html(args.url)

    output_path = Path(args.output) if args.output else Path(default_output_filename())

    df = scrape_excluded_system_users(html)
    save_excel(df, output_path)

    print(f"Saved {len(df):,} rows to {output_path}")


if __name__ == "__main__":
    main()