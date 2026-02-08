import os
import pandas as pd
import unicodedata
import re


# Constants (editable if needed)
EXCEL_ENGINE = "openpyxl"
DEFAULT_SHEET = "Certificate Database"

def find_id_column(df: pd.DataFrame) -> str:
    """
    Detect the certificate ID column.
    Prefers 'Certificate_ID', then 'Certificate ID', then fuzzy search.
    """
    cols = list(df.columns)
    if "Certificate_ID" in cols:
        return "Certificate_ID"
    if "Certificate ID" in cols:
        return "Certificate ID"
    for c in cols:
        cn = c.strip().lower().replace(" ", "_")
        if "cert" in cn and "id" in cn:
            return c
    raise KeyError(
        "Could not find a certificate ID column. "
        "Expected 'Certificate_ID' or 'Certificate ID'."
    )

def load_sheet(path: str, sheet_name: str = DEFAULT_SHEET) -> pd.DataFrame:
        """Read a sheet with all columns as strings for consistent comparison."""
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")
        return pd.read_excel(path, sheet_name=sheet_name, engine=EXCEL_ENGINE, dtype=str)


# Normalize to strings (trim whitespace) for reliable comparisons
def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in df.columns:
        df[c] = df[c].fillna("").astype(str).str.strip()
    return df


def create_certs_added(previous_fn, current_fn): 

    current_df = load_sheet(current_fn, sheet_name=DEFAULT_SHEET)
    current_df = _normalize(current_df)
    id_col = find_id_column(current_df)
    current_ids = set(current_df[id_col][current_df[id_col] != ""])

    # Try previous snapshot
    prev_path = previous_fn
    if not prev_path or not os.path.exists(prev_path):
        # First run or prev file missing: produce empty "added" report
        added_df = current_df.iloc[0:0].copy()
    else:
        prev_df = load_sheet(prev_path, sheet_name=DEFAULT_SHEET)
        prev_df = _normalize(prev_df)
        prev_id_col = find_id_column(prev_df)
        prev_ids = set(prev_df[prev_id_col][prev_df[prev_id_col] != ""])

        # New IDs present in current, not in previous
        new_ids = current_ids - prev_ids
        added_df = current_df[current_df[id_col].isin(new_ids)].copy()

    # Append as new sheet to current workbook
    # NOTE: Excel must be closed to avoid PermissionError.
    try:
        with pd.ExcelWriter(current_fn, engine=EXCEL_ENGINE, mode="a", if_sheet_exists="new") as writer:
            new_sheet_name = "Certificates Added"
            added_df.to_excel(writer, sheet_name=new_sheet_name, index=False)
    except PermissionError as e:
        raise PermissionError(
            f"Could not write to '{current_fn}'. Is it open in Excel/OneDrive? "
            "Close it and retry."
        ) from e
    
    # Simple console summary (optional)
    print(f"Compared against: {prev_path or '(no previous snapshot)'}")
    print(f"Added certificates found: {len(added_df)}")


def create_certs_removed(previous_fn, current_fn):
  
    current_df = load_sheet(current_fn, sheet_name=DEFAULT_SHEET)
    current_df = _normalize(current_df)
    id_col = find_id_column(current_df)
    current_ids = set(current_df[id_col][current_df[id_col] != ""])

    # Load previous snapshot
    prev_path = previous_fn
    if not prev_path or not os.path.exists(prev_path):
        # No previous file: nothing can be "removed"
        removed_df = current_df.iloc[0:0].copy()
    else:
        prev_df = load_sheet(prev_path, sheet_name=DEFAULT_SHEET)
        prev_df = _normalize(prev_df)
        prev_id_col = find_id_column(prev_df)
        prev_ids = set(prev_df[prev_id_col][prev_df[prev_id_col] != ""])

        # Removed IDs: present in previous, not in current
        removed_ids = prev_ids - current_ids
        removed_df = prev_df[prev_df[prev_id_col].isin(removed_ids)].copy()

    # Append as new sheet to current workbook
    # NOTE: Excel must be closed to avoid PermissionError.
    try:
        with pd.ExcelWriter(current_fn, engine=EXCEL_ENGINE, mode="a", if_sheet_exists="new") as writer:
            new_sheet_name = "Certificates Removed"
            removed_df.to_excel(writer, sheet_name=new_sheet_name, index=False)
    except PermissionError as e:
        raise PermissionError(
            f"Could not write to '{current_fn}'. Is it open in Excel/OneDrive? "
            "Close it and retry."
        ) from e

    # Simple console summary (optional)
    print(f"Compared against: {prev_path or '(no previous snapshot)'}")
    print(f"Removed certificates found: {len(removed_df)}")


def create_certs_changed(
    previous_fn,
    current_fn,
    sheet_name="Certificate Database",
    ignore_cols=None,
    ignore_patterns=None,
    case_insensitive=True,
    # enabled by default as agreed
    apply_location_suffix_rule=True,
    # You can add more suffixes here if you encounter them
    location_suffix_patterns=None,
    # Global text equivalence rules (pattern, replacement)
    custom_equivalence_rules=None
):
    """
    Compare two snapshots of the 'Certificate Database' and return rows whose values changed.

    Fixes false positives by:
      • Unicode normalization (NFKC)
      • Stripping invisible characters (ZWSP, NBSP, soft hyphen, etc.)
      • Collapsing whitespace
      • Optional case-insensitive compare (default: True)
      • Optional rules to strip location suffixes (e.g., ', Budapest, Hungary')
      • Optional custom equivalence rules
      • Deduping by ID before comparison
      • Ignoring specified columns and/or regex patterns
    Writes directly to 'output_path' (no temp file).
    """

    # -------- helpers --------
    INVISIBLE_CHARS = (
        "\u200b"  # zero width space
        "\u200c"  # zero width non-joiner
        "\u200d"  # zero width joiner
        "\ufeff"  # BOM
        "\u00a0"  # non-breaking space
        "\u2060"  # word joiner
        "\u202f"  # narrow no-break space
        "\u00ad"  # soft hyphen
    )
    INVISIBLES_RE = re.compile(f"[{re.escape(INVISIBLE_CHARS)}]")

    if location_suffix_patterns is None:
        location_suffix_patterns = [
            r",\s*Budapest,\s*Hungary\s*$",   # the case you hit
            # Add more as needed, e.g.:
            # r",\s*London,\s*UK\s*$",
            # r",\s*Singapore\s*$",
        ]

    def sanitize_text(s: pd.Series) -> pd.Series:
        # 1) Unicode normalize
        s = s.astype(str).map(lambda x: unicodedata.normalize("NFKC", x))
        # 2) Remove invisible chars
        s = s.map(lambda x: INVISIBLES_RE.sub("", x))
        # 3) Trim and collapse whitespace
        s = s.str.strip().str.replace(r"\s+", " ", regex=True)
        return s

    def apply_equivalence_rules(s: pd.Series) -> pd.Series:
        # global rules first
        if custom_equivalence_rules:
            for pat, repl in custom_equivalence_rules:
                s = s.str.replace(pat, repl, regex=True)
        # then suffix rules (end-of-string only)
        if apply_location_suffix_rule and location_suffix_patterns:
            for pat in location_suffix_patterns:
                s = s.str.replace(pat, "", regex=True)
        return s

    def normalize_cols(df):
        df.columns = df.columns.str.strip()
        return df

    def normalize_values(df, cols):
        out = df[cols].copy()
        for c in cols:
            s = out[c]
            if pd.api.types.is_datetime64_any_dtype(s):
                out[c] = s.dt.strftime("%Y-%m-%d %H:%M:%S").fillna("")
            else:
                s = sanitize_text(s).fillna("")
                if case_insensitive:
                    s = s.str.lower()
                s = apply_equivalence_rules(s)
                out[c] = s
        return out

    def drop_volatile_columns(cols):
        cols_set = set(cols)
        if ignore_cols:
            cols_set -= set(ignore_cols)
        if ignore_patterns:
            for pat in ignore_patterns:
                to_drop = {c for c in cols_set if re.search(pat, c, flags=re.IGNORECASE)}
                cols_set -= to_drop
        return list(cols_set)

    # -------- load --------
    current_df = load_sheet(current_fn, sheet_name=sheet_name)
    prev_df    = load_sheet(previous_fn, sheet_name=sheet_name)

    current_df = normalize_cols(current_df).fillna("")
    prev_df    = normalize_cols(prev_df).fillna("")

    id_col_curr = find_id_column(current_df)
    id_col_prev = find_id_column(prev_df)

    if prev_df is None or prev_df.empty:
        empty_cols = list(current_df.columns) + ["Value_Changed"]
        changed_df = pd.DataFrame(columns=empty_cols)
        # Write empty sheet (optional but keeps workflow consistent)
        try:
            with pd.ExcelWriter(current_fn, engine=EXCEL_ENGINE, mode="a", if_sheet_exists="new") as writer:
                changed_df.to_excel(writer, sheet_name="Certificates Changed", index=False)
        except PermissionError as e:
            raise PermissionError(
                f"Could not write to '{current_fn}'. Is it open in Excel/OneDrive? Close it and retry."
            ) from e
        print(f"Compared against: {previous_fn or '(no previous snapshot)'}")
        print("No previous snapshot found (or empty). No change rows generated.")
        return changed_df

    shared_cols = set(current_df.columns).intersection(prev_df.columns)
    shared_cols.discard(id_col_curr)
    shared_cols.discard(id_col_prev)
    shared_cols = drop_volatile_columns(list(shared_cols))
    if not shared_cols:
        raise ValueError("No comparable columns found after applying ignores.")

    # --- index & dedup by ID (take first) ---
    current_idx = (current_df.set_index(id_col_curr, drop=False)
                             .groupby(level=0, as_index=True).first())
    prev_idx    = (prev_df.set_index(id_col_prev, drop=False)
                           .groupby(level=0, as_index=True).first())

    # --- compare only the common IDs ---
    common_ids = current_idx.index.intersection(prev_idx.index)
    curr_common = current_idx.loc[common_ids]
    prev_common = prev_idx.loc[common_ids]

    # --- normalize, compare (vectorized) ---
    curr_norm = normalize_values(curr_common, shared_cols)
    prev_norm = normalize_values(prev_common, shared_cols)

    diff_mask = curr_norm.ne(prev_norm)
    changed_ids = diff_mask.any(axis=1)
    changed_ids = changed_ids[changed_ids].index

    changed_cols_list = diff_mask.loc[changed_ids].apply(
        lambda row: ", ".join([col for col, changed in row.items() if changed]),
        axis=1
    )

    # assemble from current snapshot + annotation
    changed_df = curr_common.loc[changed_ids].copy()
    changed_df["Value_Changed"] = changed_cols_list

    base_cols = list(current_df.columns)
    if "Value_Changed" not in base_cols:
        base_cols.append("Value_Changed")
    changed_df = changed_df[base_cols]

    print(f"Changed certificates found: {len(changed_df)}")

    # Append as new sheet to the current workbook
    try:
        with pd.ExcelWriter(current_fn, engine=EXCEL_ENGINE, mode="a", if_sheet_exists="new") as writer:
            changed_df.to_excel(writer, sheet_name="Certificates Changed", index=False)
    except PermissionError as e:
        raise PermissionError(
            f"Could not write to '{current_fn}'. Is it open in Excel/OneDrive? Close it and retry."
        ) from e

    # Summary (optional)
    print(f"Compared against: {previous_fn or '(no previous snapshot)'}")
    print(f"Changed certificates found: {len(changed_df)}")

