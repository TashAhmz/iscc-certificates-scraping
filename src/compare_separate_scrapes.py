from compare import find_id_column, load_sheet, _normalize, DEFAULT_SHEET
import os
from pathlib import Path
import pandas as pd
import unicodedata
import re


def get_certs_added(previous_fn, current_fn):

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
    
    # Simple console summary (optional)
    print(f"Added certificates found: {len(added_df)}")

def get_certs_removed(previous_fn, current_fn):
   
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

    # Simple console summary (optional)
    print()
    print(f"Removed certificates found: {len(removed_df)}")


def get_certs_changed(
    previous_fn,
    current_fn,
    sheet_name="Certificate Database",
    ignore_cols=None,
    ignore_patterns=None,
    case_insensitive=True,
    output_path="certs_changed.xlsx",
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

    # --- write directly (no tmp) ---
    with pd.ExcelWriter(output_path, engine="openpyxl") as xw:
        changed_df.to_excel(xw, sheet_name="Changed", index=False)

    return changed_df


curr_fp = r"C:\Users\tashif.ahmed\OneDrive - Shell\Documents\Projects\ISCC certificates scraping\out\ISCC_Certificates_27.01.2026_16.04.xlsx"
prev_fp = r"C:\Users\tashif.ahmed\OneDrive - Shell\Documents\Projects\ISCC certificates scraping\out\ISCC_Certificates_20.01.2026_13.16.xlsx"

get_certs_changed(Path(prev_fp), Path(curr_fp), ignore_cols=["Map"])


