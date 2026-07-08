"""
qc_core.py
----------
Reconciliation engine for Amravati / Chalisgaon property tax demand bills.

KEY LOGIC:
  - Look up Excel row by composite key: (NewPropertyNo + NewWardNo)
  - Compare every one of the 23 PDF fields against the matched Excel row
  - Single mode: one PDF (one property) vs Excel
  - Bulk mode  : one PDF (many properties) vs Excel
"""

import unicodedata
import re
import pandas as pd
from rapidfuzz import fuzz

# ---------------------------------------------------------------------------
# Devanagari digit normalisation
# ---------------------------------------------------------------------------
DEVANAGARI_DIGITS = str.maketrans("०१२३४५६७८९", "0123456789")

NUMERIC_TOLERANCE = 0.5   # allow ±0.50 rounding difference
FUZZY_TEXT_THRESHOLD = 85  # % similarity for Marathi name fields

# ---------------------------------------------------------------------------
# KEY COLUMNS  —  composite lookup key for finding the Excel row
# NewWardNo + NewPropertyNo + NewPartitionNo together uniquely identify a property.
# (Same property number can exist in different wards)
# ---------------------------------------------------------------------------
KEY_COLUMNS = ["NewWardNo", "NewPropertyNo", "NewPartitionNo"]

# ---------------------------------------------------------------------------
# COLUMN TYPE MAP
# Every column the PDF can contain, mapped to its comparison type.
# Marathi column names from Excel are aliased to English keys here.
# ---------------------------------------------------------------------------
COLUMN_TYPES = {
    # --- identity / metadata (string exact match after normalisation) ---
    "NewWardNo":           "string",
    "NewPropertyNo":       "string",
    "NewPartitionNo":      "string",
    "OldWardNo":           "string",
    "OldPropertyNo":       "string",
    "OwnerID":             "string",
    "AssesmentID":         "string",
    "FinanceYear":         "string",
    "RateableValue":       "number",
    "MobileNo":            "string",
    # --- owner names (fuzzy Marathi text) ---
    "MarathiOwnerName":    "text",
    # --- 23 tax columns (numeric) ---
    "tax_ekatrit":         "number",   # एकत्रित मालमत्ता कर
    "tax_education":       "number",   # महा. शिक्षण कर
    "tax_rojgar":          "number",   # रोजगार हमी कर
    "tax_vriksha":         "number",   # वृक्ष संवर्धन कर
    "tax_edu_penalty":     "number",   # महाराष्ट्र शिक्षण कर दंड
    "tax_swachhata":       "number",   # विशेष स्वच्छता कर
    "tax_upayogkarta":     "number",   # उपयोगकर्ता शुल्क
    "tax_panipatti":       "number",   # पाणीपट्टी कर
    "tax_path":            "number",   # पथ कर
    "tax_agni":            "number",   # अग्निशमन कर
    "tax_divabatti":       "number",   # दिवाबत्ती कर
    "tax_jallabha":        "number",   # जल लाभ कर
    "tax_thakeet":         "number",   # थकीत शास्ती
    "tax_malnissaran":     "number",   # मलनिस्सारण कर
    "tax_pani":            "number",   # पाणी कर
    "tax_anadhikrut":      "number",   # अनधिकृत बांध काम शास्ती
    "tax_moti_imaarat":    "number",   # मोठी इमारत कर
    "tax_kar3":            "number",   # कर ३
    "tax_kar4":            "number",   # कर ४
    "tax_kar5":            "number",   # कर ५
    "TM_TaxTotal":         "number",   # एकूण कर (grand total)
}

# ---------------------------------------------------------------------------
# Excel column name → internal key  (position-based fallback via index)
# ---------------------------------------------------------------------------
EXCEL_COL_ALIASES: dict[str, str] = {
    # identity columns
    "NewWardNo":           "NewWardNo",
    "NewPropertyNo":       "NewPropertyNo",
    "NewPartitionNo":      "NewPartitionNo",
    "OldWardNo":           "OldWardNo",
    "OldPropertyNo":       "OldPropertyNo",
    "OwnerID":             "OwnerID",
    "AssesmentID":         "AssesmentID",
    "FinanceYear":         "FinanceYear",
    "RateableValue":       "RateableValue",
    "MobileNo":            "MobileNo",
    "MarathiOwnerName":    "MarathiOwnerName",
    "TM_TaxTotal":         "TM_TaxTotal",
}

# Known English column names — these are NOT tax columns
_KNOWN_ENGLISH_COLS = {
    "OwnerID", "NewZoneNo", "NewWardNo", "NewPropertyNo", "NewPartitionNo",
    "NewCityServeyNo", "NewPlotNo", "OpenPlot", "PlotArea", "PropertyDescription",
    "MarathiOwnerName", "MarathiOwnerPatta", "MarathiOwnerDukanImarateNav",
    "MarathiOwnerDukanFlatNo", "CombPropRemark", "Remark", "Category",
    "MarathiSocietyName", "OccupierNameMarathi",
    "OldZoneNo", "OldWardNo", "OldPropertyNo", "OldPartitionNo",
    "AssesmentID", "FinanceYear", "RateableValue", "MobileNo",
    "TM_TaxTotal", "PendingYear",
}

# Ordered tax key names — assigned in order to the Marathi tax columns found
TAX_KEYS_IN_ORDER = [
    "tax_ekatrit", "tax_education", "tax_rojgar", "tax_vriksha",
    "tax_edu_penalty", "tax_swachhata", "tax_upayogkarta", "tax_panipatti",
    "tax_path", "tax_agni", "tax_divabatti", "tax_jallabha",
    "tax_thakeet", "tax_malnissaran", "tax_pani", "tax_anadhikrut",
    "tax_moti_imaarat", "tax_kar3", "tax_kar4", "tax_kar5",
]


def _is_marathi_col(col_name: str) -> bool:
    """Return True if the column name contains Marathi/Devanagari characters."""
    return any("\u0900" <= ch <= "\u097F" for ch in str(col_name))


def rename_excel_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Dynamically rename Marathi tax columns to internal English keys.
    Finds all columns with Devanagari characters and assigns tax key names in order.
    """
    col_map = {}
    marathi_cols = [c for c in df.columns if _is_marathi_col(c)]
    for i, mc in enumerate(marathi_cols):
        if i < len(TAX_KEYS_IN_ORDER):
            col_map[mc] = TAX_KEYS_IN_ORDER[i]
    return df.rename(columns=col_map)


# ---------------------------------------------------------------------------
# Normalisation helpers
# ---------------------------------------------------------------------------
def normalize_value(value, col_type: str):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    value = str(value).strip()
    value = unicodedata.normalize("NFC", value)
    value = value.translate(DEVANAGARI_DIGITS)

    if value in ("", "nan", "None", "-", "NULL", "NaN"):
        return None

    if col_type == "number":
        cleaned = re.sub(r"[₹$,\s]", "", value)
        cleaned = cleaned.replace("(", "-").replace(")", "")
        try:
            return round(float(cleaned), 2)
        except ValueError:
            return value

    if col_type == "date":
        try:
            return pd.to_datetime(value, format="%Y-%m-%d", errors="raise").date()
        except (ValueError, TypeError):
            pass
        try:
            return pd.to_datetime(value, dayfirst=True, errors="raise").date()
        except Exception:
            return value

    if col_type == "string":
        v = value.upper().strip()
        # Remove trailing .0 from float-integers that Excel stores (e.g. "2465.0" → "2465")
        if v.endswith(".0") and v[:-2].lstrip("-").isdigit():
            v = v[:-2]
        return v

    return value  # "text" — fuzzy match handles minor diffs


def normalize_record(record: dict, col_types: dict = COLUMN_TYPES) -> dict:
    return {col: normalize_value(record.get(col), t) for col, t in col_types.items()}


def normalize_dataframe(df: pd.DataFrame, col_types: dict = COLUMN_TYPES) -> pd.DataFrame:
    df_norm = df.copy()
    for col, col_type in col_types.items():
        if col in df_norm.columns:
            df_norm[col] = df_norm[col].apply(lambda v: normalize_value(v, col_type))
    return df_norm


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------
def values_match(report_val, excel_val, col_type: str):
    """Returns (matched: bool, issue_type: str)."""
    if report_val is None and excel_val is None:
        return True, "OK"
    if report_val is None:
        return False, "MISSING_IN_REPORT"
    if excel_val is None:
        return False, "MISSING_IN_EXCEL"

    if col_type == "number":
        try:
            return abs(float(report_val) - float(excel_val)) <= NUMERIC_TOLERANCE, "VALUE_MISMATCH"
        except (TypeError, ValueError):
            return str(report_val) == str(excel_val), "VALUE_MISMATCH"

    if col_type == "text":
        score = fuzz.ratio(str(report_val), str(excel_val))
        return score >= FUZZY_TEXT_THRESHOLD, "VALUE_MISMATCH"

    return report_val == excel_val, "VALUE_MISMATCH"


# ---------------------------------------------------------------------------
# Excel row lookup by composite key (NewWardNo + NewPropertyNo + NewPartitionNo)
# ---------------------------------------------------------------------------
def find_excel_row(report_record: dict, excel_df: pd.DataFrame) -> pd.DataFrame:
    """
    Find the Excel row(s) matching:
      - NewWardNo      (exact string match)
      - NewPropertyNo  (exact string match, handles float-int from Excel)
      - NewPartitionNo (None/blank in PDF → match NaN rows in Excel;
                        a number in PDF   → match that numeric partition)
    """
    mask = pd.Series([True] * len(excel_df), index=excel_df.index)

    for key in KEY_COLUMNS:
        if key not in excel_df.columns:
            continue

        report_val_raw = report_record.get(key)

        if key == "NewPartitionNo":
            # If PDF has no partition (or it's blank), skip filtering on it
            # so we don't accidentally fail if the Excel has a partition but PDF doesn't.
            norm = normalize_value(report_val_raw, "string")
            if norm is None:
                continue
            else:
                # Match numeric partition
                excel_col_norm = excel_df[key].apply(lambda v: normalize_value(v, "string"))
                mask &= excel_col_norm == norm
        else:
            norm = normalize_value(report_val_raw, "string")
            if norm is None:
                continue  # skip this key if PDF couldn't extract it
            excel_col_norm = excel_df[key].apply(lambda v: normalize_value(v, "string"))
            mask &= excel_col_norm == norm

    return excel_df[mask]


# ---------------------------------------------------------------------------
# MODE 1: Single-property check
# ---------------------------------------------------------------------------
def check_single(report_record: dict, excel_df: pd.DataFrame,
                 col_types: dict = COLUMN_TYPES) -> dict:
    """
    Compare a single-property report record against the Excel master.
    Lookup key: NewWardNo + NewPropertyNo + NewPartitionNo.
    All other columns are compared field-by-field.
    """
    key_values = {k: report_record.get(k) for k in KEY_COLUMNS}

    matches = find_excel_row(report_record, excel_df)

    if matches.empty:
        return {
            "found_in_excel": False,
            "key": key_values,
            "message": (
                f"No Excel row found for "
                f"Ward={key_values.get('NewWardNo')!r}, "
                f"Property={key_values.get('NewPropertyNo')!r}, "
                f"Partition={key_values.get('NewPartitionNo')!r}."
            ),
            "discrepancies": [],
            "results": [],
        }

    # Convert pandas NaNs to None to avoid JSON serialization errors
    excel_row = {k: (v if pd.notna(v) else None) for k, v in matches.iloc[0].items()}
    r_norm = normalize_record(report_record, col_types)
    e_norm = normalize_record(excel_row, col_types)

    discrepancies = []
    results = []

    for col in col_types:
        if col in KEY_COLUMNS:
            continue
        matched, issue_type = values_match(r_norm.get(col), e_norm.get(col), col_types[col])
        entry = {
            "field": col,
            "report_value": report_record.get(col),
            "excel_value":  excel_row.get(col),
            "status": "MATCH" if matched else issue_type,
        }
        results.append(entry)
        if not matched:
            discrepancies.append(entry)

    return {
        "found_in_excel": True,
        "key": key_values,
        "excel_identifiers": {
            "OwnerID": excel_row.get("OwnerID"),
            "OldPropertyNo": excel_row.get("OldPropertyNo"),
            "AssesmentID": excel_row.get("AssesmentID")
        },
        "discrepancies": discrepancies,
        "results": results,
        "status": "MATCH" if not discrepancies else "MISMATCH",
        "total_fields_checked": len(results),
        "total_mismatches": len(discrepancies),
    }


# ---------------------------------------------------------------------------
# MODE 2: Bulk check (multi-property report)
# ---------------------------------------------------------------------------
def check_bulk(report_df: pd.DataFrame, excel_df: pd.DataFrame,
               col_types: dict = COLUMN_TYPES) -> dict:
    """
    Compare every row in report_df against the Excel master using
    (NewWardNo + NewPropertyNo) as the composite lookup key.
    """
    all_results = []
    total_mismatches = 0
    compare_cols = [c for c in col_types if c not in KEY_COLUMNS]

    for _, report_row in report_df.iterrows():
        record = report_row.to_dict()
        key_values = {k: record.get(k) for k in KEY_COLUMNS}
        key_str = f"Ward={key_values.get('NewWardNo')} / Prop={key_values.get('NewPropertyNo')}"

        matches = find_excel_row(record, excel_df)

        if matches.empty:
            all_results.append({
                "key": key_str,
                "found_in_excel": False,
                "message": f"No Excel row found for {key_str}",
                "discrepancies": [],
                "status": "NOT_FOUND",
            })
            total_mismatches += 1
            continue

        # Convert pandas NaNs to None to avoid JSON serialization errors
        excel_row = {k: (v if pd.notna(v) else None) for k, v in matches.iloc[0].items()}
        r_norm = normalize_record(record, col_types)
        e_norm = normalize_record(excel_row, col_types)

        discrepancies = []
        for col in compare_cols:
            matched, issue_type = values_match(r_norm.get(col), e_norm.get(col), col_types[col])
            if not matched:
                discrepancies.append({
                    "field": col,
                    "report_value": record.get(col),
                    "excel_value": excel_row.get(col),
                    "issue_type": issue_type,
                })

        total_mismatches += len(discrepancies)
        all_results.append({
            "key": key_str,
            "found_in_excel": True,
            "discrepancies": discrepancies,
            "status": "MATCH" if not discrepancies else "MISMATCH",
        })

    return {
        "total_records_compared": len(report_df),
        "total_discrepancies": total_mismatches,
        "records": all_results,
    }
