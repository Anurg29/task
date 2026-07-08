"""
main.py — FastAPI backend for Report vs Excel QC system

Logic:
  1. Upload Excel  → find the sheet that has NewWardNo + NewPropertyNo columns
  2. Upload PDF    → extract NewWardNo, NewPropertyNo and all tax values from table
  3. Lookup Excel row by (NewWardNo + NewPropertyNo) composite key
  4. Compare every field — report MATCH, MISMATCH, or MISSING
"""

import io
import re
import pandas as pd

import tempfile
import os
import math
import asyncio
from concurrent.futures import ProcessPoolExecutor
from fastapi import WebSocket, WebSocketDisconnect

import pdfplumber
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware

from qc_core import check_single, COLUMN_TYPES, KEY_COLUMNS, rename_excel_columns

app = FastAPI(title="Report vs Excel QC API")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

STATE = {"excel_df": None, "sheet_name": None}

PROGRESS_STORE = {}

@app.websocket("/ws/progress/{client_id}")
async def websocket_endpoint(websocket: WebSocket, client_id: str):
    await websocket.accept()
    PROGRESS_STORE[client_id] = {"scanned": 0, "total": 1}
    try:
        while True:
            await asyncio.sleep(0.5)
            data = PROGRESS_STORE.get(client_id)
            if data:
                await websocket.send_json(data)
                if data.get("scanned", 0) >= data.get("total", 1):
                    break
    except WebSocketDisconnect:
        pass
    finally:
        PROGRESS_STORE.pop(client_id, None)

def worker_parse_pages(pdf_path: str, start_page: int, end_page: int, target_wards: list, is_ledger: bool) -> list:
    records = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i in range(start_page, end_page):
                page = pdf.pages[i]
                if is_ledger:
                    tbls = page.extract_tables() or []
                    for r in extract_ledger_records(tbls):
                        if target_wards:
                            w = str(r.get("NewWardNo", "")).strip().upper()
                            if w not in target_wards:
                                continue
                        records.append(r)
                else:
                    rec = extract_record_from_page(page)
                    if rec.get("NewWardNo") or rec.get("NewPropertyNo"):
                        if target_wards:
                            w = str(rec.get("NewWardNo", "")).strip().upper()
                            if w not in target_wards:
                                continue
                        records.append(rec)
    except Exception:
        pass
    return records



# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
def clean_cell(val) -> str | None:
    """Strip None, CID garbage, whitespace from a PDF table cell."""
    if val is None:
        return None
    v = re.sub(r"\(cid:\d+\)", "", str(val)).strip()
    # Also strip leading/trailing Marathi half-chars that render as junk
    v = v.strip()
    return v if v else None


def safe_get(tbl, row, col) -> str | None:
    try:
        return clean_cell(tbl[row][col])
    except (IndexError, TypeError):
        return None


def to_num(v) -> float | None:
    if v is None:
        return None
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Excel upload — auto-detect correct sheet
# ─────────────────────────────────────────────────────────────────────────────
@app.post("/excel/upload")
async def upload_excel(file: UploadFile = File(...)):
    content = await file.read()
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:
        raise HTTPException(400, f"Could not open Excel file: {e}")

    df = None
    chosen_sheet = None
    for sheet in xl.sheet_names:
        try:
            candidate = xl.parse(sheet)
            if "NewWardNo" in candidate.columns and "NewPropertyNo" in candidate.columns:
                df = candidate
                chosen_sheet = sheet
                break
        except Exception:
            continue

    if df is None:
        raise HTTPException(
            400,
            f"No sheet found with 'NewWardNo' and 'NewPropertyNo' columns. "
            f"Sheets available: {xl.sheet_names}"
        )

    # Rename positional tax columns to internal English keys
    df = rename_excel_columns(df)

    STATE["excel_df"] = df
    STATE["sheet_name"] = chosen_sheet
    return {
        "status": "ok",
        "sheet": chosen_sheet,
        "rows_loaded": len(df),
        "columns": list(df.columns),
    }


def get_excel_df() -> pd.DataFrame:
    if STATE["excel_df"] is None:
        raise HTTPException(400, "No Excel loaded. Upload master Excel first.")
    return STATE["excel_df"]


# ─────────────────────────────────────────────────────────────────────────────
# PDF extraction — Amravati Demand Bill structure
#
# Table layout (Table index 2 on page, 0-indexed):
#   row[0]: [zone_lbl, zone_val, ..., bill_lbl, ..., BILL_NO, ...]
#   row[1]: [ward_lbl,  D10,    ..., old_prop_lbl, ..., 75-1052, ...]
#   row[2]: [prop_lbl,  2465,   ..., old_comp_lbl, ..., 0475001161, ...]
#   row[3]: [UPIC_lbl,  UPIC,   ..., mobile_lbl, ..., MOBILE, ...]
#   row[4]: [owner_lbl, OWNER_NAME, ...]
#   row[5]: [occupier_lbl, OCCUPIER, ...]
#   row[6]: [rv_lbl, RV, ..., prop_desc_lbl, ..., PROP_DESC, ...]
#   row[9]: Tax header row
#   row[10..23]: Tax data rows  — col[1]=arrear, col[2]=current, col[8]=TOTAL
#   row[24]: Grand total row   — col[8]=TM_TaxTotal
#
# TAX column order in the bill (rows 10-23, col[8] = total):
#   row[10]: General tax (सामान्य कर)  → tax_ekatrit
#   row[11]: Fire tax (अग्नि)           → tax_agni
#   row[12]: Tree tax (वृक्ष)            → tax_vriksha
#   row[13]: Street tax (स्ट्रीट)        → tax_path
#   row[14]: —  (blank row, skip)
#   row[15]: Water (पाणीपट्टी)           → tax_panipatti
#   row[16]: —  (blank/note row, skip)
#   row[17]: —  (blank row, skip)
#   row[18]: Education cess (शिक्षण)     → tax_education
#   row[19]: Employment (रोजगार हमी)     → tax_rojgar
#   row[20]: Large building (मोठी इमारत)→ tax_moti_imaarat
#   row[21]: User charge (उपयोगकर्ता)    → tax_upayogkarta
#   row[22]: Penalty (शास्ती)            → tax_thakeet
# ─────────────────────────────────────────────────────────────────────────────

# Map: table row index → internal tax key
# The TOTAL value in each tax row is the LAST non-None cell
TAX_ROW_MAP = {
    10: "tax_ekatrit",      # row 10 = general + fire + tree + street + water + education + sewage + sanitation (combined label)
    11: "tax_agni",
    12: "tax_vriksha",
    13: "tax_path",
    15: "tax_panipatti",
    18: "tax_education",
    19: "tax_rojgar",
    20: "tax_moti_imaarat",
    21: "tax_upayogkarta",
    22: "tax_thakeet",
}


def get_last_num(row: list) -> float | None:
    """Return the last numeric value in a table row (= the TOTAL column)."""
    vals = []
    for cell in row:
        if cell is None:
            continue
        v = clean_cell(cell)
        if v:
            n = to_num(v)
            if n is not None:
                vals.append(n)
    return vals[-1] if vals else None

def find_header_table(tables: list) -> list | None:
    """
    Find the main data table in the PDF page.
    According to known layout, it's typically table index 2 and has ~25 rows.
    We return the first table with more than 20 rows, or fallback to index 2.
    """
    if not tables:
        return None
    
    # The main table with tax data and headers has 25+ rows.
    for tbl in tables:
        if len(tbl) > 20:
            return tbl
            
    # Fallback to the index from comments if no table is long enough
    if len(tables) > 2:
        return tables[2]
        
    return tables[-1]


def extract_record_from_page(page) -> dict:
    """
    Extract a single property record dict from one pdfplumber page.
    Returns keys: NewWardNo, NewPropertyNo, OldPropertyNo, MobileNo,
                  MarathiOwnerName, RateableValue, TM_TaxTotal, + tax_* keys
    """
    record = {}
    tables = page.extract_tables() or []
    tbl = find_header_table(tables)

    if tbl:
        # ── Identity fields ──────────────────────────────────────────────
        record["NewWardNo"] = safe_get(tbl, 1, 1)   # e.g. "D10"

        # Property number may include partition: "2465" or "2465-1"
        prop_raw = safe_get(tbl, 2, 1)
        if prop_raw and "-" in str(prop_raw):
            parts = str(prop_raw).split("-", 1)
            record["NewPropertyNo"]  = parts[0].strip()
            record["NewPartitionNo"] = parts[1].strip()
        else:
            record["NewPropertyNo"]  = prop_raw
            record["NewPartitionNo"] = None   # no partition → match NaN rows in Excel

        # Also try col 3 if col 7 is None (compact table layout)
        record["OldPropertyNo"] = safe_get(tbl, 1, 7) or safe_get(tbl, 1, 3)
        record["MobileNo"]      = safe_get(tbl, 3, 7) or safe_get(tbl, 3, 3)
        record["MarathiOwnerName"] = safe_get(tbl, 4, 1)
        rv_raw = safe_get(tbl, 6, 1)
        record["RateableValue"] = to_num(rv_raw)

        # ── Tax values: LAST numeric value in each tax row = TOTAL ────────
        for row_idx, tax_key in TAX_ROW_MAP.items():
            if row_idx < len(tbl):
                record[tax_key] = get_last_num(tbl[row_idx])

        # ── Grand total: last number in row[24] ───────────────────────────
        if len(tbl) > 24:
            grand = get_last_num(tbl[24])
            # Sanity check: if it looks like a merged-cell artifact (e.g. 24249417)
            # fall back to summing tax values
            if grand is not None and grand < 1_000_000:
                record["TM_TaxTotal"] = grand
            else:
                record["TM_TaxTotal"] = None

        # Fallback: sum individual tax values
        if record.get("TM_TaxTotal") is None:
            tax_sum = sum(v for v in [record.get(k) for k in TAX_ROW_MAP.values()] if v is not None)
            if tax_sum > 0:
                record["TM_TaxTotal"] = round(tax_sum, 2)

    return record


def extract_ledger_records(tables) -> list[dict]:
    records = []
    
    LEDGER_TAX_MAP = {
        8: "tax_ekatrit",
        9: "tax_agni",
        10: "tax_vriksha",
        11: "tax_path",
        12: "tax_panipatti",
        13: "tax_education",
        17: "tax_rojgar",
        18: "tax_moti_imaarat",
        19: "tax_upayogkarta",
        20: "tax_thakeet",
    }
    
    for tbl in tables:
        current_record = None
        for row in tbl:
            if not row: continue
            
            # Start of a new property block (Serial Number is present)
            if str(row[0]).strip().isdigit():
                current_record = {}
                current_record["NewWardNo"] = safe_get([row], 0, 1)
                prop_raw = safe_get([row], 0, 2)
                if prop_raw and "-" in str(prop_raw):
                    parts = str(prop_raw).split("-", 1)
                    current_record["NewPropertyNo"] = parts[0].strip()
                    current_record["NewPartitionNo"] = parts[1].strip()
                else:
                    current_record["NewPropertyNo"] = prop_raw
                    current_record["NewPartitionNo"] = None
                
                current_record["OldPropertyNo"] = safe_get([row], 0, 3)
                current_record["MarathiOwnerName"] = safe_get([row], 0, 4)
            
            # If we find the "Total Demand" row for the current property
            if current_record and safe_get([row], 0, 7) == "एकमण ममगणठ र.":
                for col_idx, tax_key in LEDGER_TAX_MAP.items():
                    current_record[tax_key] = to_num(safe_get([row], 0, col_idx))
                
                current_record["TM_TaxTotal"] = to_num(safe_get([row], 0, 21))
                
                records.append(current_record)
                current_record = None # Reset for next
                
    return records


async def extract_records_from_pdf(pdf_bytes: bytes, target_wards: list[str] | None = None, client_id: str | None = None) -> list[dict]:
    """Extract records from a multi-page PDF using multiprocessing."""
    records = []
    tmp_path = None
    try:
        # Write bytes to temp file for worker processes
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf_bytes)
            tmp_path = tmp.name

        with pdfplumber.open(tmp_path) as pdf:
            total_pages = len(pdf.pages)
            if total_pages == 0:
                os.remove(tmp_path)
                return []
                
            first_page_tables = pdf.pages[0].extract_tables() or []
            is_ledger = False
            for tbl in first_page_tables:
                if tbl and len(tbl) > 0 and tbl[0]:
                    headers = " ".join(str(c) for c in tbl[0] if c).replace("\n", "")
                    if "नपवनवमरर" in headers or "नपवन वमरर" in headers:
                        is_ledger = True
                        break

        CHUNK_SIZE = 25  # Process 25 pages per worker
        chunks = []
        for i in range(0, total_pages, CHUNK_SIZE):
            chunks.append((tmp_path, i, min(i + CHUNK_SIZE, total_pages), target_wards, is_ledger))

        loop = asyncio.get_running_loop()
        completed_pages = 0
        if client_id:
            PROGRESS_STORE[client_id] = {"scanned": 0, "total": total_pages}

        with ProcessPoolExecutor() as executor:
            async def run_chunk(exec_ref, ch):
                r = await loop.run_in_executor(exec_ref, worker_parse_pages, *ch)
                return r, ch[2] - ch[1]

            tasks = [run_chunk(executor, chunk) for chunk in chunks]
            
            for f in asyncio.as_completed(tasks):
                res, num_pages_in_chunk = await f
                records.extend(res)
                completed_pages += num_pages_in_chunk
                if client_id:
                    PROGRESS_STORE[client_id] = {"scanned": min(completed_pages, total_pages), "total": total_pages}

    except Exception as e:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise HTTPException(400, f"Invalid PDF file: {e}")
        
    if tmp_path and os.path.exists(tmp_path):
        os.remove(tmp_path)
        
    return records


async def extract_record_from_pdf(pdf_bytes: bytes,
                             override_ward: str | None = None,
                             override_prop: str | None = None,
                             override_partition: str | None = None) -> dict:
    try:
        records = await extract_records_from_pdf(pdf_bytes)
        if not records:
             raise ValueError("No valid property records found in PDF")
        
        record = records[0]
        
        # If overrides provided, try to find the matching record in the ledger
        if override_ward and override_prop:
            target_ward = override_ward.strip().upper()
            target_prop = override_prop.strip()
            for r in records:
                w = str(r.get("NewWardNo", "")).strip().upper()
                p = str(r.get("NewPropertyNo", "")).strip()
                if w == target_ward and p == target_prop:
                    record = r
                    break
                    
    except Exception as e:
        raise HTTPException(400, f"Invalid PDF file or unsupported format: {e}")
        
    if override_ward:
        record["NewWardNo"] = override_ward.strip().upper()
    if override_prop:
        record["NewPropertyNo"] = override_prop.strip()
    if override_partition is not None:
        val = override_partition.strip()
        record["NewPartitionNo"] = val if val else None
    return record


# ─────────────────────────────────────────────────────────────────────────────
# QC endpoints
# ─────────────────────────────────────────────────────────────────────────────
from pydantic import BaseModel
import json
import os

class TypoApproval(BaseModel):
    pdf_text: str
    excel_text: str

@app.post("/qc/approve-typo")
async def approve_typo(typo: TypoApproval):
    typo_file = os.path.join(os.path.dirname(__file__), "typos.json")
    typos = {}
    if os.path.exists(typo_file):
        try:
            with open(typo_file, "r") as f:
                typos = json.load(f)
        except Exception:
            pass
    typos[typo.pdf_text] = typo.excel_text
    with open(typo_file, "w") as f:
        json.dump(typos, f, ensure_ascii=False, indent=2)
    return {"status": "ok", "message": "Typo approved and saved."}
@app.post("/qc/check-single")
async def qc_check_single(
    file: UploadFile = File(...),
    ward_no: str | None = Form(None),
    prop_no: str | None = Form(None),
    partition_no: str | None = Form(None),
):
    pdf_bytes = await file.read()
    report_record = await extract_record_from_pdf(
        pdf_bytes,
        override_ward=ward_no,
        override_prop=prop_no,
        override_partition=partition_no,
    )

    missing = []
    if not report_record.get("NewWardNo"):
        missing.append("Ward No (e.g. D10)")
    if not report_record.get("NewPropertyNo"):
        missing.append("Property No (e.g. 2465)")

    if missing:
        raise HTTPException(
            422,
            f"Could not auto-read {', '.join(missing)} from the PDF. "
            "Please type them in the override fields."
        )

    excel_df = get_excel_df()
    result = check_single(report_record, excel_df)
    result["extracted_from_report"] = report_record
    return result


def expand_ward_ranges(ward_str: str) -> list[str]:
    if not ward_str:
        return []
    
    parts = [p.strip() for p in ward_str.split(",") if p.strip()]
    final_wards = []
    
    for part in parts:
        if "-" in part:
            subparts = [s.strip() for s in part.split("-")]
            if len(subparts) == 2:
                start_w, end_w = subparts[0].upper(), subparts[1].upper()
                m1 = re.match(r"^([A-Z]+)(\d+)$", start_w)
                m2 = re.match(r"^([A-Z]+)(\d+)$", end_w)
                if m1 and m2 and m1.group(1) == m2.group(1):
                    prefix = m1.group(1)
                    start_num = int(m1.group(2))
                    end_num = int(m2.group(2))
                    
                    step = 1 if start_num <= end_num else -1
                    for n in range(start_num, end_num + step, step):
                        final_wards.append(f"{prefix}{n}")
                    continue
        
        final_wards.append(part.upper())
    return final_wards

@app.post("/qc/check-bulk")
async def qc_check_bulk(
    file: UploadFile = File(...),
    target_wards: str | None = Form(None),
    client_id: str | None = Form(None)
):
    pdf_bytes = await file.read()
    
    wards_list = expand_ward_ranges(target_wards) if target_wards else None

    records = await extract_records_from_pdf(pdf_bytes, target_wards=wards_list, client_id=client_id)

    if not records:
        raise HTTPException(422, "Could not extract any property records from the PDF (or none matched the ward filter).")

    excel_df = get_excel_df()
    all_results = []
    for rec in records:
        r = check_single(rec, excel_df)
        r["extracted_from_report"] = rec
        all_results.append(r)

    return {
        "total_records": len(all_results),
        "total_mismatches": sum(r.get("total_mismatches", 0) for r in all_results),
        "records": all_results,
    }


@app.get("/health")
async def health():
    df = STATE["excel_df"]
    return {
        "status": "ok",
        "excel_loaded": df is not None,
        "sheet": STATE.get("sheet_name"),
        "rows": len(df) if df is not None else 0,
    }
