import argparse
import io
import sys
import pandas as pd
import pdfplumber
from pathlib import Path
try:
    from tqdm import tqdm
except ImportError:
    tqdm = lambda x, **kwargs: x

# Import backend logic from existing files
from main import extract_record_from_page, extract_ledger_records
from qc_core import rename_excel_columns, check_single

def load_excel(excel_path):
    print(f"Loading Excel file: {excel_path}...")
    xl = pd.ExcelFile(excel_path)
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
        print("Error: No sheet found with 'NewWardNo' and 'NewPropertyNo' columns.")
        sys.exit(1)
        
    print(f"Detected valid sheet: {chosen_sheet} ({len(df)} rows)")
    df = rename_excel_columns(df)
    return df

def parse_pdf(pdf_path, target_wards=None):
    print(f"Parsing PDF file: {pdf_path}...")
    records = []
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            return []
            
        # Detect ledger vs normal format
        first_page_tables = pdf.pages[0].extract_tables() or []
        is_ledger = False
        for tbl in first_page_tables:
            if tbl and len(tbl) > 0 and tbl[0]:
                headers = " ".join(str(c) for c in tbl[0] if c).replace("\n", "")
                if "नपवनवमरर" in headers or "नपवन वमरर" in headers:
                    is_ledger = True
                    break
        
        print(f"Format detected: {'Ledger' if is_ledger else 'Standard Report'}")
        
        if is_ledger:
            for page in tqdm(pdf.pages, desc="Extracting pages"):
                tbls = page.extract_tables() or []
                page_records = extract_ledger_records(tbls)
                for rec in page_records:
                    if not target_wards or str(rec.get("NewWardNo")).strip().upper() in target_wards:
                        records.append(rec)
        else:
            for page in tqdm(pdf.pages, desc="Extracting pages"):
                rec = extract_record_from_page(page)
                if rec.get("NewWardNo") or rec.get("NewPropertyNo"):
                    if not target_wards or str(rec.get("NewWardNo")).strip().upper() in target_wards:
                        records.append(rec)
                        
    return records

def run_reconciliation(records, df):
    print("Running reconciliation against Excel master...")
    all_results = []
    for rec in tqdm(records, desc="Reconciling records"):
        r = check_single(rec, df)
        r["extracted_from_report"] = rec
        all_results.append(r)
    return all_results

def export_report(all_results, output_file):
    print(f"Generating Excel report: {output_file}...")
    
    # Process results into flat lists for DataFrames
    discrepancies = []
    missing_in_excel = []
    
    total_records = len(all_results)
    total_matched = 0
    total_mismatched = 0
    total_missing = 0
    
    for r in all_results:
        rec = r["extracted_from_report"]
        status = r.get("status")
        
        ward = rec.get("NewWardNo")
        prop = rec.get("NewPropertyNo")
        part = rec.get("NewPartitionNo")
        
        if status == "MISSING_IN_EXCEL":
            total_missing += 1
            row = {
                "Owner ID": "NOT_FOUND",
                "Ward": ward,
                "Property No": prop,
                "Partition No": part,
                "Overall Result": status
            }
            # Fill missing with report values
            for k, v in rec.items():
                if k not in ["NewWardNo", "NewPropertyNo", "NewPartitionNo"]:
                    row[f"{k} (PDF)"] = v
                    row[f"{k} (Excel)"] = "N/A"
                    row[f"{k} (Result)"] = "MISSING_IN_EXCEL"
            missing_in_excel.append(row)
            continue
            
        if status == "MATCH":
            total_matched += 1
        else:
            total_mismatched += 1
            
        owner_id = r.get("key", {}).get("OwnerID", "")
        row = {
            "Owner ID": owner_id,
            "Ward": ward,
            "Property No": prop,
            "Partition No": part,
            "Overall Result": status
        }
        
        # Populate all fields side-by-side
        if r.get("results"):
            for res in r["results"]:
                field = res.get("field")
                if field not in ["NewWardNo", "NewPropertyNo", "NewPartitionNo", "OwnerID"]:
                    row[f"{field} (PDF)"] = res.get("report_value")
                    row[f"{field} (Excel)"] = res.get("excel_value")
                    row[f"{field} (Result)"] = res.get("status")
                    
        discrepancies.append(row)

    # Summary df
    summary_df = pd.DataFrame([{
        "Metric": "Total Records Parsed from PDF",
        "Value": total_records
    }, {
        "Metric": "Total Perfect Matches",
        "Value": total_matched
    }, {
        "Metric": "Total Mismatched Properties",
        "Value": total_mismatched
    }, {
        "Metric": "Total Properties Missing in Excel",
        "Value": total_missing
    }])
    
    # Write to excel using pandas ExcelWriter
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        if discrepancies:
            pd.DataFrame(discrepancies).to_excel(writer, sheet_name='All_Properties', index=False)
        else:
            pd.DataFrame([{"Message": "No records found!"}]).to_excel(writer, sheet_name='All_Properties', index=False)
            
        if missing_in_excel:
            pd.DataFrame(missing_in_excel).to_excel(writer, sheet_name='Missing_In_Excel', index=False)
        else:
            pd.DataFrame([{"Message": "All properties found in Excel!"}]).to_excel(writer, sheet_name='Missing_In_Excel', index=False)
            
    print(f"Success! Report generated at {output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Bulk reconcile PDF report with Excel Master")
    parser.add_argument("--pdf", default="a1-798.pdf", help="Path to PDF file")
    parser.add_argument("--excel", default="amravti demand bill.xlsx", help="Path to Excel file")
    parser.add_argument("--wards", nargs="*", default=["A1", "A2", "A3", "A4", "A5", "A6"], help="Wards to filter (e.g. A1 A2 A3)")
    parser.add_argument("--output", default="reconciliation_report_A1_A6.xlsx", help="Output Excel file path")
    
    args = parser.parse_args()
    
    # Normalize target wards
    target_wards = [w.strip().upper() for w in args.wards]
    print(f"Target wards for filtering: {target_wards}")
    
    df = load_excel(args.excel)
    records = parse_pdf(args.pdf, target_wards=target_wards)
    print(f"Total records extracted for target wards: {len(records)}")
    
    if not records:
        print("No records found in PDF for the specified wards.")
        sys.exit(0)
        
    results = run_reconciliation(records, df)
    export_report(results, args.output)
