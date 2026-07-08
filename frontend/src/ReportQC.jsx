import { useState } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import "jspdf-autotable";

const API_BASE = "http://localhost:8000";

// ─── Status badge colours ───────────────────────────────────────────────────
const STATUS_COLOR = {
  MATCH:              { bg: "#d4edda", text: "#155724" },
  VALUE_MISMATCH:     { bg: "#f8d7da", text: "#721c24" },
  MISSING_IN_REPORT:  { bg: "#fff3cd", text: "#856404" },
  MISSING_IN_EXCEL:   { bg: "#d1ecf1", text: "#0c5460" },
  NOT_FOUND:          { bg: "#f8d7da", text: "#721c24" },
};

function Badge({ status }) {
  const c = STATUS_COLOR[status] || { bg: "#e2e3e5", text: "#383d41" };
  return (
    <span style={{
      background: c.bg, color: c.text,
      padding: "2px 8px", borderRadius: 12,
      fontSize: 12, fontWeight: 600,
    }}>
      {status}
    </span>
  );
}

export default function ReportQC() {
  const [mode, setMode]           = useState("single");
  const [excelFile, setExcelFile] = useState(null);
  const [reportFile, setReportFile] = useState(null);
  const [wardNo, setWardNo]       = useState("");
  const [propNo, setPropNo]       = useState("");
  const [partitionNo, setPartitionNo] = useState("");
  const [result, setResult]       = useState(null);
  const [loading, setLoading]     = useState(false);
  const [error, setError]         = useState(null);

  async function runCheck() {
    if (!excelFile || !reportFile) {
      setError("Please select both the Master Excel and the Report PDF.");
      return;
    }
    setLoading(true);
    setError(null);
    setResult(null);

    try {
      // Step 1 — upload Excel
      const ef = new FormData();
      ef.append("file", excelFile);
      const er = await fetch(`${API_BASE}/excel/upload`, { method: "POST", body: ef });
      const ed = await er.json();
      if (!er.ok) throw new Error(ed.detail || "Excel upload failed");

      // Step 2 — upload PDF + run check
      const rf = new FormData();
      rf.append("file", reportFile);
      if (mode === "single") {
        if (wardNo)      rf.append("ward_no",      wardNo);
        if (propNo)      rf.append("prop_no",      propNo);
        // Always send partition_no (empty string = no partition = NaN match)
        rf.append("partition_no", partitionNo);
      }

      const endpoint = mode === "single" ? "/qc/check-single" : "/qc/check-bulk";
      const rr = await fetch(`${API_BASE}${endpoint}`, { method: "POST", body: rf });
      const rd = await rr.json();
      if (!rr.ok) throw new Error(rd.detail || "QC check failed");

      setResult({ mode, ...rd });
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={{ maxWidth: 900, margin: "0 auto", padding: "24px 16px", fontFamily: "system-ui, sans-serif" }}>
      <h2 style={{ marginBottom: 4 }}>Report vs Excel QC Check</h2>
      <p style={{ color: "#666", marginTop: 0 }}>
        Upload the master Excel and the report PDF. The system looks up each property by
        <strong> Ward No + Property No</strong> and compares all 23 columns.
      </p>

      {/* ── Upload form ── */}
      <section style={{ border: "1px solid #ddd", borderRadius: 8, padding: 20, marginBottom: 20 }}>
        {/* Excel */}
        <div style={{ marginBottom: 16 }}>
          <h3 style={{ margin: "0 0 8px" }}>1. Master Excel</h3>
          <input type="file" id="excel-input" accept=".xlsx,.xls"
            onChange={e => setExcelFile(e.target.files[0])} />
        </div>

        {/* Mode */}
        <div style={{ marginBottom: 16 }}>
          <h3 style={{ margin: "0 0 8px" }}>2. Check type</h3>
          <label style={{ marginRight: 20 }}>
            <input type="radio" checked={mode === "single"} onChange={() => setMode("single")} /> Single Property
          </label>
          <label>
            <input type="radio" checked={mode === "bulk"} onChange={() => setMode("bulk")} /> Bulk (multi-page PDF)
          </label>

          {mode === "single" && (
            <div style={{ marginTop: 12 }}>
              <p style={{ margin: "0 0 8px", fontSize: 13, color: "#555" }}>
                These 3 fields are the <strong>lookup key</strong> — the system finds the Excel row by matching all three.
                Leave blank to auto-read from the PDF.
              </p>
              <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
                <div>
                  <label style={{ display: "block", fontSize: 12, color: "#555", marginBottom: 3 }}>Ward No</label>
                  <input id="ward-no-input" type="text" value={wardNo}
                    onChange={e => setWardNo(e.target.value)}
                    placeholder="e.g. D10"
                    style={{ padding: "6px 10px", width: 110, border: "1px solid #ccc", borderRadius: 4 }} />
                </div>
                <div>
                  <label style={{ display: "block", fontSize: 12, color: "#555", marginBottom: 3 }}>New Property No</label>
                  <input id="prop-no-input" type="text" value={propNo}
                    onChange={e => setPropNo(e.target.value)}
                    placeholder="e.g. 2465"
                    style={{ padding: "6px 10px", width: 120, border: "1px solid #ccc", borderRadius: 4 }} />
                </div>
                <div>
                  <label style={{ display: "block", fontSize: 12, color: "#555", marginBottom: 3 }}>New Partition No <em style={{fontWeight:400}}>(blank = none)</em></label>
                  <input id="partition-no-input" type="text" value={partitionNo}
                    onChange={e => setPartitionNo(e.target.value)}
                    placeholder="e.g. 1  (or blank)"
                    style={{ padding: "6px 10px", width: 140, border: "1px solid #ccc", borderRadius: 4 }} />
                </div>
              </div>
            </div>
          )}
        </div>

        {/* PDF */}
        <div style={{ marginBottom: 20 }}>
          <h3 style={{ margin: "0 0 8px" }}>3. Report PDF / RPT</h3>
          <input type="file" id="pdf-input" accept=".pdf,.rpt"
            onChange={e => setReportFile(e.target.files[0])} />
        </div>

        <button id="run-qc-btn" onClick={runCheck}
          disabled={!excelFile || !reportFile || loading}
          style={{
            padding: "10px 28px", fontSize: 15, background: "#4472C4",
            color: "white", border: "none", borderRadius: 6, cursor: "pointer",
          }}>
          {loading ? "Processing…" : "▶ Run QC Check"}
        </button>
      </section>

      {error && (
        <div style={{ background: "#f8d7da", border: "1px solid #f5c2c7", borderRadius: 6,
          padding: "12px 16px", color: "#842029", marginBottom: 16 }}>
          ⚠ {error}
        </div>
      )}

      {/* ── Results ── */}
      {result && result.mode === "single" && <SingleResult result={result} />}
      {result && result.mode === "bulk"   && <BulkResult result={result} />}
    </div>
  );
}


// ─────────────────────────────────────────────────────────────────────────────
// Single result panel
// ─────────────────────────────────────────────────────────────────────────────
function SingleResult({ result }) {
  const exportButtons = (
    <div style={{ marginTop: 24, textAlign: "right", display: "flex", gap: "12px", justifyContent: "flex-end" }}>
      <button onClick={() => exportDiscrepanciesExcel([result])}
        style={{ padding: "8px 16px", background: "#217346", color: "white", border: "none", borderRadius: 4, cursor: "pointer" }}>
        📥 Export to Excel (.xlsx)
      </button>
      <button onClick={() => exportDiscrepanciesPDF([result])}
        style={{ padding: "8px 16px", background: "#d93025", color: "white", border: "none", borderRadius: 4, cursor: "pointer" }}>
        📥 Export to PDF
      </button>
    </div>
  );

  if (!result.found_in_excel) {
    return (
      <div style={{ background: "#fff3cd", border: "1px solid #ffc107", borderRadius: 6, padding: 16 }}>
        <h3 style={{ margin: "0 0 8px 0", color: "#856404" }}>⚠️ Not found in Excel</h3>
        <p style={{ margin: "0 0 16px 0", color: "#856404" }}>{result.message}</p>
        
        <div style={{ fontSize: 13, color: "#333", padding: "12px", background: "#ffffff", borderRadius: "6px", border: "1px solid #e9ecef" }}>
          <h4 style={{ margin: "0 0 12px 0", color: "#4472C4" }}>Data Extracted from PDF (Cannot find matching Excel row)</h4>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: "12px" }}>
            {result.extracted_from_report 
              ? Object.entries(result.extracted_from_report).map(([key, value]) => (
                  <div key={key} style={{ background: "#f8f9fa", padding: "8px", borderRadius: "4px", border: "1px solid #eee" }}>
                    <strong style={{ color: "#555", fontSize: "11px", textTransform: "uppercase" }}>{key}</strong>
                    <div style={{ marginTop: "4px", wordBreak: "break-word" }}>
                      {value !== null && value !== undefined && value !== "" ? String(value) : <em style={{color: "#aaa"}}>N/A</em>}
                    </div>
                  </div>
                ))
              : <div>No data extracted.</div>
            }
          </div>
        </div>

        <div style={{ marginTop: 16, background: "white", padding: 12, borderRadius: 6, border: "1px solid #e9ecef" }}>
          <FieldTable rows={
            result.extracted_from_report && Object.keys(result.extracted_from_report).length > 0
              ? Object.entries(result.extracted_from_report).map(([k, v]) => ({
                  field: k,
                  report_value: v !== null && v !== undefined && v !== "" ? String(v) : "MISSING",
                  excel_value: "MISSING",
                  status: "ROW_MISSING_IN_EXCEL"
                }))
              : [{
                  field: "(entire row)",
                  report_value: "Present",
                  excel_value: "MISSING",
                  status: "ROW_MISSING_IN_EXCEL"
                }]
          } 
          rowKey={result.extracted_from_report?.NewPropertyNo || "UNKNOWN"} />
        </div>
        
        {exportButtons}
      </div>
    );
  }

  const matches    = result.results?.filter(r => r.status === "MATCH") ?? [];
  const mismatches = result.results?.filter(r => r.status !== "MATCH") ?? [];

  return (
    <section>
      {/* Summary bar */}
      <div style={{ display: "flex", gap: 12, marginBottom: 16 }}>
        <SummaryCard label="Total Fields Checked" value={result.total_fields_checked} color="#4472C4" />
        <SummaryCard label="Matching"   value={matches.length}    color="#28a745" />
        <SummaryCard label="Mismatches" value={mismatches.length} color="#dc3545" />
        <SummaryCard label="Overall"    value={result.status}     color={result.status === "MATCH" ? "#28a745" : "#dc3545"} />
      </div>

      {/* Key used for lookup & Excel identifiers */}
      <div style={{ fontSize: 13, color: "#333", marginBottom: 16, padding: "12px", background: "#f8f9fa", borderRadius: "6px", border: "1px solid #e9ecef" }}>
        <h4 style={{ margin: "0 0 8px 0", color: "#4472C4" }}>Property Identification</h4>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "8px" }}>
          <div><strong>Correct Owner ID:</strong> <br/>{result.excel_identifiers?.OwnerID || "N/A"}</div>
          <div><strong>Property No:</strong> <br/>{result.key?.NewPropertyNo}</div>
          <div><strong>Ward No:</strong> <br/>{result.key?.NewWardNo}</div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "8px", marginTop: "8px", color: "#666" }}>
          <div><strong>Assessment ID:</strong> <br/>{result.excel_identifiers?.AssesmentID || "N/A"}</div>
          <div><strong>Partition No:</strong> <br/>{result.key?.NewPartitionNo || "none"}</div>
          <div><strong>Old Property No:</strong> <br/>{result.excel_identifiers?.OldPropertyNo || "N/A"}</div>
        </div>
      </div>

      {/* Mismatch table — only show if there are issues */}
      {mismatches.length > 0 ? (
        <>
          <h3 style={{ color: "#dc3545" }}>❌ Mismatched Fields ({mismatches.length})</h3>
          <FieldTable rows={mismatches} rowKey={result.excel_identifiers?.OwnerID || result.key?.NewPropertyNo || "UNKNOWN"} />
        </>
      ) : (
        <h3 style={{ color: "#28a745" }}>✅ All fields matched perfectly!</h3>
      )}

      {/* Full comparison table (collapsed by default to compress UI) */}
      <details style={{ marginTop: 24, border: "1px solid #ddd", borderRadius: 6, padding: 12 }}>
        <summary style={{ cursor: "pointer", fontSize: "1.1em", fontWeight: "bold" }}>
          📋 Full Field Comparison ({result.results?.length} fields)
        </summary>
        <div style={{ marginTop: 12 }}>
          <FieldTable rows={result.results ?? []} rowKey={result.excel_identifiers?.OwnerID || result.key?.NewPropertyNo || "UNKNOWN"} />
        </div>
      </details>
      
      {exportButtons}
    </section>
  );
}


// ─────────────────────────────────────────────────────────────────────────────
// Bulk result panel
// ─────────────────────────────────────────────────────────────────────────────
function BulkResult({ result }) {
  return (
    <section>
      <div style={{ display: "flex", gap: 12, marginBottom: 16 }}>
        <SummaryCard label="Properties Checked" value={result.total_records}    color="#4472C4" />
        <SummaryCard label="Total Mismatches"   value={result.total_mismatches} color="#dc3545" />
      </div>

      {result.records?.map((rec, i) => (
        <details key={i} style={{ border: "1px solid #ddd", borderRadius: 6, marginBottom: 8, padding: 12 }}>
          <summary style={{ cursor: "pointer", fontWeight: 600 }}>
            {rec.key?.NewWardNo} / {rec.key?.NewPropertyNo} &nbsp;—&nbsp;
            <span style={{ color: rec.status === "MATCH" ? "#28a745" : "#dc3545" }}>
              {rec.status}
            </span>
            {" "}({rec.total_mismatches ?? 0} issues)
          </summary>
          <div style={{ marginTop: 12 }}>
            {rec.found_in_excel ? (
              <FieldTable rows={rec.discrepancies ?? []} rowKey={rec.excel_identifiers?.OwnerID || rec.key?.NewPropertyNo || "UNKNOWN"} />
            ) : (
              <FieldTable rows={
                rec.extracted_from_report && Object.keys(rec.extracted_from_report).length > 0
                  ? Object.entries(rec.extracted_from_report).map(([k, v]) => ({
                      field: k,
                      report_value: v !== null && v !== undefined && v !== "" ? String(v) : "MISSING",
                      excel_value: "MISSING",
                      status: "ROW_MISSING_IN_EXCEL"
                    }))
                  : [{
                      field: "(entire row)",
                      report_value: "Present",
                      excel_value: "MISSING",
                      status: "ROW_MISSING_IN_EXCEL"
                    }]
              } 
              rowKey={rec.extracted_from_report?.NewPropertyNo || "UNKNOWN"} />
            )}
          </div>
        </details>
      ))}

      {/* Export Buttons */}
      <div style={{ marginTop: 24, textAlign: "right", display: "flex", gap: "12px", justifyContent: "flex-end" }}>
        <button onClick={() => exportDiscrepanciesExcel(result.records || [])}
          style={{ padding: "8px 16px", background: "#217346", color: "white", border: "none", borderRadius: 4, cursor: "pointer" }}>
          📥 Export Discrepancies to Excel (.xlsx)
        </button>
        <button onClick={() => exportDiscrepanciesPDF(result.records || [])}
          style={{ padding: "8px 16px", background: "#d93025", color: "white", border: "none", borderRadius: 4, cursor: "pointer" }}>
          📥 Export Discrepancies to PDF
        </button>
      </div>
    </section>
  );
}


// ─────────────────────────────────────────────────────────────────────────────
// Export Logic
// ─────────────────────────────────────────────────────────────────────────────
function formatExportData(records) {
  const data = [];
  records.forEach(rec => {
    // Determine a Row Key (use Owner ID if available, otherwise Property No)
    let rowKey = rec.excel_identifiers?.OwnerID;
    if (!rowKey || rowKey === "NOT_FOUND") {
      rowKey = rec.key?.NewPropertyNo || rec.extracted_from_report?.NewPropertyNo || "UNKNOWN";
    }

    if (!rec.found_in_excel) {
      if (rec.extracted_from_report && Object.keys(rec.extracted_from_report).length > 0) {
        Object.entries(rec.extracted_from_report).forEach(([key, value]) => {
          data.push({
            "Row Key": rowKey,
            "Field": key,
            "Report Value": value !== null && value !== undefined && value !== "" ? String(value) : "MISSING",
            "Excel Value": "MISSING",
            "Issue Type": "ROW_MISSING_IN_EXCEL"
          });
        });
      } else {
        data.push({
          "Row Key": rowKey,
          "Field": "(entire row)",
          "Report Value": "Present",
          "Excel Value": "MISSING",
          "Issue Type": "ROW_MISSING_IN_EXCEL"
        });
      }
      return;
    }

    if (rec.discrepancies && rec.discrepancies.length > 0) {
      rec.discrepancies.forEach(d => {
        data.push({
          "Row Key": rowKey,
          "Field": d.field,
          "Report Value": d.report_value === null ? "MISSING" : String(d.report_value),
          "Excel Value": d.excel_value === null ? "MISSING" : String(d.excel_value),
          "Issue Type": d.status || d.issue_type || "VALUE_MISMATCH"
        });
      });
    } else {
        // If they want to export even perfect matches (usually bulk just exports discrepancies)
        // We will just skip rows with no discrepancies to match the issue screenshot
    }
  });
  return data;
}

function exportDiscrepanciesExcel(records) {
  const data = formatExportData(records);
  if (data.length === 0) {
    alert("No discrepancies to export!");
    return;
  }
  const ws = XLSX.utils.json_to_sheet(data);
  
  // Basic column widths
  ws["!cols"] = [{ wch: 15 }, { wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 25 }];
  
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "QC Report");
  XLSX.writeFile(wb, "Report_QC_Results.xlsx");
}

function exportDiscrepanciesPDF(records) {
  const data = formatExportData(records);
  if (data.length === 0) {
    alert("No discrepancies to export!");
    return;
  }

  const doc = new jsPDF();
  doc.setFontSize(16);
  doc.text("QC Discrepancy Report", 14, 15);
  doc.setFontSize(10);
  doc.text(`Generated on: ${new Date().toLocaleString()}`, 14, 22);

  const tableColumn = ["Row Key", "Field", "Report Value", "Excel Value", "Issue Type"];
  const tableRows = data.map(item => [
    item["Row Key"],
    item["Field"],
    item["Report Value"],
    item["Excel Value"],
    item["Issue Type"]
  ]);

  doc.autoTable({
    startY: 28,
    head: [tableColumn],
    body: tableRows,
    theme: 'grid',
    headStyles: { fillColor: [68, 114, 196] },
    styles: { fontSize: 9 }
  });

  doc.save("Report_QC_Results.pdf");
}

// ─────────────────────────────────────────────────────────────────────────────
// Reusable components
// ─────────────────────────────────────────────────────────────────────────────
function SummaryCard({ label, value, color }) {
  return (
    <div style={{
      flex: 1, border: `2px solid ${color}`, borderRadius: 8,
      padding: "12px 16px", textAlign: "center",
    }}>
      <div style={{ fontSize: 22, fontWeight: 700, color }}>{value}</div>
      <div style={{ fontSize: 12, color: "#555" }}>{label}</div>
    </div>
  );
}

function FieldTable({ rows, rowKey }) {
  if (!rows || rows.length === 0) return null;
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
        <thead>
          <tr style={{ background: "#4472C4", color: "white" }}>
            <th style={TH}>Row Key</th>
            <th style={TH}>Field</th>
            <th style={TH}>Report Value</th>
            <th style={TH}>Excel Value</th>
            <th style={TH}>Issue Type</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => {
            const isOk = r.status === "MATCH";
            return (
              <tr key={i} style={{ background: isOk ? "white" : "#fff0f0" }}>
                <td style={TD}>{rowKey || "UNKNOWN"}</td>
                <td style={{ ...TD, fontFamily: "monospace", fontWeight: 500 }}>{r.field}</td>
                <td style={{ ...TD, color: isOk ? "inherit" : "#c0392b", fontWeight: isOk ? 400 : 600 }}>
                  {fmt(r.report_value)}
                </td>
                <td style={TD}>{fmt(r.excel_value)}</td>
                <td style={TD}><Badge status={r.status || r.issue_type} /></td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

const fmt = v => (v === null || v === undefined ? <em style={{ color: "#aaa" }}>—</em> : String(v));
const TH = { border: "1px solid #3a5fa0", padding: "8px 12px", textAlign: "left" };
const TD = { border: "1px solid #ddd", padding: "7px 12px" };
