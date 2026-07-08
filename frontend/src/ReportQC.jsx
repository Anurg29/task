import { useState } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import "jspdf-autotable";

const API_BASE = import.meta.env.VITE_API_URL || "http://localhost:8000";

const STATUS_COLOR = {
  MATCH:              { bg: "var(--success-bg)", text: "var(--success)" },
  VALUE_MISMATCH:     { bg: "var(--error-bg)", text: "var(--error)" },
  MISSING_IN_REPORT:  { bg: "var(--warning-bg)", text: "var(--warning)" },
  MISSING_IN_EXCEL:   { bg: "rgba(14, 165, 233, 0.2)", text: "#0ea5e9" },
  NOT_FOUND:          { bg: "var(--error-bg)", text: "var(--error)" },
  MINOR_TYPO:         { bg: "rgba(234, 179, 8, 0.2)", text: "#eab308" }, // Yellow
};

function Badge({ status }) {
  const c = STATUS_COLOR[status] || { bg: "rgba(0,0,0,0.1)", text: "var(--text-muted)" };
  return (
    <span className="badge" style={{ background: c.bg, color: c.text }}>
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
  const [targetWards, setTargetWards] = useState(""); // For bulk mode filtering
  const [result, setResult]       = useState(null);
  const [loading, setLoading]     = useState(false);
  const [error, setError]         = useState(null);

  // Progress tracking
  const [progressMsg, setProgressMsg] = useState("");
  const [progressPct, setProgressPct] = useState(0);

  async function runCheck() {
    if (!excelFile || !reportFile) {
      setError("Please select both the Master Excel and the Report PDF.");
      return;
    }
    setLoading(true);
    setError(null);
    setResult(null);
    setProgressMsg("Starting scan...");
    setProgressPct(0);

    let ws = null;
    let clientId = null;

    try {
      if (mode === "bulk") {
        clientId = Math.random().toString(36).substring(2, 15);
        const wsProtocol = API_BASE.startsWith("https") ? "wss" : "ws";
        const wsHost = API_BASE.replace(/^https?:\/\//, "");
        ws = new WebSocket(`${wsProtocol}://${wsHost}/ws/progress/${clientId}`);
        ws.onmessage = (event) => {
          const data = JSON.parse(event.data);
          const scanned = data.scanned || 0;
          const total = data.total || 1;
          const pct = Math.round((scanned / total) * 100);
          setProgressPct(pct);
          setProgressMsg(`Scanning Page ${scanned} of ${total}... ${pct}% Completed`);
        };
      }

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
      } else if (mode === "bulk") {
        if (targetWards) rf.append("target_wards", targetWards);
        if (clientId)    rf.append("client_id", clientId);
      }

      const endpoint = mode === "single" ? "/qc/check-single" : "/qc/check-bulk";
      const rr = await fetch(`${API_BASE}${endpoint}`, { method: "POST", body: rf });
      const rd = await rr.json();
      if (!rr.ok) throw new Error(rd.detail || "QC check failed");

      setResult({ mode, ...rd });
    } catch (e) {
      setError(e.message);
    } finally {
      if (ws) ws.close();
      setLoading(false);
    }
  }

  return (
    <div className="app-container animate-fade-in">
      <header className="header animate-fade-in">
        <h1>TaxLedger QC</h1>
        <p style={{ color: "var(--text-muted)", marginTop: "12px", lineHeight: "1.6", maxWidth: "800px", margin: "12px auto 0" }}>
          An intelligent reconciliation platform designed to cross-verify your master Excel database against PDF tax ledgers. 
          The system instantly audits all 23 tax heads and property details, automatically detecting numeric discrepancies, missing records, and typographical errors in real-time.
        </p>
      </header>

      {/* ── Upload form ── */}
      <section className="glass-panel upload-section">
        {/* Excel */}
        <div className="upload-card">
          <h3>📊 1. Master Excel</h3>
          <label className="file-upload-wrapper">
            <div className="upload-icon">📁</div>
            <div>{excelFile ? excelFile.name : "Click or drag Excel file (.xlsx)"}</div>
            <input type="file" accept=".xlsx,.xls" onChange={e => setExcelFile(e.target.files[0])} />
          </label>
        </div>

        {/* PDF & Controls */}
        <div className="upload-card">
          <h3>📑 2. Report PDF & Mode</h3>
          
          <div className="mode-selector">
            <label className={`mode-radio ${mode === "single" ? "active" : ""}`}>
              <input type="radio" checked={mode === "single"} onChange={() => setMode("single")} /> 
              Single Property
            </label>
            <label className={`mode-radio ${mode === "bulk" ? "active" : ""}`}>
              <input type="radio" checked={mode === "bulk"} onChange={() => setMode("bulk")} /> 
              Bulk (multi-page)
            </label>
          </div>

          <label className="file-upload-wrapper" style={{ padding: "20px" }}>
            <div className="upload-icon" style={{ fontSize: "1.5rem" }}>📄</div>
            <div style={{ fontSize: "0.9rem" }}>{reportFile ? reportFile.name : "Select PDF Report (.pdf)"}</div>
            <input type="file" accept=".pdf,.rpt" onChange={e => setReportFile(e.target.files[0])} />
          </label>

          {mode === "single" && (
            <div className="single-mode-inputs animate-fade-in">
              <input className="input-modern" type="text" value={wardNo}
                onChange={e => setWardNo(e.target.value)} placeholder="Ward (e.g. D10)" />
              <input className="input-modern" type="text" value={propNo}
                onChange={e => setPropNo(e.target.value)} placeholder="Property (e.g. 2465)" />
              <input className="input-modern" type="text" value={partitionNo}
                onChange={e => setPartitionNo(e.target.value)} placeholder="Partition (opt)" />
            </div>
          )}

          {mode === "bulk" && (
            <div className="single-mode-inputs animate-fade-in">
              <input className="input-modern" type="text" value={targetWards}
                onChange={e => setTargetWards(e.target.value)} 
                placeholder="Filter by Wards (e.g. A1, A2 or A1-A10)" 
                style={{ width: "100%" }} />
            </div>
          )}
        </div>
      </section>

      {error && (
        <div className="error-banner animate-fade-in">
          <span style={{ fontSize: "1.2rem" }}>⚠️</span>
          <span>{error}</span>
        </div>
      )}

      <div className="action-row">
        <button className="btn-primary run-btn" onClick={runCheck} disabled={loading}>
          {loading ? (
            <span className="spinner"></span>
          ) : "▶ Run QC Check"}
        </button>
      </div>

      {/* ── Results ── */}
      {result && (
        <div className="results-container glass-panel animate-fade-in" style={{ padding: "32px" }}>
          {result.mode === "single" ? <SingleResult result={result} /> : <BulkResult result={result} />}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Single result panel
// ─────────────────────────────────────────────────────────────────────────────
function SingleResult({ result }) {
  const exportButtons = (
    <div style={{ marginTop: 32, display: "flex", gap: "16px", justifyContent: "flex-end" }}>
      <button className="btn-primary" onClick={() => exportDiscrepanciesExcel([result])}
        style={{ background: "linear-gradient(135deg, #10b981, #059669)" }}>
        📥 Export Excel
      </button>
      <button className="btn-primary" onClick={() => exportDiscrepanciesPDF([result])}
        style={{ background: "linear-gradient(135deg, #ef4444, #dc2626)" }}>
        📥 Export PDF
      </button>
    </div>
  );

  if (!result.found_in_excel) {
    return (
      <div className="animate-fade-in">
        <h3 style={{ color: "var(--warning)", marginBottom: "8px" }}>⚠️ Not found in Excel</h3>
        <p style={{ marginBottom: "24px" }}>{result.message}</p>
        
        <div className="glass-card" style={{ padding: "20px", marginBottom: "24px" }}>
          <h4 style={{ margin: "0 0 16px 0", color: "var(--primary)" }}>Data Extracted from PDF</h4>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: "16px" }}>
            {result.extracted_from_report 
              ? Object.entries(result.extracted_from_report).map(([key, value]) => (
                  <div key={key}>
                    <div style={{ fontSize: "0.75rem", textTransform: "uppercase", color: "var(--text-muted)", marginBottom: "4px" }}>{key}</div>
                    <div style={{ fontWeight: "500" }}>
                      {value !== null && value !== undefined && value !== "" ? String(value) : <em style={{opacity: 0.5}}>N/A</em>}
                    </div>
                  </div>
                ))
              : <div>No data extracted.</div>
            }
          </div>
        </div>

        <div className="table-container">
          <FieldTable records={[result]} />
        </div>
        
        {exportButtons}
      </div>
    );
  }

  const matches    = result.results?.filter(r => r.status === "MATCH") ?? [];
  const mismatches = result.results?.filter(r => r.status !== "MATCH") ?? [];

  return (
    <section className="animate-fade-in">
      {/* Summary bar */}
      <div className="summary-cards">
        <div className="glass-card summary-card">
          <div className="value" style={{ color: "var(--primary)" }}>{result.total_fields_checked}</div>
          <div className="label">Fields Checked</div>
        </div>
        <div className="glass-card summary-card">
          <div className="value" style={{ color: "var(--success)" }}>{matches.length}</div>
          <div className="label">Matching</div>
        </div>
        <div className="glass-card summary-card">
          <div className="value" style={{ color: "var(--error)" }}>{mismatches.length}</div>
          <div className="label">Mismatches</div>
        </div>
      </div>

      {/* Key used for lookup & Excel identifiers */}
      <div className="glass-card" style={{ padding: "20px", marginBottom: "32px" }}>
        <h4 style={{ margin: "0 0 16px 0", color: "var(--primary)" }}>Property Identification</h4>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: "16px" }}>
          <div>
            <div style={{ fontSize: "0.75rem", textTransform: "uppercase", color: "var(--text-muted)" }}>Owner ID</div>
            <div style={{ fontWeight: "500" }}>{result.excel_identifiers?.OwnerID || "N/A"}</div>
          </div>
          <div>
            <div style={{ fontSize: "0.75rem", textTransform: "uppercase", color: "var(--text-muted)" }}>Property No</div>
            <div style={{ fontWeight: "500" }}>{result.key?.NewPropertyNo}</div>
          </div>
          <div>
            <div style={{ fontSize: "0.75rem", textTransform: "uppercase", color: "var(--text-muted)" }}>Ward No</div>
            <div style={{ fontWeight: "500" }}>{result.key?.NewWardNo}</div>
          </div>
          <div>
            <div style={{ fontSize: "0.75rem", textTransform: "uppercase", color: "var(--text-muted)" }}>Assessment ID</div>
            <div style={{ fontWeight: "500" }}>{result.excel_identifiers?.AssesmentID || "N/A"}</div>
          </div>
        </div>
      </div>

      {/* Mismatch table */}
      <div style={{ marginBottom: "24px" }}>
        {mismatches.length > 0 ? (
          <>
            <h3 style={{ color: "var(--error)", marginBottom: "16px", display: "flex", alignItems: "center", gap: "8px" }}>
              ❌ Property Details ({mismatches.length} Mismatches)
            </h3>
            <div className="table-container">
              <FieldTable records={[result]} />
            </div>
          </>
        ) : (
          <div className="glass-card" style={{ padding: "24px", textAlign: "center" }}>
            <h3 style={{ color: "var(--success)", margin: 0 }}>✨ All fields matched perfectly!</h3>
            <div className="table-container" style={{ marginTop: "24px" }}>
              <FieldTable records={[result]} />
            </div>
          </div>
        )}
      </div>
      
      {exportButtons}
    </section>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Bulk result panel
// ─────────────────────────────────────────────────────────────────────────────
function BulkResult({ result }) {
  const [showOnlyMismatches, setShowOnlyMismatches] = useState(false);
  
  // Filter records based on toggle
  const displayedRecords = result.records?.filter(rec => 
    showOnlyMismatches ? (rec.total_mismatches > 0 || rec.overall_result === "MISSING_IN_EXCEL") : true
  );

  return (
    <section className="animate-fade-in">
      <div className="summary-cards">
        <div className="glass-card summary-card">
          <div className="value" style={{ color: "var(--primary)" }}>{result.total_records}</div>
          <div className="label">Properties Parsed</div>
        </div>
        <div className="glass-card summary-card">
          <div className="value" style={{ color: "var(--error)" }}>{result.total_mismatches}</div>
          <div className="label">Total Mismatches Found</div>
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: "16px" }}>
        <label style={{ display: "flex", alignItems: "center", gap: "8px", cursor: "pointer", color: "var(--text-muted)", fontSize: "0.9rem" }}>
          <input 
            type="checkbox" 
            checked={showOnlyMismatches} 
            onChange={(e) => setShowOnlyMismatches(e.target.checked)}
            style={{ width: "16px", height: "16px", accentColor: "var(--primary)" }}
          />
          Show Only Mismatches
        </label>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: "12px", marginBottom: "32px" }}>
        {displayedRecords?.length === 0 && showOnlyMismatches && (
          <div className="glass-card" style={{ padding: "32px", textAlign: "center", color: "var(--success)" }}>
            ✨ All records are matching perfectly!
          </div>
        )}
        {displayedRecords?.slice(0, 10).map((rec, i) => (
          <details key={i} className="glass-card">
            <summary style={{ cursor: "pointer", padding: "16px", display: "flex", alignItems: "center", gap: "12px", outline: "none" }}>
              <strong style={{ minWidth: "140px" }}>{rec.key?.NewWardNo} / {rec.key?.NewPropertyNo}</strong>
              <Badge status={rec.overall_result} />
              <span style={{ color: "var(--text-muted)", fontSize: "0.9rem", marginLeft: "auto" }}>
                {rec.total_mismatches ?? 0} issues
              </span>
            </summary>
            
            <div className="table-container" style={{ border: "none", borderTop: "1px solid var(--surface-border)", borderRadius: "0 0 12px 12px" }}>
              <FieldTable records={[rec]} />
            </div>
          </details>
        ))}
        {result.records?.length > 10 && (
          <div className="glass-card" style={{ padding: "16px", textAlign: "center", color: "var(--text-muted)" }}>
            Showing first <strong>10</strong> records out of <strong>{result.records.length}</strong>. 
            <br/>Please click <strong>Export to Excel</strong> to view all records.
          </div>
        )}
      </div>

      {/* Export Buttons */}
      <div style={{ display: "flex", gap: "16px", justifyContent: "flex-end" }}>
        <button className="btn-primary" onClick={() => exportDiscrepanciesExcel(result.records || [])}
          style={{ background: "linear-gradient(135deg, #10b981, #059669)" }}>
          📥 Export to Excel
        </button>
        <button className="btn-primary" onClick={() => exportDiscrepanciesPDF(result.records || [])}
          style={{ background: "linear-gradient(135deg, #ef4444, #dc2626)" }}>
          📥 Export to PDF
        </button>
      </div>
    </section>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Export Logic
// ─────────────────────────────────────────────────────────────────────────────
function formatExportData(records) {
  return records.map(rec => {
    const row = {};
    const ownerId = rec.excel_identifiers?.OwnerID;
    
    row["Owner ID"] = (ownerId && ownerId !== "NOT_FOUND") ? ownerId : "N/A";
    row["Ward No"] = rec.key?.NewWardNo || rec.extracted_from_report?.NewWardNo || "";
    row["Property No"] = rec.key?.NewPropertyNo || rec.extracted_from_report?.NewPropertyNo || "";
    row["Partition No"] = rec.key?.NewPartitionNo || rec.extracted_from_report?.NewPartitionNo || "";
    row["Overall Result"] = rec.overall_result;

    if (!rec.found_in_excel) {
      if (rec.extracted_from_report) {
        Object.entries(rec.extracted_from_report).forEach(([k, v]) => {
          if (!["NewWardNo", "NewPropertyNo", "NewPartitionNo"].includes(k)) {
            row[`${k} (PDF)`] = v !== null && v !== undefined && v !== "" ? String(v) : "MISSING";
            row[`${k} (Excel)`] = "N/A";
            row[`${k} (Result)`] = "MISSING_IN_EXCEL";
          }
        });
      }
    } else {
      if (rec.results) {
        rec.results.forEach(res => {
          const field = res.field;
          if (!["NewWardNo", "NewPropertyNo", "NewPartitionNo", "OwnerID"].includes(field)) {
            row[`${field} (PDF)`] = res.report_value !== null ? String(res.report_value) : "MISSING";
            row[`${field} (Excel)`] = res.excel_value !== null ? String(res.excel_value) : "MISSING";
            row[`${field} (Result)`] = res.status;
          }
        });
      }
    }
    return row;
  });
}

async function exportDiscrepanciesExcel(records) {
  const data = formatExportData(records);
  if (data.length === 0) {
    alert("No records to export!");
    return;
  }
  
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("QC Report");
  
  if (data.length > 0) {
    const columns = Object.keys(data[0]).map(key => ({ header: key, key: key, width: 20 }));
    worksheet.columns = columns;
    
    // Style Header
    worksheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } }; // Indigo
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    // Add rows and style mismatches
    data.forEach(rowData => {
      const row = worksheet.addRow(rowData);
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const colHeader = columns[colNumber - 1].header;
        const val = cell.value;
        
        if (colHeader.includes("Result") || colHeader.includes("Status") || colHeader.includes("Issue Type")) {
           if (val === "MATCH" || val === "OK") {
             cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } }; // Light Green
             cell.font = { color: { argb: 'FF065F46' }, bold: true };
           } else if (val === "MINOR_TYPO") {
             cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF08A' } }; // Light Yellow
             cell.font = { color: { argb: 'FF854D0E' }, bold: true };
           } else {
             cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } }; // Light Red
             cell.font = { color: { argb: 'FF991B1B' }, bold: true };
           }
        }
      });
    });
  }

  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), "Report_QC_Results.xlsx");
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
    headStyles: { fillColor: [99, 102, 241] },
    styles: { fontSize: 9 }
  });
  doc.save("Report_QC_Results.pdf");
}

// ─────────────────────────────────────────────────────────────────────────────
// Reusable components
// ─────────────────────────────────────────────────────────────────────────────
function FieldTable({ records }) {
  if (!records || records.length === 0) return null;
  
  // Reuse the export formatting to get wide rows for display
  const data = formatExportData(records);
  if (data.length === 0) return null;
  
  const columns = Object.keys(data[0]);

  return (
    <div style={{ overflowX: "auto" }}>
      <table className="modern-table" style={{ whiteSpace: "nowrap" }}>
        <thead>
          <tr>
            {columns.map(col => <th key={col}>{col}</th>)}
          </tr>
        </thead>
        <tbody>
          {data.map((row, i) => (
            <tr key={i}>
              {columns.map(col => {
                const val = row[col];
                if (col.includes("Result") || col.includes("Status") || col.includes("Issue Type")) {
                  const isMinorTypo = val === "MINOR_TYPO";
                  return (
                    <td key={col} style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                      <Badge status={val} />
                      {isMinorTypo && (
                        <button 
                          className="btn-primary" 
                          style={{ padding: "4px 8px", fontSize: "0.75rem", background: "linear-gradient(135deg, #10b981, #059669)" }}
                          onClick={async () => {
                            const pdfText = row[col.replace(" (Result)", " (PDF)")];
                            const excelText = row[col.replace(" (Result)", " (Excel)")];
                            try {
                              await fetch(`${API_BASE}/qc/approve-typo`, {
                                method: "POST",
                                headers: { "Content-Type": "application/json" },
                                body: JSON.stringify({ pdf_text: pdfText, excel_text: excelText })
                              });
                              alert(`Approved typo: ${pdfText} -> ${excelText}\nIt will be matched correctly next time!`);
                            } catch(e) {
                              alert("Error approving typo.");
                            }
                          }}
                        >
                          Approve
                        </button>
                      )}
                    </td>
                  );
                }
                
                let style = { color: val === "MISSING" || val === "N/A" ? "var(--text-muted)" : "inherit" };
                if (col.endsWith(" (PDF)")) {
                  const resCol = col.replace(" (PDF)", " (Result)");
                  if (row[resCol] && !["MATCH", "MINOR_TYPO", "OK"].includes(row[resCol])) {
                     style.color = "var(--error)";
                     style.fontWeight = "bold";
                  }
                } else if (col.endsWith(" (Excel)")) {
                  const resCol = col.replace(" (Excel)", " (Result)");
                  if (row[resCol] && !["MATCH", "MINOR_TYPO", "OK"].includes(row[resCol])) {
                     style.color = "var(--success)";
                     style.fontWeight = "bold";
                  }
                }

                return (
                  <td key={col} style={style}>
                    {fmt(val)}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

const fmt = v => (v === null || v === undefined ? <em style={{ opacity: 0.5 }}>—</em> : String(v));
