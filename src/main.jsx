import React, { useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import ExcelJS from "exceljs";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import { Check, ChevronDown, Download, FileSpreadsheet, FileText, LockKeyhole, Pencil, RotateCcw, Sparkles, Upload } from "lucide-react";
import { cleanGrid, normalizeHeader, safeBaseName } from "./cleaning";
import "./styles.css";

const readRows = (worksheet) => {
  const data = [];
  for (let r = 1; r <= worksheet.rowCount; r += 1) {
    const row = [];
    for (let c = 1; c <= worksheet.columnCount; c += 1) {
      const value = worksheet.getCell(r, c).value;
      row.push(value && typeof value === "object" ? (value.text ?? value.result ?? value.richText?.map((x) => x.text).join("") ?? String(value)) : (value ?? ""));
    }
    data.push(row);
  }
  return data;
};

function App() {
  const inputRef = useRef(null);
  const [source, setSource] = useState(null);
  const [workbook, setWorkbook] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [fileName, setFileName] = useState("cleaned_students");
  const [editing, setEditing] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const summary = useMemo(() => sheets.reduce((acc, sheet) => ({
    rows: acc.rows + Math.max(0, sheet.rows.length - (sheet.headerIndex ?? 0) - 1),
    emails: acc.emails + sheet.stats.emails,
    phones: acc.phones + sheet.stats.phones,
    rolls: acc.rolls + sheet.stats.rolls,
    warnings: acc.warnings + sheet.warnings.length,
  }), { rows: 0, emails: 0, phones: 0, rolls: 0, warnings: 0 }), [sheets]);

  async function loadFile(file) {
    if (!file || !file.name.toLowerCase().endsWith(".xlsx")) { setError("Please choose an .xlsx Excel file."); return; }
    setBusy(true); setError("");
    try {
      const bytes = await file.arrayBuffer();
      const nextWorkbook = new ExcelJS.Workbook();
      await nextWorkbook.xlsx.load(bytes);
      const parsed = nextWorkbook.worksheets.map((worksheet) => ({ name: worksheet.name, ...cleanGrid(readRows(worksheet)) }));
      if (!parsed.length) throw new Error("The workbook has no worksheets.");
      setSource(file); setWorkbook(nextWorkbook); setSheets(parsed); setActiveSheet(0); setEditing(false);
      setFileName(`cleaned_${file.name.replace(/\.xlsx$/i, "")}`);
    } catch (cause) { setError(cause?.message || "This workbook could not be read."); }
    finally { setBusy(false); }
  }

  function updateCell(rowIndex, columnIndex, value) {
    setSheets((current) => current.map((sheet, index) => index === activeSheet
      ? { ...sheet, rows: sheet.rows.map((row, r) => r === rowIndex ? row.map((cell, c) => c === columnIndex ? value : cell) : row) }
      : sheet));
  }

  function applyRowsToWorkbook() {
    sheets.forEach((sheetData, index) => {
      const worksheet = workbook.worksheets[index];
      const originalHeaders = readRows(worksheet)[sheetData.headerIndex ?? 0]?.map(normalizeHeader) ?? [];
      const paidColumn = originalHeaders.indexOf("paid amount") + 1;
      if (paidColumn > 0) worksheet.spliceColumns(paidColumn, 1);
      sheetData.rows.forEach((row, r) => row.forEach((value, c) => {
        const cell = worksheet.getCell(r + 1, c + 1);
        cell.value = value;
        if (r > (sheetData.headerIndex ?? 0) && normalizeHeader(sheetData.rows[sheetData.headerIndex ?? 0]?.[c]) === "roll") cell.numFmt = "@";
      }));
      // Fit output columns to their longest visible value, within practical limits.
      const columnCount = Math.max(0, ...sheetData.rows.map((row) => row.length));
      for (let c = 0; c < columnCount; c += 1) {
        const longest = Math.max(0, ...sheetData.rows.map((row) => String(row[c] ?? "").length));
        worksheet.getColumn(c + 1).width = Math.min(48, Math.max(12, longest + 2));
      }
      worksheet.eachRow((row) => { row.alignment = { ...row.alignment, vertical: "middle" }; });
    });
  }

  async function downloadXlsx() {
    setBusy(true); setError("");
    try {
      applyRowsToWorkbook();
      const buffer = await workbook.xlsx.writeBuffer();
      downloadBlob(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `${safeBaseName(fileName)}.xlsx`);
    } catch (cause) { setError(cause?.message || "Could not create the Excel file."); }
    finally { setBusy(false); }
  }

  function downloadPdf() {
    setBusy(true); setError("");
    try {
      const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
      sheets.forEach((sheet, index) => {
        if (index) doc.addPage("a4", "landscape");
        doc.setFont("helvetica", "bold"); doc.setFontSize(18); doc.setTextColor(28, 45, 39); doc.text(sheet.name, 40, 38);
        doc.setFont("helvetica", "normal"); doc.setFontSize(9); doc.setTextColor(100, 116, 108); doc.text(`Cleaned export • ${new Date().toLocaleDateString()}`, 40, 54);
        const header = sheet.headerIndex ?? 0;
        autoTable(doc, { startY: 68, head: [sheet.rows[header]], body: sheet.rows.slice(header + 1), theme: "plain", margin: { left: 40, right: 40 },
          styles: { font: "helvetica", fontSize: 8, cellPadding: 6, lineColor: [220, 228, 223], lineWidth: 0.5, textColor: [38, 52, 47], overflow: "linebreak" },
          headStyles: { fillColor: [28, 45, 39], textColor: 255, fontStyle: "bold" }, alternateRowStyles: { fillColor: [244, 247, 245] },
          didDrawPage: (data) => { doc.setFontSize(8); doc.setTextColor(120); doc.text(`Page ${doc.getNumberOfPages()}`, data.settings.margin.left, doc.internal.pageSize.height - 18); } });
      });
      doc.save(`${safeBaseName(fileName)}.pdf`);
    } catch (cause) { setError(cause?.message || "Could not create the PDF."); }
    finally { setBusy(false); }
  }

  const current = sheets[activeSheet];
  return <main>
    <header className="topbar"><a className="brand" href="#"><span><Sparkles size={18}/></span> Student File Cleaner</a><div className="privacy"><LockKeyhole size={15}/> Files stay on your device</div></header>
    <section className="hero">
      <div className="eyebrow">PRIVATE • FAST • NO SIGN-IN</div>
      <h1>Clean student files.<br/><em>Keep the important parts.</em></h1>
      <p>Mask private contact details, standardize roll numbers, and remove payment data—in one careful pass.</p>
    </section>

    {!source ? <section className={`dropzone ${dragging ? "dragging" : ""}`} onDragOver={(e) => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={(e) => { e.preventDefault(); setDragging(false); loadFile(e.dataTransfer.files[0]); }}>
      <input ref={inputRef} hidden type="file" accept=".xlsx" onChange={(e) => loadFile(e.target.files[0])}/>
      <div className="upload-icon"><Upload size={28}/></div><h2>{busy ? "Reading workbook…" : "Drop your Excel file here"}</h2><p>or choose a file from your computer</p>
      <button className="primary" onClick={() => inputRef.current?.click()} disabled={busy}><FileSpreadsheet size={18}/> Select Excel file</button><small>.XLSX files only • Processed locally in your browser</small>
    </section> : <>
      <section className="filebar">
        <div className="filemeta"><span className="excel"><FileSpreadsheet size={21}/></span><div><strong>{source.name}</strong><small>{(source.size / 1024).toFixed(1)} KB • {sheets.length} sheet{sheets.length === 1 ? "" : "s"}</small></div></div>
        <button className="ghost" onClick={() => { setSource(null); setWorkbook(null); setSheets([]); setError(""); }}><RotateCcw size={16}/> Start over</button>
      </section>

      <section className="results">
        <div className="result-title"><div><span className="success"><Check size={18}/></span><div><h2>Your file is clean</h2><p>{summary.rows} student records processed successfully.</p></div></div><button className={`edit-toggle ${editing ? "active" : ""}`} onClick={() => setEditing(!editing)}><Pencil size={15}/>{editing ? "Finish editing" : "Edit data"}</button></div>
        <div className="metrics"><article><strong>{summary.emails}</strong><span>Emails masked</span></article><article><strong>{summary.phones}</strong><span>Phones masked</span></article><article><strong>{summary.rolls}</strong><span>Rolls cleaned</span></article><article><strong>Removed</strong><span>Paid Amount</span></article></div>
        {summary.warnings > 0 && <div className="warning">{summary.warnings} item{summary.warnings === 1 ? "" : "s"} kept unchanged to prevent data loss.</div>}
      </section>

      <section className="preview-card">
        <div className="preview-head"><div><h2>Preview</h2><p>{editing ? "Click any cell below to make a correction." : "Review the cleaned data before downloading."}</p></div>{sheets.length > 1 && <label className="sheet-select">Sheet<select value={activeSheet} onChange={(e) => setActiveSheet(Number(e.target.value))}>{sheets.map((s, i) => <option value={i} key={s.name}>{s.name}</option>)}</select><ChevronDown size={14}/></label>}</div>
        <div className="table-wrap"><table><thead><tr>{current?.rows[current.headerIndex ?? 0]?.map((cell, i) => <th key={i}>{String(cell)}</th>)}</tr></thead><tbody>{current?.rows.slice((current.headerIndex ?? 0) + 1).map((row, r) => <tr key={r}>{row.map((cell, c) => <td key={c}>{editing ? <input aria-label={`Row ${r + 2}, ${current.rows[current.headerIndex ?? 0][c]}`} value={String(cell ?? "")} onChange={(e) => updateCell(r + (current.headerIndex ?? 0) + 1, c, e.target.value)}/> : String(cell ?? "")}</td>)}</tr>)}</tbody></table></div>
        <div className="table-foot"><span>Showing all {Math.max(0, (current?.rows.length ?? 1) - (current?.headerIndex ?? 0) - 1)} rows</span><span>{current?.name}</span></div>
      </section>

      <section className="downloads"><div className="name-field"><label htmlFor="filename">Output file name</label><div><input id="filename" value={fileName} onChange={(e) => setFileName(e.target.value)}/><span>.xlsx / .pdf</span></div><small>You can rename it before downloading.</small></div>
        <div className="download-actions"><button className="primary dark" onClick={downloadXlsx} disabled={busy}><Download size={18}/> Download clean XLSX</button><button className="secondary" onClick={downloadPdf} disabled={busy}><FileText size={18}/> Download PDF</button></div>
      </section>
    </>}
    {error && <div className="error" role="alert">{error}</div>}
    <footer>Built for repeatable student-file cleanup. Nothing leaves this browser.</footer>
  </main>;
}

function downloadBlob(blob, name) { const url = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = url; anchor.download = name; anchor.click(); setTimeout(() => URL.revokeObjectURL(url), 1000); }

createRoot(document.getElementById("root")).render(<React.StrictMode><App/></React.StrictMode>);
