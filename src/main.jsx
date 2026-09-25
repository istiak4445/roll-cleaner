import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import ExcelJS from "exceljs";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import JSZip from "jszip";
import {
  Archive,
  Check,
  ChevronDown,
  Database,
  Download,
  ExternalLink,
  FileSpreadsheet,
  FileText,
  FolderArchive,
  Layers,
  Link,
  LockKeyhole,
  Pencil,
  Plus,
  RefreshCw,
  RotateCcw,
  Sparkles,
  Trash2,
  Upload,
  X
} from "lucide-react";
import {
  buildGoogleSheetLookup,
  cleanGrid,
  combineStandardSheets,
  combineWebSheets,
  formatForWebPortal,
  normalizeHeader,
  safeBaseName,
  toGoogleSheetCsvUrl
} from "./cleaning";
import "./styles.css";

const DEFAULT_SHEETS = [
  {
    id: "default-batch-28",
    title: "Batch 28 Google Sheet",
    url: "https://docs.google.com/spreadsheets/d/1d3d2CF4HzIpDyz2I4QRf4zeNGLH7wL2H-mOUPS8BQww/edit?gid=0#gid=0",
    isDefault: true
  }
];

const readRows = (worksheet) => {
  const data = [];
  for (let r = 1; r <= worksheet.rowCount; r += 1) {
    const row = [];
    for (let c = 1; c <= worksheet.columnCount; c += 1) {
      const value = worksheet.getCell(r, c).value;
      row.push(
        value && typeof value === "object"
          ? value.text ?? value.result ?? value.richText?.map((x) => x.text).join("") ?? String(value)
          : value ?? ""
      );
    }
    data.push(row);
  }
  return data;
};

async function fetchSheetCsv(url) {
  const csvUrl = toGoogleSheetCsvUrl(url);
  // Try Vercel / serverless proxy first to guarantee no CORS blocking
  try {
    const apiRes = await fetch(`/api/fetch-sheet?url=${encodeURIComponent(csvUrl)}`);
    if (apiRes.ok) {
      return await apiRes.text();
    }
  } catch (e) {
    // API not running, fall back to direct fetch
  }

  const directRes = await fetch(csvUrl);
  if (!directRes.ok) {
    throw new Error(`Google Sheets fetch failed with status ${directRes.status}`);
  }
  return await directRes.text();
}

function App() {
  const inputRef = useRef(null);
  const addFilesInputRef = useRef(null);

  const [rawWorkbooksData, setRawWorkbooksData] = useState([]);
  const [exportMode, setExportMode] = useState("web_portal"); // default to web_portal as requested
  const [defaultBatch, setDefaultBatch] = useState("Morning");
  const [standardSheets, setStandardSheets] = useState([]);
  const [webSheets, setWebSheets] = useState([]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [fileName, setFileName] = useState("cleaned_students");
  const [editing, setEditing] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  // Google Sheets Management State
  const [googleSheets, setGoogleSheets] = useState(DEFAULT_SHEETS);
  const [googleSheetMap, setGoogleSheetMap] = useState(new Map());
  const [syncingSheets, setSyncingSheets] = useState(false);
  const [sheetsModalOpen, setSheetsModalOpen] = useState(false);
  const [newSheetUrl, setNewSheetUrl] = useState("");
  const [newSheetTitle, setNewSheetTitle] = useState("");
  const [storageType, setStorageType] = useState("local");
  const [lastSyncedTime, setLastSyncedTime] = useState("");

  // Load configured sheets from API or localStorage on mount
  useEffect(() => {
    async function initSheets() {
      let initialSheets = DEFAULT_SHEETS;
      try {
        const local = localStorage.getItem("roll_cleaner_google_sheets");
        if (local) {
          const parsed = JSON.parse(local);
          if (Array.isArray(parsed) && parsed.length > 0) {
            initialSheets = parsed;
          }
        }
      } catch (e) {}

      try {
        const res = await fetch("/api/sheets");
        if (res.ok) {
          const data = await res.json();
          if (data.sheets && Array.isArray(data.sheets) && data.sheets.length > 0) {
            initialSheets = data.sheets;
          }
          if (data.storage) {
            setStorageType(data.storage);
          }
        }
      } catch (e) {}

      setGoogleSheets(initialSheets);
      syncGoogleSheets(initialSheets);
    }

    initSheets();
  }, []);

  async function syncGoogleSheets(sheetsList) {
    setSyncingSheets(true);
    try {
      const mergedMap = new Map();
      const updatedSheets = [];

      for (const sheet of sheetsList) {
        try {
          const csvText = await fetchSheetCsv(sheet.url);
          const lookup = buildGoogleSheetLookup(csvText, sheet.title || sheet.url);
          for (const [key, val] of lookup.entries()) {
            if (!mergedMap.has(key) || val.phoneSource === "gsheet_secondary") {
              mergedMap.set(key, val);
            }
          }
          updatedSheets.push({
            ...sheet,
            status: "connected",
            recordCount: lookup.size,
            lastSynced: new Date().toLocaleTimeString()
          });
        } catch (err) {
          console.error("Error loading sheet:", sheet.url, err);
          updatedSheets.push({
            ...sheet,
            status: "error",
            errorMessage: err.message
          });
        }
      }

      setGoogleSheets(updatedSheets);
      setGoogleSheetMap(mergedMap);
      setLastSyncedTime(new Date().toLocaleTimeString());

      // If files are already loaded, re-evaluate web sheets with updated map
      if (rawWorkbooksData.length > 0) {
        reprocessWebSheets(rawWorkbooksData, mergedMap, defaultBatch);
      }
    } catch (err) {
      setError("Failed to sync Google Sheets: " + err.message);
    } finally {
      setSyncingSheets(false);
    }
  }

  function reprocessWebSheets(workbooks, gMap, batch) {
    const individual = [];
    workbooks.forEach((wbData) => {
      wbData.worksheets.forEach((ws) => {
        const displayName =
          workbooks.length > 1 ? `${wbData.fileName} • ${ws.name}` : ws.name;
        individual.push({
          fileName: wbData.fileName,
          sheetName: ws.name,
          name: displayName,
          ...formatForWebPortal(ws.rawRows, batch, gMap)
        });
      });
    });

    let finalWeb = individual;
    if (individual.length > 1) {
      const combined = combineWebSheets(individual);
      finalWeb = [combined, ...individual];
    }
    setWebSheets(finalWeb);
  }

  async function persistSheets(sheetsList) {
    setGoogleSheets(sheetsList);
    try {
      localStorage.setItem("roll_cleaner_google_sheets", JSON.stringify(sheetsList));
    } catch (e) {}

    try {
      await fetch("/api/sheets", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sheets: sheetsList })
      });
    } catch (e) {}
  }

  function handleAddSheet(e) {
    e.preventDefault();
    if (!newSheetUrl.trim()) return;
    const newSheet = {
      id: "sheet-" + Date.now(),
      title: newSheetTitle.trim() || `Google Sheet ${googleSheets.length + 1}`,
      url: newSheetUrl.trim(),
      createdAt: new Date().toISOString()
    };
    const updated = [...googleSheets, newSheet];
    persistSheets(updated);
    setNewSheetUrl("");
    setNewSheetTitle("");
    syncGoogleSheets(updated);
  }

  function handleRemoveSheet(id) {
    const updated = googleSheets.filter((s) => s.id !== id);
    persistSheets(updated);
    syncGoogleSheets(updated);
  }

  const sheets = exportMode === "standard" ? standardSheets : webSheets;

  const summary = useMemo(() => {
    if (sheets[0]?.isCombined) {
      const m = sheets[0];
      return {
        rows: Math.max(0, m.rows.length - (m.headerIndex ?? 0) - 1),
        emails: m.stats?.emails || 0,
        phones: m.stats?.phones || 0,
        rolls: m.stats?.rolls || 0,
        matchedCount: m.stats?.matchedCount || 0,
        secondaryCount: m.stats?.secondaryCount || 0,
        primaryCount: m.stats?.primaryCount || 0,
        unmatchedCount: m.stats?.unmatchedCount || 0,
        warnings: m.warnings?.length || 0
      };
    }
    return sheets.reduce(
      (acc, sheet) => ({
        rows: acc.rows + Math.max(0, sheet.rows.length - (sheet.headerIndex ?? 0) - 1),
        emails: acc.emails + (sheet.stats?.emails || 0),
        phones: acc.phones + (sheet.stats?.phones || 0),
        rolls: acc.rolls + (sheet.stats?.rolls || 0),
        matchedCount: acc.matchedCount + (sheet.stats?.matchedCount || 0),
        secondaryCount: acc.secondaryCount + (sheet.stats?.secondaryCount || 0),
        primaryCount: acc.primaryCount + (sheet.stats?.primaryCount || 0),
        unmatchedCount: acc.unmatchedCount + (sheet.stats?.unmatchedCount || 0),
        warnings: acc.warnings + (sheet.warnings?.length || 0)
      }),
      {
        rows: 0,
        emails: 0,
        phones: 0,
        rolls: 0,
        matchedCount: 0,
        secondaryCount: 0,
        primaryCount: 0,
        unmatchedCount: 0,
        warnings: 0
      }
    );
  }, [sheets]);

  async function loadFiles(fileList, append = false) {
    const xlsxFiles = Array.from(fileList || []).filter((f) => f.name.toLowerCase().endsWith(".xlsx"));
    if (xlsxFiles.length === 0) {
      setError("Please choose one or more .xlsx Excel files.");
      return;
    }
    setBusy(true);
    setError("");
    try {
      const workbooksToProcess = append ? [...rawWorkbooksData] : [];

      for (const file of xlsxFiles) {
        const bytes = await file.arrayBuffer();
        const nextWorkbook = new ExcelJS.Workbook();
        await nextWorkbook.xlsx.load(bytes);

        const worksheetsData = nextWorkbook.worksheets.map((ws) => ({
          name: ws.name,
          rawRows: readRows(ws)
        }));

        workbooksToProcess.push({
          fileName: file.name,
          workbook: nextWorkbook,
          worksheets: worksheetsData
        });
      }

      const individualStandard = [];
      const individualWeb = [];

      workbooksToProcess.forEach((wbData) => {
        wbData.worksheets.forEach((ws) => {
          const displayName =
            workbooksToProcess.length > 1
              ? `${wbData.fileName} • ${ws.name}`
              : ws.name;

          individualStandard.push({
            fileName: wbData.fileName,
            sheetName: ws.name,
            name: displayName,
            ...cleanGrid(ws.rawRows)
          });

          individualWeb.push({
            fileName: wbData.fileName,
            sheetName: ws.name,
            name: displayName,
            ...formatForWebPortal(ws.rawRows, defaultBatch, googleSheetMap)
          });
        });
      });

      let finalStandard = individualStandard;
      if (individualStandard.length > 1) {
        const combinedStd = combineStandardSheets(individualStandard);
        finalStandard = [combinedStd, ...individualStandard];
      }

      let finalWeb = individualWeb;
      if (individualWeb.length > 1) {
        const combinedWeb = combineWebSheets(individualWeb);
        finalWeb = [combinedWeb, ...individualWeb];
      }

      setRawWorkbooksData(workbooksToProcess);
      setStandardSheets(finalStandard);
      setWebSheets(finalWeb);
      setActiveSheet(0);
      setEditing(false);

      if (workbooksToProcess.length > 1) {
        const totalCount = finalWeb[0].rows.length - 1;
        setFileName(`cleaned_combined_master_${totalCount}_students`);
      } else {
        setFileName(`cleaned_${workbooksToProcess[0].fileName.replace(/\.xlsx$/i, "")}`);
      }
    } catch (cause) {
      setError(cause?.message || "Could not read workbook file(s).");
    } finally {
      setBusy(false);
    }
  }

  function updateCell(rowIndex, columnIndex, value) {
    const targetSetter = exportMode === "standard" ? setStandardSheets : setWebSheets;
    targetSetter((current) =>
      current.map((sheet, index) =>
        index === activeSheet
          ? {
              ...sheet,
              rows: sheet.rows.map((row, r) =>
                r === rowIndex ? row.map((cell, c) => (c === columnIndex ? value : cell)) : row
              )
            }
          : sheet
      )
    );
  }

  function handleBatchChange(value) {
    setDefaultBatch(value);
    setWebSheets((current) =>
      current.map((sheet) => ({
        ...sheet,
        rows: sheet.rows.map((row, r) => {
          if (r === 0) return row;
          const newRow = [...row];
          newRow[3] = value;
          return newRow;
        })
      }))
    );
  }

  // Download active sheet in Standard mode
  async function downloadXlsxStandard() {
    if (rawWorkbooksData.length === 0) return;
    setBusy(true);
    setError("");
    try {
      const activeSheetData = standardSheets[activeSheet];
      const exportWb = new ExcelJS.Workbook();
      const ws = exportWb.addWorksheet(activeSheetData.sheetName || "Cleaned");

      activeSheetData.rows.forEach((row, r) => {
        ws.addRow(row);
        if (r > (activeSheetData.headerIndex ?? 0)) {
          const rollIndex = activeSheetData.rows[activeSheetData.headerIndex ?? 0]?.findIndex(
            (c) => normalizeHeader(c) === "roll"
          );
          if (rollIndex >= 0) {
            ws.getCell(r + 1, rollIndex + 1).numFmt = "@";
          }
        }
      });

      const colCount = Math.max(0, ...activeSheetData.rows.map((r) => r.length));
      for (let c = 0; c < colCount; c += 1) {
        const longest = Math.max(0, ...activeSheetData.rows.map((row) => String(row[c] ?? "").length));
        ws.getColumn(c + 1).width = Math.min(48, Math.max(12, longest + 2));
      }
      ws.eachRow((row) => {
        row.alignment = { ...row.alignment, vertical: "middle" };
      });

      const buffer = await exportWb.xlsx.writeBuffer();
      downloadBlob(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }),
        `${safeBaseName(fileName)}.xlsx`
      );
    } catch (cause) {
      setError(cause?.message || "Could not create the Excel file.");
    } finally {
      setBusy(false);
    }
  }

  function downloadPdf() {
    setBusy(true);
    setError("");
    try {
      const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
      sheets.forEach((sheet, index) => {
        if (index) doc.addPage("a4", "landscape");
        doc.setFont("helvetica", "bold");
        doc.setFontSize(18);
        doc.setTextColor(28, 45, 39);
        doc.text(sheet.name, 40, 38);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(9);
        doc.setTextColor(100, 116, 108);
        doc.text(`Cleaned export • ${new Date().toLocaleDateString()}`, 40, 54);
        const header = sheet.headerIndex ?? 0;
        autoTable(doc, {
          startY: 68,
          head: [sheet.rows[header]],
          body: sheet.rows.slice(header + 1),
          theme: "plain",
          margin: { left: 40, right: 40 },
          styles: {
            font: "helvetica",
            fontSize: 8,
            cellPadding: 6,
            lineColor: [220, 228, 223],
            lineWidth: 0.5,
            textColor: [38, 52, 47],
            overflow: "linebreak"
          },
          headStyles: { fillColor: [28, 45, 39], textColor: 255, fontStyle: "bold" },
          alternateRowStyles: { fillColor: [244, 247, 245] },
          didDrawPage: (data) => {
            doc.setFontSize(8);
            doc.setTextColor(120);
            doc.text(
              `Page ${doc.getNumberOfPages()}`,
              data.settings.margin.left,
              doc.internal.pageSize.height - 18
            );
          }
        });
      });
      doc.save(`${safeBaseName(fileName)}.pdf`);
    } catch (cause) {
      setError(cause?.message || "Could not create the PDF.");
    } finally {
      setBusy(false);
    }
  }

  // Download active Web Sheet as XLSX
  async function downloadXlsxWebPortal() {
    setBusy(true);
    setError("");
    try {
      const currentSheet = webSheets[activeSheet];
      if (!currentSheet) return;
      const exportWorkbook = new ExcelJS.Workbook();
      const worksheet = exportWorkbook.addWorksheet(currentSheet.sheetName || "WebPortal");

      currentSheet.rows.forEach((row, r) => {
        worksheet.addRow(row);
        if (r > 0) {
          worksheet.getCell(r + 1, 1).numFmt = "@"; // roll_number
          worksheet.getCell(r + 1, 3).numFmt = "@"; // mobile
        }
      });

      for (let c = 0; c < 4; c += 1) {
        const longest = Math.max(0, ...currentSheet.rows.map((row) => String(row[c] ?? "").length));
        worksheet.getColumn(c + 1).width = Math.min(48, Math.max(14, longest + 2));
      }
      worksheet.eachRow((row) => {
        row.alignment = { ...row.alignment, vertical: "middle" };
      });

      const buffer = await exportWorkbook.xlsx.writeBuffer();
      const outName = currentSheet.isCombined
        ? `${safeBaseName(fileName)}.xlsx`
        : `${safeBaseName(fileName)}_${safeBaseName(currentSheet.sheetName || currentSheet.name)}.xlsx`;

      downloadBlob(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }),
        outName
      );
    } catch (cause) {
      setError(cause?.message || "Could not create the Excel file.");
    } finally {
      setBusy(false);
    }
  }

  // Download CSV of active sheet
  function downloadCsv(sheet) {
    if (!sheet) return;
    setBusy(true);
    setError("");
    try {
      const csvContent = sheet.rows
        .map((row) =>
          row
            .map((cell) => {
              const str = String(cell ?? "");
              if (str.includes(",") || str.includes('"') || str.includes("\n") || str.includes("\r")) {
                return `"${str.replace(/"/g, '""')}"`;
              }
              return str;
            })
            .join(",")
        )
        .join("\r\n");

      const outName = sheet.isCombined
        ? `${safeBaseName(fileName)}.csv`
        : `${safeBaseName(fileName)}_${safeBaseName(sheet.sheetName || sheet.name)}.csv`;

      downloadBlob(
        new Blob([csvContent], { type: "text/csv;charset=utf-8;" }),
        outName
      );
    } catch (cause) {
      setError(cause?.message || "Could not create the CSV file.");
    } finally {
      setBusy(false);
    }
  }

  // Combined Master Excel of ALL uploaded sheets
  async function downloadCombinedMasterXlsx() {
    if (webSheets.length === 0) return;
    setBusy(true);
    setError("");
    try {
      const exportWorkbook = new ExcelJS.Workbook();
      const worksheet = exportWorkbook.addWorksheet("Master_Web_Portal");

      // Add Header once
      worksheet.addRow(["roll_number", "name", "mobile", "batch"]);

      let rowIndex = 2;
      webSheets.forEach((sheetData) => {
        sheetData.rows.slice(1).forEach((row) => {
          worksheet.addRow(row);
          worksheet.getCell(rowIndex, 1).numFmt = "@";
          worksheet.getCell(rowIndex, 3).numFmt = "@";
          rowIndex += 1;
        });
      });

      for (let c = 0; c < 4; c += 1) {
        worksheet.getColumn(c + 1).width = 22;
      }
      worksheet.eachRow((row) => {
        row.alignment = { ...row.alignment, vertical: "middle" };
      });

      const buffer = await exportWorkbook.xlsx.writeBuffer();
      downloadBlob(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }),
        `${safeBaseName(fileName)}_MASTER_ALL.xlsx`
      );
    } catch (cause) {
      setError(cause?.message || "Could not create combined master Excel.");
    } finally {
      setBusy(false);
    }
  }

  // Download all cleaned sheets as ZIP
  async function downloadAllZip() {
    if (sheets.length === 0) return;
    setBusy(true);
    setError("");
    try {
      const zip = new JSZip();

      for (let i = 0; i < sheets.length; i++) {
        const sheetData = sheets[i];
        const exportWorkbook = new ExcelJS.Workbook();
        const ws = exportWorkbook.addWorksheet(sheetData.sheetName || `Sheet_${i + 1}`);

        sheetData.rows.forEach((row, r) => {
          ws.addRow(row);
          if (exportMode === "web_portal" && r > 0) {
            ws.getCell(r + 1, 1).numFmt = "@";
            ws.getCell(r + 1, 3).numFmt = "@";
          }
        });

        for (let c = 0; c < (sheetData.rows[0]?.length || 4); c += 1) {
          ws.getColumn(c + 1).width = 20;
        }

        const buffer = await exportWorkbook.xlsx.writeBuffer();
        const baseName = safeBaseName(sheetData.name);
        zip.file(`${baseName}.xlsx`, buffer);
      }

      const zipBlob = await zip.generateAsync({ type: "blob" });
      downloadBlob(zipBlob, `${safeBaseName(fileName)}_bundle.zip`);
    } catch (cause) {
      setError(cause?.message || "Could not create the ZIP file.");
    } finally {
      setBusy(false);
    }
  }

  const current = sheets[activeSheet];

  return (
    <main>
      <header className="topbar">
        <a className="brand" href="#">
          <span>
            <Sparkles size={18} />
          </span>
          Student File Cleaner
        </a>
        <div className="topbar-actions">
          <button
            className={`sheet-status-pill ${googleSheetMap.size > 0 ? "connected" : "idle"}`}
            onClick={() => setSheetsModalOpen(true)}
            title="Manage connected Google Sheets"
          >
            <span className="dot"></span>
            <span className="pill-text">
              {syncingSheets
                ? "Syncing Google Sheets…"
                : googleSheetMap.size > 0
                ? `${googleSheetMap.size} Web Rolls linked`
                : "Connect Google Sheets"}
            </span>
            <ExternalLink size={13} />
          </button>
          <div className="privacy">
            <LockKeyhole size={15} /> Files stay on your device
          </div>
        </div>
      </header>

      <section className="hero">
        <div className="eyebrow">GOOGLE SHEETS LINKED • PRIVATE • FAST • NO SIGN-IN</div>
        <h1>
          Clean student files.
          <br />
          <em>Keep the important parts.</em>
        </h1>
        <p>
          Auto-matches roll numbers against Google Sheets, prioritizing secondary phone numbers, and
          prepares flawless web portal exports.
        </p>

        <div className="hero-sheet-bar">
          <div className="sheet-bar-summary">
            <strong>
              <Database size={15} /> {googleSheets.length} Google Sheet{googleSheets.length === 1 ? "" : "s"} Active
            </strong>
            <span>• {googleSheetMap.size} Web Rolls indexed</span>
            {lastSyncedTime && <small className="sync-time">Synced at {lastSyncedTime}</small>}
          </div>
          <div className="sheet-bar-btns">
            <button
              className="sync-btn"
              onClick={() => syncGoogleSheets(googleSheets)}
              disabled={syncingSheets}
              title="Re-fetch latest rows from Google Sheets"
            >
              <RefreshCw size={13} className={syncingSheets ? "spin" : ""} />
              {syncingSheets ? "Syncing…" : "Sync Sheets"}
            </button>
            <button className="manage-sheets-btn" onClick={() => setSheetsModalOpen(true)}>
              Manage Sheets
            </button>
          </div>
        </div>
      </section>

      {rawWorkbooksData.length === 0 ? (
        <section
          className={`dropzone ${dragging ? "dragging" : ""}`}
          onDragOver={(e) => {
            e.preventDefault();
            setDragging(true);
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragging(false);
            loadFiles(e.dataTransfer.files);
          }}
        >
          <input
            ref={inputRef}
            hidden
            type="file"
            accept=".xlsx"
            multiple
            onChange={(e) => loadFiles(e.target.files)}
          />
          <div className="upload-icon">
            <Upload size={28} />
          </div>
          <h2>{busy ? "Reading workbook(s)…" : "Drop your Excel files here"}</h2>
          <p>Select one or multiple .xlsx sheets to clean at once</p>
          <button className="primary" onClick={() => inputRef.current?.click()} disabled={busy}>
            <FileSpreadsheet size={18} /> Select Excel files
          </button>
          <small>Multiple .XLSX files supported • Processed locally in your browser</small>
        </section>
      ) : (
        <>
          <section className="filebar">
            <div className="filemeta">
              <span className="excel">
                <FileSpreadsheet size={21} />
              </span>
              <div>
                <strong>
                  {rawWorkbooksData.length === 1
                    ? rawWorkbooksData[0].fileName
                    : `${rawWorkbooksData.length} Files Uploaded`}
                </strong>
                <small>
                  {sheets.length} total worksheet{sheets.length === 1 ? "" : "s"} ready
                </small>
              </div>
            </div>
            <div className="filebar-controls">
              <input
                ref={addFilesInputRef}
                hidden
                type="file"
                accept=".xlsx"
                multiple
                onChange={(e) => loadFiles(e.target.files, true)}
              />
              <button
                className="secondary small-btn"
                onClick={() => addFilesInputRef.current?.click()}
                disabled={busy}
              >
                <Plus size={14} /> Add more files
              </button>
              <button
                className="ghost"
                onClick={() => {
                  setRawWorkbooksData([]);
                  setStandardSheets([]);
                  setWebSheets([]);
                  setError("");
                }}
              >
                <RotateCcw size={16} /> Start over
              </button>
            </div>
          </section>

          <div className="mode-tabs">
            <button
              className={`tab-btn ${exportMode === "web_portal" ? "active" : ""}`}
              onClick={() => {
                setExportMode("web_portal");
                setEditing(false);
              }}
            >
              Web Portal Formatter (Google Sheet Phone Integration)
            </button>
            <button
              className={`tab-btn ${exportMode === "standard" ? "active" : ""}`}
              onClick={() => {
                setExportMode("standard");
                setEditing(false);
              }}
            >
              Standard Cleaner (Masked)
            </button>
          </div>

          <section className="results">
            <div className="result-title">
              <div>
                <span className="success">
                  <Check size={18} />
                </span>
                <div>
                  <h2>Your {sheets[0]?.isCombined ? "Combined Master file" : "file"} is {exportMode === "standard" ? "clean" : "ready for Web Portal"}</h2>
                  <p>
                    {summary.rows} student records {sheets[0]?.isCombined ? `merged across ${rawWorkbooksData.length} uploaded files` : `across ${sheets.length} sheet`}.
                  </p>
                </div>
              </div>
              <button
                className={`edit-toggle ${editing ? "active" : ""}`}
                onClick={() => setEditing(!editing)}
              >
                <Pencil size={15} />
                {editing ? "Finish editing" : "Edit data"}
              </button>
            </div>

            {exportMode === "web_portal" ? (
              <div className="metrics metrics-web">
                <article>
                  <strong>{summary.secondaryCount}</strong>
                  <span>GSheet Secondary (Col J)</span>
                </article>
                <article>
                  <strong>{summary.primaryCount}</strong>
                  <span>GSheet Primary (Col I)</span>
                </article>
                <article>
                  <strong>{summary.unmatchedCount}</strong>
                  <span>Original / Fallback</span>
                </article>
                <article>
                  <strong>{defaultBatch || "None"}</strong>
                  <span>Active Batch</span>
                </article>
              </div>
            ) : (
              <div className="metrics">
                <article>
                  <strong>{summary.emails}</strong>
                  <span>Emails masked</span>
                </article>
                <article>
                  <strong>{summary.phones}</strong>
                  <span>Phones masked</span>
                </article>
                <article>
                  <strong>{summary.rolls}</strong>
                  <span>Rolls cleaned</span>
                </article>
                <article>
                  <strong>Removed</strong>
                  <span>Paid Amount</span>
                </article>
              </div>
            )}

            {summary.warnings > 0 && (
              <div className="warning">
                {summary.warnings} note{summary.warnings === 1 ? "" : "s"} (e.g. rolls with fallback phone or unpadded formats).
              </div>
            )}
          </section>

          <section className="preview-card">
            <div className="preview-head">
              <div>
                <h2>Preview: {current?.name}</h2>
                <p>
                  {editing
                    ? "Click any cell below to make corrections."
                    : exportMode === "web_portal"
                    ? "Roll numbers normalized to 8 digits • Phone replaced from Google Sheet"
                    : "Review the cleaned data before downloading."}
                </p>
              </div>
              {sheets.length > 1 && (
                <label className="sheet-select">
                  Sheet
                  <select
                    value={activeSheet}
                    onChange={(e) => setActiveSheet(Number(e.target.value))}
                  >
                    {sheets.map((s, i) => (
                      <option value={i} key={i}>
                        {s.name}
                      </option>
                    ))}
                  </select>
                  <ChevronDown size={14} />
                </label>
              )}
            </div>

            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    {current?.rows[current.headerIndex ?? 0]?.map((cell, i) => (
                      <th key={i}>{String(cell)}</th>
                    ))}
                    {exportMode === "web_portal" && <th>Source</th>}
                  </tr>
                </thead>
                <tbody>
                  {current?.rows.slice((current.headerIndex ?? 0) + 1).map((row, r) => {
                    const phoneSrc = current.phoneSources ? current.phoneSources[r + 1] : null;
                    return (
                      <tr key={r}>
                        {row.map((cell, c) => (
                          <td key={c}>
                            {editing ? (
                              <input
                                aria-label={`Row ${r + 2}, Col ${c + 1}`}
                                value={String(cell ?? "")}
                                onChange={(e) =>
                                  updateCell(
                                    r + (current.headerIndex ?? 0) + 1,
                                    c,
                                    e.target.value
                                  )
                                }
                              />
                            ) : (
                              String(cell ?? "")
                            )}
                          </td>
                        ))}
                        {exportMode === "web_portal" && (
                          <td>
                            {phoneSrc === "gsheet_secondary" && (
                              <span className="badge badge-sec" title="From Google Sheet Secondary Phone (Col J)">
                                Col J (Secondary)
                              </span>
                            )}
                            {phoneSrc === "gsheet_primary" && (
                              <span className="badge badge-pri" title="From Google Sheet Primary Phone (Col I)">
                                Col I (Primary)
                              </span>
                            )}
                            {(phoneSrc === "uploaded_fallback" || phoneSrc === "uploaded_sheet") && (
                              <span className="badge badge-orig" title="From Uploaded Sheet (Not in Google Sheet)">
                                Uploaded Sheet
                              </span>
                            )}
                          </td>
                        )}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="table-foot">
              <span>
                Showing {Math.max(0, (current?.rows.length ?? 1) - (current?.headerIndex ?? 0) - 1)} rows in this sheet
              </span>
              <span>{current?.name}</span>
            </div>
          </section>

          <section className="downloads">
            <div className="name-field">
              <label htmlFor="filename">Output file name</label>
              <div>
                <input
                  id="filename"
                  value={fileName}
                  onChange={(e) => setFileName(e.target.value)}
                />
                <span>{exportMode === "standard" ? ".xlsx / .pdf" : ".xlsx / .csv"}</span>
              </div>
              <small>Applied to downloaded files.</small>
            </div>

            {exportMode === "web_portal" && (
              <div className="batch-field">
                <label htmlFor="batchname">Default Batch</label>
                <input
                  id="batchname"
                  value={defaultBatch}
                  onChange={(e) => handleBatchChange(e.target.value)}
                  placeholder="e.g. Morning"
                />
                <small>Assigned to the 'batch' column.</small>
              </div>
            )}

            <div className="download-actions">
              {exportMode === "standard" ? (
                <>
                  <button className="primary dark" onClick={downloadXlsxStandard} disabled={busy}>
                    <Download size={18} /> {current?.isCombined ? `Download Combined Clean XLSX (${current.rows.length - 1} records)` : "Download Clean XLSX"}
                  </button>
                  <button className="secondary" onClick={downloadPdf} disabled={busy}>
                    <FileText size={18} /> Download PDF
                  </button>
                  {sheets.length > 1 && (
                    <button className="secondary" onClick={downloadAllZip} disabled={busy}>
                      <FolderArchive size={18} /> All Files (.ZIP)
                    </button>
                  )}
                </>
              ) : (
                <>
                  <button className="primary dark" onClick={downloadXlsxWebPortal} disabled={busy}>
                    <Download size={18} /> {current?.isCombined ? `Download Combined Web XLSX (${current.rows.length - 1} records)` : "Download Web XLSX"}
                  </button>
                  <button className="secondary" onClick={() => downloadCsv(current)} disabled={busy}>
                    <FileText size={18} /> {current?.isCombined ? "Download Combined CSV" : "Download CSV"}
                  </button>
                  {sheets.length > 1 && (
                    <button className="secondary" onClick={downloadAllZip} disabled={busy}>
                      <Archive size={18} /> All Files (.ZIP)
                    </button>
                  )}
                </>
              )}
            </div>
          </section>
        </>
      )}

      {/* Google Sheets Management Modal */}
      {sheetsModalOpen && (
        <div className="modal-overlay" onClick={() => setSheetsModalOpen(false)}>
          <div className="modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div>
                <h2>Manage Google Sheets</h2>
                <p>Add and synchronize multiple Google Sheets for roll and phone lookup.</p>
              </div>
              <button className="close-btn" onClick={() => setSheetsModalOpen(false)}>
                <X size={20} />
              </button>
            </div>

            <div className="db-status-bar">
              <Database size={15} />
              <span>
                Storage mode: <strong>{storageType === "postgres" ? "Vercel Postgres" : storageType === "kv" ? "Vercel KV" : "Browser Storage + Vercel API"}</strong>
              </span>
              <button
                className="sync-modal-btn"
                onClick={() => syncGoogleSheets(googleSheets)}
                disabled={syncingSheets}
              >
                <RefreshCw size={13} className={syncingSheets ? "spin" : ""} />
                {syncingSheets ? "Syncing…" : "Sync All Now"}
              </button>
            </div>

            <form className="add-sheet-form" onSubmit={handleAddSheet}>
              <div className="form-row">
                <input
                  type="text"
                  placeholder="Sheet Label / Batch Name (e.g. Batch 28 Sheet)"
                  value={newSheetTitle}
                  onChange={(e) => setNewSheetTitle(e.target.value)}
                />
              </div>
              <div className="form-row flex-row">
                <input
                  type="url"
                  required
                  placeholder="Paste Google Sheet URL (https://docs.google.com/spreadsheets/d/...)"
                  value={newSheetUrl}
                  onChange={(e) => setNewSheetUrl(e.target.value)}
                />
                <button type="submit" className="primary small-submit">
                  <Plus size={16} /> Add Sheet
                </button>
              </div>
              <small className="form-hint">
                Make sure the Google Sheet sharing setting is set to "Anyone with the link can view".
              </small>
            </form>

            <div className="sheet-list">
              <h3>Connected Sheets ({googleSheets.length})</h3>
              {googleSheets.length === 0 ? (
                <div className="empty-sheets">No Google Sheets connected yet.</div>
              ) : (
                googleSheets.map((s) => (
                  <div className="sheet-item" key={s.id || s.url}>
                    <div className="sheet-info">
                      <div className="sheet-title-line">
                        <strong>{s.title || "Untitled Sheet"}</strong>
                        {s.isDefault && <span className="default-tag">Default</span>}
                        {s.status === "connected" && (
                          <span className="count-tag">{s.recordCount || 0} Web Rolls</span>
                        )}
                        {s.status === "error" && (
                          <span className="error-tag">Sync Failed</span>
                        )}
                      </div>
                      <a href={s.url} target="_blank" rel="noreferrer" className="sheet-url-link">
                        <Link size={12} /> {s.url}
                      </a>
                      {s.lastSynced && (
                        <span className="sync-note">Last synced: {s.lastSynced}</span>
                      )}
                    </div>
                    <div className="sheet-actions">
                      <button
                        className="delete-sheet-btn"
                        onClick={() => handleRemoveSheet(s.id)}
                        title="Remove Google Sheet"
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>

            <div className="modal-footer">
              <button className="primary" onClick={() => setSheetsModalOpen(false)}>
                Done
              </button>
            </div>
          </div>
        </div>
      )}

      {error && (
        <div className="error" role="alert">
          {error}
          <button className="close-err" onClick={() => setError("")}>
            ×
          </button>
        </div>
      )}

      <footer>
        Roll Cleaner & Web Portal Integrator • Google Sheets Secondary-Phone Engine • Nothing leaves your browser.
      </footer>
    </main>
  );
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
