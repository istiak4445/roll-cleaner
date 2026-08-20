# Student File Cleaner

A privacy-first browser app for cleaning weekly student Excel files. It masks email addresses and phone numbers, standardizes numeric roll IDs to eight-character text, removes the entire `Paid Amount` column, allows manual review/editing, and downloads clean XLSX or PDF files.

## Run locally

```bash
npm install
npm run dev
```

## Deploy to Vercel

Import this folder as a Vercel project. The default Vite build settings are detected automatically (`npm run build`, output directory `dist`). No environment variables or backend services are required.

All workbook processing happens locally in the browser; files are never uploaded.
