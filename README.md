# Trial Balance Consolidation Tool

A browser-based tool that consolidates trial balances exported from Thailand's GFMIS for many disbursement units (หน่วยเบิกจ่าย) into a single table, joined by account code. It compares balances between two years and produces an Excel workbook in which every summary cell is a formula that can be traced back to the raw data.

Built for financial-statement audits of government agencies with dozens to hundreds of disbursement units.

**Developed by MRTEAM** · Version 1.0

---

## Features

- **Flexible import** — one file per unit, or one file with many sheets (one sheet per unit). Supports `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.csv`.
- **Automatic layout detection** — finds the header row and the account-code column whether or not the GFMIS report header is present. Non-trial-balance sheets (e.g. supporting schedules) are skipped.
- **Column role detection** — account code, account name, opening balance, debit, credit and closing balance (net or split debit/credit). Roles can be edited; sheets with identical headers are grouped so one edit applies to all of them.
- **GFMIS negative credits** — detected and normalised automatically.
- **Unit matching by 10-digit disbursement-unit code** — read from the “รหัสหน่วยเบิกจ่าย :” line of the report header, independent of file names.
- **Fiscal year detection** — year labels are read from the report period.
- **Three comparison modes**
  1. *Single file with both years* (default) — opening balance = prior year, closing balance = current year.
  2. *Single file + prior-year trial balance* — selected account groups (default: 3, 4, 5) take the closing balance from the prior-year file, so equity, revenue and expense accounts are compared on the same basis.
  3. *Separate files per year.*
- **Display options** — balance type per year, account-code filter, aggregation by the first 1/2/4/6/8 digits, unit and account selection, hide zero rows, show credit-side accounts as positive.
- **Table orientation** — accounts as rows and units as columns, or the reverse.
- **Pre-download summary** — shows which column of which file feeds each sheet.

## Excel output

### Formula workbook (recommended)

| Sheet | Contents |
|---|---|
| คำอธิบาย (Instructions) | How to use the file, tab colours, how to paste new data |
| ตั้งค่า (Settings) | Highlight thresholds (percent / amount), expected report period, list of raw-data sheets |
| Year sheets × 2 | Consolidated table per year. Every cell is a `SUMIFS` formula on the raw sheets by account code, with row and column totals |
| เปรียบเทียบ (Comparison) | Change amount and percentage for each unit/account pair, totals on both axes, yellow highlight when thresholds are met |
| ตรวจสอบ (Checks) | Per-sheet validation (see below) |
| สรุปการตั้งค่า (Settings summary) | Columns, filters and options used — suitable for audit working papers |
| Raw sheets (named by unit code) | Original data, values only |

Tab colours: 🟩 summary sheets · 🟦 current-year raw data · 🟧 prior-year raw data

Formulas search for account codes across rows 1–2,000, so **new data can be pasted over a raw sheet** — with or without the report header, starting at any row (clear the old data first and keep the column positions).

### Values-only workbook

A smaller file with year sheets, comparison, long-format comparison (for filters/pivot tables), settings summary and checks. Always uses accounts-as-rows orientation.

## Automatic checks

The Checks sheet marks each raw sheet **ผ่าน (pass)** or **ตรวจสอบ (review)** by testing that:

- the report period matches the expected period (flags “ไม่พบหัวรายงาน” when the header is missing);
- total opening and total closing balances are zero;
- total debits equal total credits;
- for every account, opening + debit − credit = closing;
- no account in the raw sheet is missing from the summary tables;
- current-year opening equals prior-year closing for accounts not taken from the prior-year file (when a prior-year file is loaded).

## Usage

1. Open the web page and drop the trial-balance files onto **Year 1** (and prior-year files onto the orange box, if any).
2. Wait for the **“Ready”** panel and review any warnings.
3. Click **Download formula workbook** — or **Customise first** to review column roles (step 3) and display options (step 4).
4. Open the file in Excel (click *Enable Editing* if prompted) and review the Checks sheet before using the figures.

A **sample data** button lets you try the tool without real files.

## Data privacy

All files are read and processed entirely in the user's browser. **No data is sent to any server.** Data lives only in the page's memory and is discarded when the page is closed or reloaded.

## Requirements

- A current web browser (Chrome, Edge, Safari, Firefox)
- Internet access on first load to fetch the spreadsheet-reading library
- Microsoft Excel 2010 or later to open the formula workbook (tested with Excel 365)

## Known limitations

- The formula workbook uses the account list that existed when it was generated. If newly pasted data contains new accounts, the Checks sheet reports them under “บัญชีที่ไม่อยู่ในชีตสรุป”; regenerating the workbook is recommended.
- Year labels in sheet titles are static text and do not update when data is pasted over.
- Revenue and expense accounts (groups 4–5) have zero opening balances; without a prior-year file they appear in the comparison as a full-amount increase.
- Credit-side balances are shown as negative numbers, as in GFMIS, unless the positive-credit option is selected.

## Technology

- Single-page application (HTML, CSS, JavaScript); no backend
- Reading files: [SheetJS](https://sheetjs.com/) (`xlsx` 0.18.5)
- Values-only workbook: [ExcelJS](https://github.com/exceljs/exceljs) 4.4.0
- Formula workbook: a built-in OOXML writer (worksheet XML plus ZIP via `CompressionStream`) giving full control over formulas, array formulas, conditional formatting, defined names, freeze panes and tab colours
- Font: IBM Plex Sans Thai (Google Fonts)

## Testing

Tested with real GFMIS trial balances for 130 disbursement units and 317 accounts (20 files). Generated workbooks were recalculated in LibreOffice and compared with the source figures: every cell matched in all modes, both orientations and several paste-over scenarios. Results were also confirmed in Excel 365.

## License

Released under the MIT License — see [LICENSE](LICENSE).

© 2026 MRTEAM

This software is a data-preparation aid. Users should verify the results against the source data before relying on them in audit work.
