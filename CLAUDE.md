# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Summary

VBA-based automation for transforming daily NAV reports into portfolio reports. Two components:
1. **Outlook VBA** - Monitors inbox, saves Custom attachment, triggers transformation
2. **Excel VBA** - Transforms data into Dashboard + Stocks + Options + Currencies + Other tabs

## Architecture (v5.7.1 - Simplified)

```
Outlook (monitors inbox)
    ↓ detects Custom portfolio report email by subject pattern
    ↓ saves attachment to C:\Mobius Reports\Incoming\
    ↓ triggers transformation immediately (no waiting for second email)
    ↓
Excel (transforms data)
    ↓ reads Custom file (positions + YTD return from Column H)
    ↓ outputs to C:\Mobius Reports\Transformed\
```

**Note:** As of v5.7.0, only the Custom email is required. The Daily Reports email with 15 attachments is no longer needed since YTD Return is now read from the Custom file.

## Data Sources

| Email | Subject Pattern | Key File | Used For |
|-------|-----------------|----------|----------|
| **Custom** (REQUIRED) | `...Custom daily portfolio report MMDDYYYY` | `Gain And Exposure_Custom_...XLSX` | All position data + YTD Return |
| Daily Reports (optional) | `...Daily Reports MMDDYYYY` | 15 files incl. NAV Portfolio Notebook | Future enhancements |

## Custom File Column Mapping

| Col | Header | Used | Purpose |
|-----|--------|------|---------|
| A | Product Name | Yes | Security name, classify stock/option |
| B | Ticker | Yes | Bloomberg ticker or OCC option symbol |
| C | ISIN | No | International identifier |
| D | Portfolio Weight % | Yes | Position size (last row = 1.0 for total) |
| E | Unit Cost (USD) | Yes | Cost basis per share |
| **F** | **Today (USD)** | **Yes** | **Current market price** |
| G | % Daily Gain/Loss | No | Daily change % |
| **H** | **Jan 1 ROR** | **Yes** | **YTD Return (last row only)** |
| I | Total Cost (USD) | No | Total cost basis |
| J | Market Value (USD) | Yes | Current position value |
| K | Total P&L YTD | Yes | P&L for position |
| L | # of Shares | Yes | Quantity held |

**YTD Return Location:** Last row, Column H. Validated by checking Column D = 1.0 (100% weight).

## Key Code Locations

### OutlookMonitor.txt
- `Application_Startup` - Initializes monitor on Outlook launch
- `InboxItems_ItemAdd` - Event handler for new emails
- `ProcessIncomingEmail` - Checks subject, extracts date, saves attachments
- `StripForwardPrefixes` - Removes FW:/RE: prefixes (handles multiples)
- `TriggerTransformation` - Launches Excel immediately (no waiting for second email)

### BloombergDataTransformer.vba
- `TransformBloombergData` - Main entry point
- `ReadYTDFundReturn(ws)` - Reads YTD Return from Custom file Column H (last row)
- `ProcessStock` / `ProcessOption` - Maps source columns to output
- `IsCurrency` / `IsAdminExpense` - Routes positions to Currencies/Other tabs
- `CreateDashboard` - Creates summary dashboard with KPIs and charts

## Configuration Constants

In `OutlookMonitor.txt`:
```vba
Private Const BASE_FOLDER As String = "C:\Mobius Reports"
Private Const INCOMING_FOLDER As String = "C:\Mobius Reports\Incoming"
Private Const TRANSFORMED_FOLDER As String = "C:\Mobius Reports\Transformed"
Private Const EXCEL_TRANSFORMER_PATH As String = "C:\Mobius Reports\Portfolio Transformer.xlsm"
Private Const SUBJECT_CUSTOM As String = "Mobius Emerging Opportunities Fund LP| Custom daily portfolio report"
```

## Data Flow (No External Dependencies)

As of v5.3, all data comes from the NAV reports - no Bloomberg Terminal required:
- Stock prices: Read from "Today USD" column (Column F) in Custom file
- Option underlying prices: Looked up from stock positions in same report
- YTD Return: Read from Column H of Custom file's last row
- FX conversion: Applied for non-USD tickers (JP, LN, GY suffixes)

OCC format for options: `META  260116P00700000` (TICKER  YYMMDD[P/C]STRIKE)

## Output File Structure

Output filename: `Portfolio_MMDDYYYY_gen_MMDDYYYY.xlsx`
- First date = Report date (from source file)
- Second date = Processing date (when macro ran)

Tabs created:
1. **Dashboard** - KPIs, charts (Top Holdings, Allocation, P&L, Performance)
2. **Stocks** - Stock positions with prices, P&L, attribution
3. **Options** - Option positions with details
4. **Currencies** - Cash positions (USD, CAD, JPY, etc.)
5. **Other** - Admin items (fees, payables, accruals)

## Testing

- Forward the Custom email to yourself (handles FW: prefixes)
- Use `ProcessSelectedEmail` macro to manually process
- Use `ShowTrackerState` to verify email tracking
- Use `RunManualTest` to check monitor status

## Distribution Procedure (CRITICAL)

When creating a new version release, ALL of the following steps MUST be completed:

### 1. Version String Updates (ALL files)

Update version numbers in these locations:

**Root files:**
- `BloombergDataTransformer.vba` line 2: `' Portfolio Data Transformer - Version X.Y.Z`
- `BloombergDataTransformer.vba` MsgBox: `"Portfolio Data Transformer vX.Y.Z"`
- `BloombergDataTransformer.vba` Dashboard title: `"PORTFOLIO DASHBOARD (vX.Y.Z)"`
- `CLAUDE.md` Architecture header: `## Architecture (vX.Y.Z - Simplified)`
- `CHANGELOG.md` - Add new version entry at top

**Dist files (in `dist/Mobius_Portfolio_Reporter_vX.Y.Z/files/`):**
- `TransformerMacro.txt` - Same 3 locations as BloombergDataTransformer.vba
- `OutlookMonitor.txt` line 2: `' Mobius Portfolio Report - Outlook Email Monitor (vX.Y.Z - Simplified)`
- `OutlookMonitor.txt` RunManualTest msg: `"=== Mobius Report Monitor Test (vX.Y.Z) ==="`

### 2. Sync VBA Code

Copy root `BloombergDataTransformer.vba` to `dist/.../files/TransformerMacro.txt`
- Do NOT just copy-paste the old dist version
- The root file is the source of truth

### 3. Create/Update Documentation

**Source files in `dist/src/`:**
- Create `RELEASE_NOTES_vX.Y.Z.tex` (copy from previous, update content)
- Update `SETUP_GUIDE.tex` version in header and title page

**Compile PDFs:**
```bash
cd dist/src
pdflatex -interaction=nonstopmode RELEASE_NOTES_vX.Y.Z.tex
pdflatex -interaction=nonstopmode RELEASE_NOTES_vX.Y.Z.tex  # Run twice for refs
pdflatex -interaction=nonstopmode SETUP_GUIDE.tex
pdflatex -interaction=nonstopmode SETUP_GUIDE.tex  # Run twice for refs
```

**Copy to dist folder:**
```bash
cp RELEASE_NOTES_vX.Y.Z.pdf SETUP_GUIDE.pdf ../Mobius_Portfolio_Reporter_vX.Y.Z/docs/
```

### 4. Create Dist Folder Structure

```
dist/Mobius_Portfolio_Reporter_vX.Y.Z/
├── README.txt              # Update version and "What's New"
├── docs/
│   ├── RELEASE_NOTES_vX.Y.Z.pdf
│   └── SETUP_GUIDE.pdf
└── files/
    ├── OutlookMonitor.txt      # From previous version, update version strings
    ├── TransformerMacro.txt    # Copy from root BloombergDataTransformer.vba
    └── Portfolio Transformer.xlsm  # Excel workbook (update macro inside)
```

### 5. Create Zip File

```bash
cd dist
zip -r Mobius_Portfolio_Reporter_vX.Y.Z.zip Mobius_Portfolio_Reporter_vX.Y.Z/
```

### 6. Verification Checklist

Before finalizing, verify:
- [ ] `grep -r "X.Y.Z" dist/Mobius_Portfolio_Reporter_vX.Y.Z/` shows correct version everywhere
- [ ] `grep -r "X.Y.Z-1" dist/Mobius_Portfolio_Reporter_vX.Y.Z/` shows NO old version refs
- [ ] PDF filenames match version (RELEASE_NOTES_vX.Y.Z.pdf)
- [ ] SETUP_GUIDE.pdf header shows correct version
- [ ] README.txt has correct version and changelog

### Note on Portfolio Transformer.xlsm

The .xlsm file contains embedded VBA code. To fully update it:
1. Open in Excel
2. Press Alt+F11 to open VBA Editor
3. Replace module code with contents of TransformerMacro.txt
4. Save and close

## Untapped Data (Future Enhancements)

The Daily Reports email contains a NAV Portfolio Notebook (172KB, 13 sheets) with:
- Attribution by Sector/Country/Product Type
- Top 10 Movers (gainers/decliners)
- Realized vs Unrealized P&L breakdown
- Tax Lot details with purchase dates
- Dividend tracking

This data could enhance the dashboard in future versions.
