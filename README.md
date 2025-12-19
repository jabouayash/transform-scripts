# Mobius Portfolio Report Automation

Automated system for transforming daily NAV (Net Asset Value) reports into structured portfolio reports with Bloomberg API integration.

## Overview

This system monitors Outlook for incoming daily report emails, extracts attachments, and automatically transforms them into formatted portfolio reports.

```
┌─────────────────────────────────────────────────────────────┐
│                    OUTLOOK                                   │
│  Monitors inbox for two specific email subjects             │
│  Saves attachments when both emails arrive                  │
└─────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────┐
│                    EXCEL                                     │
│  Transforms NAV data into Stocks + Options tabs             │
│  Adds Bloomberg live price formulas                         │
│  Outputs formatted report                                   │
└─────────────────────────────────────────────────────────────┘
```

## Input Requirements

### Email 1: Custom Daily Portfolio Report
- **Subject:** `Mobius Emerging Opportunities Fund LP| Custom daily portfolio report MMDDYYYY`
- **Required attachment:** `Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP_MMDDYYYY.XLSX`
- **Contains:** Individual position data (stocks, options, quantities, P&L, weights)

### Email 2: Daily Reports
- **Subject:** `Mobius Emerging Opportunities Fund LP| Daily Reports MMDDYYYY`
- **Required attachment:** `Gain And Exposure_MOBIUS EMERGING OPPORTUNITIES FUND LP_MMDDYYYY.XLSX`
- **Contains:** Fund-level performance (YTD Fund Return in cell K94)

### Input File Structure (Custom File)

| Row | Content |
|-----|---------|
| 1-3 | Empty/formatting (ignored) |
| 4 | Sub-headers |
| 5 | Column headers |
| 6+ | Position data |

| Column | Header | Description |
|--------|--------|-------------|
| A | Product Name | Security name (used to identify stocks vs options) |
| B | Ticker | Bloomberg ticker or OCC format for options |
| C | ISIN | Security identifier |
| D | Portfolio Weight % | Position weight |
| E | Unit Cost USD | Average cost basis |
| F | Today USD | Current price |
| G | % Daily Gain/Loss | Daily change |
| H | Contribution to Performance | Attribution |
| I | Total Cost USD | Total cost basis |
| J | Market Value USD | Current market value |
| K | Total Net P&L YTD | Year-to-date P&L |
| L | # of Shares | Position quantity (negative = short) |

## Output

### Location
`C:\Mobius Reports\Transformed\Transformed_Portfolio_DD MMMM YYYY.xlsx`

### Stocks Tab
| Column | Header | Description |
|--------|--------|-------------|
| A | Name | Security name |
| B | Ticker | Bloomberg ticker |
| C | Portfolio Wgt | Position weight % |
| D | % Diff (Cost) | Gain/loss vs cost basis |
| E | Daily Chg % | Today's price change |
| F | Unit Cost | Average cost (USD, rounded to nearest dollar) |
| G | Current Px | Live price via Bloomberg BDP() (USD, rounded) |
| H | Total Cost | Cost basis (USD, rounded) |
| I | Mkt Value | Current value (USD, rounded) |
| J | P&L | Year-to-date P&L (USD) |
| K | Attribution | Performance contribution % |

Note: All currency values display as numbers only (no $ symbols). Headers indicate USD.

### Options Tab
Separated into PUTS and CALLS sections:

| Column | Header | Description |
|--------|--------|-------------|
| A | Name | Option description (e.g., META 01/16/2026 PUT 700) |
| B | Quantity | # of contracts |
| C | Underlying Qty | Shares of underlying stock owned |
| D | % Hedged | Coverage ratio |
| E | Strike Px | Strike price (USD) |
| F | Underlying Px | Live underlying price via Bloomberg (USD) |
| G | % Moneyness | ITM/OTM percentage |
| H | Expiry | Expiration date |
| I | Unit Cost | Average cost per contract (USD) |
| J | Total Cost | Total cost basis (USD) |
| K | Current Px | Current option price (USD) |
| L | Mkt Value | Current market value (USD) |
| M | P&L ($) | Year-to-date P&L (USD) |

Note: All currency values display as numbers only (no $ symbols). Headers indicate USD.

### Fund Performance Summary
Added at bottom of Stocks tab:
- Total Portfolio Value
- NAV Per Share
- Fund Inception Date
- YTD Fund Return (from K94 of non-custom file)
- MTD Net Return

### Output Formatting

The transformed output matches the professional styling of the input files:

| Element | Style |
|---------|-------|
| **Sub-headers** | Navy blue background (`#003366`) with white text |
| **Column headers** | Bold, gray text (`#595959`) |
| **Data rows** | Alternating white (`#FFFFFF`) and light gray (`#F2F2F2`) zebra striping |
| **Data text** | Dark gray (`#404040`) |
| **Borders** | Light gray thin borders |

## Prerequisites

- Windows 10/11
- Microsoft Outlook (Classic/Desktop version, NOT "New Outlook")
- Microsoft Excel with macros enabled
- Bloomberg Terminal (for live price formulas)

**Important:** The "New Outlook" does not support VBA macros. You must use Classic Outlook.

## Installation

### 1. Create Folder Structure
```
C:\Mobius Reports\
├── Incoming\        ← Attachments saved here
├── Transformed\     ← Output reports saved here
└── Archive\         ← For manual archival
```

### 2. Set Up Excel

1. Open Excel
2. Press `Alt + F11` (or Developer → Visual Basic)
3. Insert → Module
4. Paste contents of `BloombergDataTransformer.vba`
5. Save as `C:\Mobius Reports\Portfolio Transformer.xlsm` (Macro-Enabled Workbook)

### 3. Set Up Outlook

1. Open Outlook (Classic version)
2. Enable Developer tab: File → Options → Customize Ribbon → check Developer
3. Click Developer → Visual Basic
4. Double-click `ThisOutlookSession` in Project Explorer
5. Paste contents of `OutlookEmailMonitor.vba`
6. Press `Ctrl + S` to save
7. Enable macros: File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros
8. Restart Outlook

### 4. Verify Setup

1. In Outlook, press `Alt + F8`
2. Select `RunManualTest` → Run
3. You should see a status popup confirming the monitor is active

## Usage

### Automatic (Default)
Once set up, the system runs automatically:
1. Both daily emails arrive in your inbox
2. Outlook detects and saves the attachments
3. Excel transformation runs automatically
4. Output saved to `C:\Mobius Reports\Transformed\`

### Manual Processing (via Outlook)
To process emails manually:
1. Select an email in Outlook
2. Press `Alt + F8` → `ProcessSelectedEmail` → Run

### Manual Processing (without email)
To test or run the transformation directly:
1. Open the Custom NAV file from `C:\Mobius Reports\Incoming\`
2. Make sure `Portfolio Transformer.xlsm` is also open
3. Press `Alt + F8` → `TransformBloombergData` → Run
4. Output appears in `C:\Mobius Reports\Transformed\`

### Testing with Forwarded Emails
The system handles `FW:`, `Fwd:`, and `RE:` prefixes (including multiples like `FW: FW: FW:`), so you can test by forwarding old emails to yourself.

## Available Macros

### Outlook (Alt + F8)
| Macro | Purpose |
|-------|---------|
| `RunManualTest` | Check if monitor is active and configured |
| `ProcessSelectedEmail` | Manually process a selected email |
| `ShowTrackerState` | See which emails have been tracked |
| `CheckFolderContents` | List files in Incoming folder |
| `ClearIncomingFolder` | Delete all files in Incoming folder |
| `ResetTracker` | Clear email tracking state |

### Excel (Alt + F8)
| Macro | Purpose |
|-------|---------|
| `TransformBloombergData` | Run the transformation manually |

## Troubleshooting

### "Macros are disabled"
File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros

### Monitor not starting
Run `InitializeMonitor` macro manually (Alt + F8)

### Emails not detected
1. Check subject line matches expected pattern exactly
2. Run `ShowTrackerState` to see what's being tracked
3. Try `ProcessSelectedEmail` on the email directly

### Bloomberg formulas show #N/A
- Ensure Bloomberg Terminal is running and logged in
- Test manually: type `=BDP("AAPL US Equity","PX_LAST")` in Excel

### Transformation errors
1. Ensure both files are in `C:\Mobius Reports\Incoming\`
2. Ensure `Portfolio Transformer.xlsm` exists in `C:\Mobius Reports\`
3. Check that file names match expected pattern

## Files

| File | Description |
|------|-------------|
| `BloombergDataTransformer.vba` | Excel VBA - transforms NAV data |
| `OutlookEmailMonitor.vba` | Outlook VBA - monitors emails, saves attachments |
| `CLAUDE.md` | Technical documentation for AI assistants |

## Configuration

Default paths are configured in `OutlookEmailMonitor.vba`:
```vba
Private Const BASE_FOLDER As String = "C:\Mobius Reports"
Private Const INCOMING_FOLDER As String = "C:\Mobius Reports\Incoming"
Private Const TRANSFORMED_FOLDER As String = "C:\Mobius Reports\Transformed"
Private Const EXCEL_TRANSFORMER_PATH As String = "C:\Mobius Reports\Portfolio Transformer.xlsm"
```

To change locations, edit these constants and re-import the VBA code.
