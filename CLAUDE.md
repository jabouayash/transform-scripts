# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Summary

VBA-based automation for transforming daily NAV reports into portfolio reports. Two components:
1. **Outlook VBA** - Monitors inbox, saves attachments when both daily emails arrive
2. **Excel VBA** - Transforms data into formatted Stocks + Options tabs with Bloomberg formulas

## Architecture

```
Outlook (monitors inbox)
    ↓ detects emails by subject pattern
    ↓ saves attachments to C:\Mobius Reports\Incoming\
    ↓ when both emails arrive for same date
    ↓
Excel (transforms data)
    ↓ reads Custom file (positions)
    ↓ reads Non-Custom file K94 (YTD return)
    ↓ outputs to C:\Mobius Reports\Transformed\
```

## Data Sources

| Email Subject Pattern | Key Attachment | Used For |
|-----------------------|----------------|----------|
| `...Custom daily portfolio report MMDDYYYY` | `Gain And Exposure_Custom_...MMDDYYYY.XLSX` | Position details |
| `...Daily Reports MMDDYYYY` | `Gain And Exposure_...MMDDYYYY.XLSX` | K94 = YTD Fund Return |

## Key Code Locations

### OutlookEmailMonitor.vba
- `Application_Startup` - Initializes monitor on Outlook launch
- `InboxItems_ItemAdd` - Event handler for new emails
- `ProcessIncomingEmail` - Checks subject, extracts date, saves attachments
- `StripForwardPrefixes` - Removes FW:/RE: prefixes (handles multiples)
- `TriggerTransformation` - Launches Excel when both emails arrive

### BloombergDataTransformer.vba
- `TransformBloombergData` - Main entry point
- `ReadYTDFundReturn` - Reads K94 from non-custom file
- `ProcessStock` / `ProcessOption` - Maps source columns to output
- `AddBottomTotals` - Adds fund performance summary section

## Input File Structure (Custom)

- Rows 1-3: Ignored
- Row 4: Sub-headers
- Row 5: Column headers
- Row 6+: Data

Key columns: A=Product Name, B=Ticker, D=Weight, E=Unit Cost, F=Price, G=Daily Change, H=Attribution, I=Total Cost, J=Market Value, K=P&L, L=# Shares

## Configuration Constants

In `OutlookEmailMonitor.vba`:
```vba
Private Const BASE_FOLDER As String = "C:\Mobius Reports"
Private Const INCOMING_FOLDER As String = "C:\Mobius Reports\Incoming"
Private Const TRANSFORMED_FOLDER As String = "C:\Mobius Reports\Transformed"
Private Const EXCEL_TRANSFORMER_PATH As String = "C:\Mobius Reports\Portfolio Transformer.xlsm"
Private Const SUBJECT_CUSTOM As String = "Mobius Emerging Opportunities Fund LP| Custom daily portfolio report"
Private Const SUBJECT_DAILY As String = "Mobius Emerging Opportunities Fund LP| Daily Reports"
```

## Data Flow (No External Dependencies)

As of v5.3, all data comes from the NAV reports - no Bloomberg Terminal required:
- Stock prices: Read from "Today USD" column in source file
- Option underlying prices: Looked up from stock positions in same report
- FX conversion: Applied for non-USD tickers (JP, LN, GY suffixes)

OCC format for options: `META  260116P00700000` (TICKER  YYMMDD[P/C]STRIKE)

## Testing

- Forward emails to yourself (handles FW: prefixes)
- Use `ProcessSelectedEmail` macro to manually process
- Use `ShowTrackerState` to verify email tracking
