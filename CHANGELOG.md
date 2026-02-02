# Changelog

All notable changes to the Mobius Portfolio Reporter are documented here.

Format: [Semantic Versioning](https://semver.org/) - MAJOR.MINOR.PATCH
- MAJOR: Breaking changes (new input format, incompatible output)
- MINOR: New features (backward compatible)
- PATCH: Bug fixes

---

## [5.7.1] - 2026-02-01
### Added
- Auto-open output file after transformation (file now appears in front automatically)

### Fixed
- Synced distribution files with source code (v5.7.0 dist was missing YTD Column H changes)

---

## [5.7.0] - 2026-01-22
### Added
- New "Currencies" tab for cash positions (USD, CAD, JPY, EUR, GBP, AED, MAD, etc.)
- New "Other" tab for admin/expense items (Accrued Dividends, Payables, Management Fees, etc.)

### Changed
- **Simplified workflow**: Only Custom email required (Daily Reports email no longer needed)
- Outlook monitor now triggers immediately on Custom email (no waiting for second email)
- Output filename now includes both report date and processing date: `Portfolio_MMDDYYYY_gen_MMDDYYYY.xlsx`
- Removed Fund Performance Summary from Stocks tab
- Stocks column A expanded to 45 width (names on single line, no text wrap)
- Top Holdings bar chart now shows biggest holdings at top (reversed axis)
- Gainers/Losers bar chart now shows gainers at top (reversed axis)
- Gainers/Losers chart labels positioned at far left (xlTickLabelPositionLow) for readability
- Gainers/Losers chart names truncated to 20 chars to prevent overflow
- Performance chart now shows YTD Return % instead of portfolio value
- Performance chart Y-axis uses 2% increments
- YTD Return now reads from Custom file Column H "Jan 1 ROR" (dynamically finds last row)
- Removed dependency on non-custom file for YTD Return
- Removed `SetDailyFilePath` function (no longer needed)

### Fixed
- Total Portfolio Value now includes: Stocks + Options + Currencies + Other items
- YTD Return now displays correct fund return from source data
- YTD Return location is now robust (validates Column D = 1.0 to find total row)

---

## [5.6.3] - 2026-01-11
### Fixed
- Total Portfolio Value now includes cash positions (USD, CAD, etc.) that were missing
- Total Portfolio Value now includes Options market value (was only including P&L)
- Portfolio Allocation section now shows dollar signs on values

---

## [5.6.2] - 2026-01-01
### Changed
- Removed empty row 1 from Stocks and Options tabs (headers now start at row 1)
- Data now starts at row 3 for Stocks, row 4 for Options (saves one row)

---

## [5.6.1] - 2026-01-01
### Fixed
- Performance line chart now positioned at top-right of Dashboard (was overlapping data)
- Name column now wraps text to new line instead of cutting off (per John's feedback)
- CSV line ending handling works with both Unix and Windows formats

---

## [5.6.0] - 2025-12-23
### Added
- Performance history tracking (saves daily metrics to Performance_History.csv)
- Performance line chart showing portfolio value over time (Robinhood-style)
- Chart appears automatically after 2+ days of data accumulates
- Version indicator on Dashboard title (v5.6)

### Changed
- Narrower columns with text-wrapped headers for better readability
- Numeric columns now have fixed widths (8-12) instead of AutoFit
- Header rows allow text wrapping for longer column names

---

## [5.5.0] - 2025-12-22
### Fixed
- Dashboard YTD P&L now includes Options P&L (was showing only Stocks)
- Portfolio Allocation "Other" calculation now sums correctly to Total Portfolio Value
- Options expiry dates now display as MM/DD/YYYY (was showing serial numbers)
- #DIV/0! errors prevented with IFERROR wrapper on % Diff formula
- Holdings count now excludes cash positions (USD, JPY, CAD, EUR, GBP)
- GOOG options now correctly look up GOOGL stock price (ticker alias)

---

## [5.4.0] - 2025-12-22
### Added
- Dashboard sheet with charts (appears first in workbook)
- KPI summary section: Total Portfolio Value, YTD P&L, YTD Return %, Holdings count
- Bar chart: Top 10 holdings by market value
- Pie chart: Portfolio allocation (Top 5 + Other)
- Bar chart: YTD P&L by position (top gainers and losers)

---

## [5.3.0] - 2025-12-22
### Removed
- Bloomberg Terminal dependency - all data now comes from NAV reports
- BDP() formulas for stock prices (now uses "Today USD" from source)
- BDP() formulas for option underlying prices (now looks up from stock dictionary)

### Changed
- Renamed project from "Bloomberg Portfolio Data Transformer" to "Portfolio Data Transformer"

---

## [5.2.0] - 2025-12-22
### Changed
- Removed blank column A (data starts in column A now)
- Reordered Stocks columns: Name, Ticker, Portfolio Wgt, % Diff, Daily Chg, Unit Cost, Current Px, Total Cost, Mkt Value, P&L, Attribution
- Removed currency symbols from cells (numbers only, headers indicate USD)
- Narrowed Name column width to 30
- Rounded prices to nearest dollar

### Removed
- Yield % column from Options tab

### Fixed
- Japan FX conversion formula (was showing incorrect % returns)
- Added ($) to P&L header on Options tab

---

## [5.1.0] - 2025-12-19
### Added
- FX conversion for non-USD tickers (JP, LN, GY, etc.)
- Automatic currency conversion to USD for Japanese, European, and other foreign stocks

---

## [5.0.0] - 2025-12-07
### Added
- Professional formatting matching input file styling
- Navy blue sub-headers with white text (#003366)
- Alternating row colors (white/light gray zebra striping)
- Consistent font colors and borders

---

## [4.0.0] - 2025-12-03
### Added
- Outlook email monitor integration (automatic triggering)
- YTD Fund Return from non-custom file (cell K94)
- Output saved to C:\Mobius Reports\Transformed\ folder
- Can be run manually or triggered by Outlook VBA

---

## [3.0.0] - 2025-11-26
### Added
- Multi-file support (Custom + non-Custom NAV files)
- DailyRor file reading for additional metrics
- Fund Performance Summary section

---

## Version Tags

To checkout a specific version:
```bash
git checkout v5.3.0
```

To see all available versions:
```bash
git tag -l
```
