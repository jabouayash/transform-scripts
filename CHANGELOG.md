# Changelog

All notable changes to the Mobius Portfolio Reporter are documented here.

Format: [Semantic Versioning](https://semver.org/) - MAJOR.MINOR.PATCH
- MAJOR: Breaking changes (new input format, incompatible output)
- MINOR: New features (backward compatible)
- PATCH: Bug fixes

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
