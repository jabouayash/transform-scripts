# Backlog

Product backlog for the Mobius Portfolio Reporter.

---

## High Priority

### Line Chart X-Axis Scaling
- **Issue**: As more dates accumulate in Performance_History.csv, the x-axis will get crowded
- **Current behavior**: Plots all data points, Excel auto-compresses labels
- **Proposed solution**: Limit to last 90 days, or add weekly aggregation
- **When to address**: After 60+ days of data accumulates

### Line Chart Positioning
- **Issue**: Line chart Left position (420) doesn't perfectly align with other charts (250)
- **Context**: Left:=250 causes overlap with data table, Left:=420 clears data but misaligns
- **Proposed solution**: Investigate why performance data table is positioned differently, or hide data table since chart shows same info
- **Workaround**: Current Left:=420 works, just slightly misaligned

---

## Medium Priority

### Column Width Fine-Tuning
- **Issue**: Balance between narrow columns (John's feedback) and readable content
- **Current state**: Name column at 25 with text wrap, Ticker at 8 with wrap
- **May need**: Further adjustment based on actual data lengths

---

## Low Priority / Future Enhancements

### Performance Chart Enhancements
- Add YTD return % as secondary line
- Add benchmark comparison (S&P 500)
- Interactive date range selection (would require different technology)

### Dashboard Improvements
- Add MTD (month-to-date) metrics
- Add QTD (quarter-to-date) metrics
- Sector allocation pie chart

### Export Options
- PDF export of dashboard
- Email summary generation

---

## Completed

### v5.6.1 - 2026-01-01
- [x] Performance line chart added
- [x] Text wrap on Name/Ticker columns
- [x] Narrower column widths
- [x] CSV line ending fix (Unix/Windows)

### v5.6.0 - 2025-12-23
- [x] Performance history tracking (CSV)
- [x] Dashboard version indicator

---

## Notes

- Performance_History.csv location: `C:\Mobius Reports\Performance_History.csv`
- Chart positions use pixel values (Left, Top, Width, Height)
- Dashboard column widths: A=40, B=15, C=18, D-G=12
