# Bloomberg NAV Data Transformer

Complete solution for transforming daily NAV calculation files into structured portfolio reports with Bloomberg API integration.

---

## 🎯 At a Glance

**What:** Automated transformation of your daily NAV reports
**Input:** `Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP_[DATE].XLSX`
**Output:** `Transformed_Portfolio_[DATE].xlsx` with Stocks & Options tabs
**Time:** One-click, 10-30 seconds
**Works with:** ANY daily NAV file with the same format

**Key Features:**
- ✅ Separates stocks from options automatically
- ✅ Calculates option analytics (strike, expiry, % moneyness, % hedged)
- ✅ Bloomberg API integration for live prices and Greeks
- ✅ Handles long and short positions correctly
- ✅ One-time setup, works forever

---

## 📋 What This Does

Transforms your daily NAV export from:

**Source Format (Mixed):**
- Stocks, ETFs, and Options all in one list
- Basic metrics only
- No option analytics

**To Output Format (Organized):**
- **Stocks Tab:** Clean equity/ETF positions
- **Options Tab:** Separated PUTs and CALLs with full analytics
  - Strike price, Expiry, % Moneyness
  - % Hedged (vs underlying position)
  - Bloomberg live prices and Greeks

---

## 📁 Files Delivered

### 1. `BloombergDataTransformer_v2.vba` ⭐ **USE THIS ONE**
**Production-ready VBA script** with:
- ✅ Uses OCC ticker format from Column B (reliable)
- ✅ Handles negative quantities (short positions)
- ✅ Matches options to underlying stocks automatically
- ✅ Bloomberg API integration for live data
- ✅ Calculates % Hedged, % Moneyness, % Yield
- ✅ Proper formatting and error handling

### 2. `SETUP_GUIDE.md`
**Complete Windows setup instructions:**
- How to install the VBA script
- Creating a transformation button
- Bloomberg API configuration
- Troubleshooting guide
- Excel function reference

### 3. `BLOOMBERG_TERMINAL_GUIDE.md`
**Bloomberg Terminal navigation:**
- Essential commands
- Option analysis tools (OMON, GRKS, OSA)
- Excel integration (XLTP, BDP, BDH)
- Daily workflow examples
- Field reference guide

---

## 🚀 Quick Start (Windows)

### Step 1: One-Time Setup
1. Open Excel on Windows with Bloomberg Terminal running
2. Press `Alt + F11` to open VBA Editor
3. Insert → Module
4. Copy all code from `BloombergDataTransformer_v2.vba`
5. Paste into the module
6. Save as "Portfolio Transformer.xlsm" (macro-enabled)
   - Save location: `C:\Bloomberg\Portfolio Transformer.xlsm` (or your preferred location)

### Step 2: Add Button (Optional but Recommended)
1. Developer tab → Insert → Button
2. Assign macro: `TransformBloombergData`
3. Label it: "Transform NAV Data"
4. Close and save the workbook

---

## 📅 Daily Workflow - Transform ANY NAV Report

Every day when you receive your NAV export file:

### Method 1: Using the Saved Template (Easiest)

**Step 1:** Download/save your daily NAV export
- File format: `Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP_[DATE].XLSX`
- Example: `Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP_11102025.XLSX`
- Save to: Your Downloads folder or a dedicated NAV folder

**Step 2:** Open the saved transformer
1. Open `Portfolio Transformer.xlsm` (the file you saved in setup)
2. When prompted "Enable Macros" → Click **Enable**

**Step 3:** Open today's NAV file
1. File → Open → Select today's NAV export file
2. The NAV file opens in a new window

**Step 4:** Transform
1. Make sure the NAV file is the active window (click on it)
2. Click the "Transform NAV Data" button you created
   - OR press `Alt + F8` → Select `TransformBloombergData` → Run
3. Wait 10-30 seconds (depending on position count)

**Step 5:** Review output
- New file created: `Transformed_Portfolio_DD MMMM YYYY.xlsx`
- Location: Same folder as the source NAV file
- Contains:
  - **Stocks Tab:** All equity positions with Bloomberg live prices
  - **Options Tab:** All options separated into PUTS and CALLS sections

**Done!** Close the original NAV file (don't need to save it).

### Method 2: Direct from NAV File

**Alternative approach if you prefer:**

**Step 1:** Open today's NAV export file directly
- Double-click `Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP_[DATE].XLSX`

**Step 2:** Load the macro
1. Press `Alt + F11` to open VBA Editor
2. File → Import File
3. Select `BloombergDataTransformer_v2.vba`
4. Close VBA Editor

**Step 3:** Run transformation
1. Press `Alt + F8`
2. Select `TransformBloombergData`
3. Click Run

**Step 4:** Review output (same as Method 1)

---

## 📂 File Naming & Organization

### Expected Source File Format
Your NAV system exports files with this naming pattern:
```
Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP_[MMDDYYYY].XLSX
```

Examples:
- `Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP_11052025.XLSX` (Nov 5, 2025)
- `Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP_11102025.XLSX` (Nov 10, 2025)

### Output File Naming
The script automatically creates:
```
Transformed_Portfolio_DD MMMM YYYY.xlsx
```

Examples:
- `Transformed_Portfolio_05 November 2025.xlsx`
- `Transformed_Portfolio_10 November 2025.xlsx`

**Note:** If a file with the same name exists, the script adds a timestamp:
```
Transformed_Portfolio_YYYYMMDD_HHMMSS.xlsx
```

### Recommended Folder Structure

```
C:\Bloomberg\
├── Portfolio Transformer.xlsm          ← Your saved template with macro
├── NAV Reports\
│   ├── Gain And Exposure_..._11052025.XLSX
│   ├── Gain And Exposure_..._11102025.XLSX
│   └── Gain And Exposure_..._11152025.XLSX
└── Transformed Reports\
    ├── Transformed_Portfolio_05 November 2025.xlsx
    ├── Transformed_Portfolio_10 November 2025.xlsx
    └── Transformed_Portfolio_15 November 2025.xlsx
```

**Tip:** Customize output path in VBA (see Customization Options below)

---

## ✅ File Format Requirements

The script works with **any** NAV export file that has this structure:

**Required Structure:**
- **Rows 1-3:** Empty/formatting (ignored)
- **Row 4:** Sub-headers (Unit Cost, Today, etc.)
- **Row 5:** Column headers (Product Name, Ticker, ISIN, etc.)
- **Row 6+:** Position data (stocks, options, cash)

**Required Columns:**
| Column | Header | Required | Used For |
|--------|--------|----------|----------|
| A | Product Name | ✅ Yes | Identify stocks vs options |
| B | Ticker | ✅ Yes | OCC format for Bloomberg API |
| C | ISIN | Optional | Reference only |
| D | Portfolio Weight % | ✅ Yes | Output |
| E | Unit Cost USD | ✅ Yes | Output |
| F | Today USD | ✅ Yes | Current price |
| G | % Daily Gain/Loss | Optional | Skipped |
| H | Contribution to Performance | ✅ Yes | Attribution |
| I | Total Cost USD | ✅ Yes | Output |
| J | Market Value USD | ✅ Yes | Output |
| K | Total Net P&L YTD | ✅ Yes | Output |
| L | # of Shares | ✅ Yes | Critical for matching |

**As long as your daily NAV exports follow this format, the script will work!**

---

## 📊 Output Format

### Stocks Tab
| Name | Ticker | Quantity | Unit Cost | Current Px | Total Cost | Mkt Value | P&L | Portfolio Wgt | Attribution |
|------|--------|----------|-----------|------------|------------|-----------|-----|---------------|-------------|
| Meta Platforms Inc | META US | 5,158 | $675 | `=BDP()` | $3.49M | `=F*D` | -$232K | 6.0% | -0.4% |

### Options Tab - PUTS Section
| Name | Qty | Underlying Qty | % Hedged | Strike | Underlying Px | % Moneyness | Expiry | Unit Cost | % Yield | Total Cost | Current Px | Mkt Value | P&L |
|------|-----|----------------|----------|--------|---------------|-------------|--------|-----------|---------|------------|------------|-----------|-----|
| META 01/16/2026 PUT 700 | 27 | 5,158 | 52.3% | $700 | `=BDP()` | -9.1% | 01/16/2026 | $31.78 | 5.0% | $85,807 | $30.77 | $83,079 | -$2,728 |

### Options Tab - CALLS Section
| Name | Qty | Underlying Qty | % Hedged | Strike | Underlying Px | % Moneyness | Expiry | Unit Cost | % Yield | Total Cost | Current Px | Mkt Value | P&L |
|------|-----|----------------|----------|--------|---------------|-------------|--------|-----------|---------|------------|------------|-----------|-----|
| META 01/16/2026 CALL 805 | -27 | 5,158 | -52.3% | $805 | `=BDP()` | -20.9% | 01/16/2026 | $36.62 | 5.7% | -$98,873 | $22.12 | -$59,724 | $39,149 |

---

## 🔧 How It Works

### Data Flow

```
NAV Export File (Daily)
        ↓
[VBA Script Reads Data]
        ↓
┌───────────────────────┐
│ 1. Build Stock Dict   │  Maps: TICKER → # of shares
│    META → 5,158       │
│    UBER → 10,369      │
└───────────────────────┘
        ↓
┌───────────────────────┐
│ 2. Separate Stocks    │  Identify: No "PUT" or "CALL" in name
│    vs Options         │
└───────────────────────┘
        ↓
┌───────────────────────┐
│ 3. Process Stocks     │  Output → "Stocks" tab
│    - Copy data        │  - Live prices via Bloomberg
│    - Add formulas     │
└───────────────────────┘
        ↓
┌───────────────────────┐
│ 4. Process Options    │  Output → "Options" tab
│    - Parse name       │  - Extract: Strike, Expiry, Type
│    - Use OCC ticker   │  - Bloomberg: Greeks, IV, Price
│    - Match underlying │  - Calculate: % Hedged, Moneyness
│    - Separate P/C     │
└───────────────────────┘
        ↓
Formatted Excel Output
```

### Key Algorithms

**1. Option Identification:**
```vba
If InStr(productName, " PUT ") > 0 Or InStr(productName, " CALL ") > 0 Then
    ' It's an option
End If
```

**2. Underlying Matching:**
```vba
' Extract "META" from "META 01/16/2026 PUT 700"
baseTicker = ExtractTickerFromOptionName(productName)

' Lookup shares from stock dictionary
underlyingShares = stockPositions(baseTicker) ' Returns 5,158 for META
```

**3. Bloomberg API Calls:**
```vba
' Use OCC ticker from Column B
occTicker = "META  260116P00700000"

' Get underlying price
=BDP("META  260116P00700000 Equity", "OPT_UNDL_PX")

' Get option Greeks (optional enhancement)
=BDP("META  260116P00700000 Equity", "OPT_DELTA")
=BDP("META  260116P00700000 Equity", "OPT_IMPLIED_VOLATILITY")
```

**4. % Hedged Calculation:**
```vba
' For PUTS (long protection):
% Hedged = (Contracts × 100) / Underlying Shares
         = (27 × 100) / 5,158
         = 52.3%

' For CALLS (short = negative):
% Hedged = -(Contracts × 100) / Underlying Shares
         = -(-27 × 100) / 5,158
         = 52.3%
```

**5. % Moneyness:**
```vba
% Moneyness = (Underlying Price - Strike Price) / Strike Price

' Example: PUT at $700 strike, underlying at $636
= ($636 - $700) / $700
= -9.1% (out of the money)
```

---

## 🎯 Source File Schema (NAV Export)

### Column Mapping

| Col | Header | Description | Used For |
|-----|--------|-------------|----------|
| A | Product Name | Full name | Identify stocks vs options |
| B | Ticker | OCC format for options | **Bloomberg API calls** |
| C | ISIN | Security identifier | Reference |
| D | Portfolio Weight % | Position size | Output |
| E | Unit Cost USD | Average cost | Output |
| F | Today USD | Current price | Output (can refresh via Bloomberg) |
| G | % Daily Gain/Loss | Daily change | Skip |
| H | Contribution to Performance | Attribution | Output |
| I | Total Cost USD | Cost basis | Output |
| J | Market Value USD | Current value | Output |
| K | Total Net P&L YTD | Year-to-date P&L | Output |
| L | # of Shares | **Shares or Contracts** | Critical for matching |

### Sample Rows

**Row 6:** USD Cash
```
USD | | | 0.144 | 1 | 1 | | 0 | 7,882,960 | 7,882,960 | 0 | 7,882,960
```

**Row 10:** Stock Position
```
Meta Platforms Inc | META US | US30303M1027 | 0.06 | 675 | 636 | 0.014 | -0.004 | 3,493,599 | 3,280,230 | -232,941 | 5,158
```

**Row 30:** PUT Option (Long)
```
META 01/16/2026 PUT 700 | META  260116P00700000 | US30303M1027 | 0.004 | 32 | 77 | -0.089 | 0.002 | 85,807 | 206,820 | 121,013 | 27
```

**Row 42:** CALL Option (Short)
```
UBER 01/16/2026 CALL 105 | UBER  260116C00105000 | US90353T1007 | 0 | 6 | 2 | -0.311 | 0.001 | -57,369 | -20,085 | 37,284 | -103
```

**Note:** Column L (# of Shares) is:
- **Positive** for long positions
- **Negative** for short positions

---

## 🔍 Bloomberg API Enhancement (Optional)

Want more option analytics? Add these to the VBA script:

### Additional Bloomberg Fields

```vba
' In ProcessOption subroutine, add after line setting Underlying Price:

' Column P: Delta
wsTarget.Cells(targetRow, 16).FormulaArray = "=BDP(""" & occTicker & " Equity"",""OPT_DELTA"")"

' Column Q: Gamma
wsTarget.Cells(targetRow, 17).FormulaArray = "=BDP(""" & occTicker & " Equity"",""OPT_GAMMA"")"

' Column R: Theta (time decay)
wsTarget.Cells(targetRow, 18).FormulaArray = "=BDP(""" & occTicker & " Equity"",""OPT_THETA"")"

' Column S: Vega (vol sensitivity)
wsTarget.Cells(targetRow, 19).FormulaArray = "=BDP(""" & occTicker & " Equity"",""OPT_VEGA"")"

' Column T: Implied Volatility
wsTarget.Cells(targetRow, 20).FormulaArray = "=BDP(""" & occTicker & " Equity"",""OPT_IMPLIED_VOLATILITY"")"
```

Then update `SetupOptionsHeaders` to add column headers for these fields.

---

## ⚙️ Customization Options

### Change Output File Location
```vba
' In TransformBloombergData, find:
outputPath = Application.ActiveWorkbook.Path & "\Transformed_Portfolio_" & todayDate & ".xlsx"

' Change to:
outputPath = "C:\Portfolio Reports\Transformed_Portfolio_" & todayDate & ".xlsx"
```

### Disable Bloomberg Live Prices (Use NAV Prices)
```vba
' In ProcessStock, find:
wsTarget.Cells(targetRow, 6).FormulaArray = "=BDP(""" & ticker & " Equity"",""PX_LAST"")"

' Replace with:
wsTarget.Cells(targetRow, 6).Value = wsSource.Cells(sourceRow, 6).Value
```

### Add More Stock Columns
```vba
' Example: Add Market Cap
wsTarget.Cells(targetRow, 12).FormulaArray = "=BDP(""" & ticker & " Equity"",""CUR_MKT_CAP"")"

' Example: Add 52-Week High
wsTarget.Cells(targetRow, 13).FormulaArray = "=BDP(""" & ticker & " Equity"",""HIGH_52WEEK"")"
```

---

## 🐛 Troubleshooting

### Issue: "Compile Error: Can't find project or library"
**Cause:** Missing reference
**Solution:**
1. VBA Editor → Tools → References
2. Uncheck any references marked as "MISSING"
3. Check "Microsoft Scripting Runtime"

### Issue: Bloomberg formulas return #N/A
**Cause:** Bloomberg Terminal not running or not logged in
**Solution:**
1. Launch Bloomberg Terminal
2. Log in
3. Verify connection: Type `AAPL US Equity <GO>`
4. Re-run transformation

### Issue: Options not matching to underlying stocks
**Cause:** Ticker mismatch
**Solution:**
1. Check that stock ticker in Column B matches option ticker prefix
2. Example: Stock should be "META US", option should be "META  260116P00700000"
3. The script extracts "META" from both

### Issue: % Hedged shows "N/A"
**Cause:** No underlying stock position found
**Solution:**
- Normal for options without underlying stock positions
- If you do have the stock, check ticker mapping

### Issue: Negative % Hedged for long puts
**Cause:** Script issue
**Solution:** Check the formula in column E, should be positive for long puts

---

## 📝 Next Steps

### Immediate (Windows Machine)
1. ✅ Install VBA script following SETUP_GUIDE.md
2. ✅ Test with sample NAV file
3. ✅ Create transformation button
4. ✅ Run first transformation

### Short Term
1. ⏭️ Set up daily export from NAV system to a standard folder
2. ⏭️ Customize output format to your preferences
3. ⏭️ Add Bloomberg Greeks if desired (Delta, Gamma, Theta, Vega, IV)
4. ⏭️ Create custom Bloomberg views in CUST<GO>

### Long Term
1. ⏭️ Automate: VBA script to auto-run on file open
2. ⏭️ Dashboard: Create summary dashboard with charts
3. ⏭️ Alerts: Bloomberg alerts for position changes
4. ⏭️ Risk: Integrate portfolio Greeks for total risk exposure

---

## 📚 Additional Resources

- **Bloomberg Terminal Guide:** See `BLOOMBERG_TERMINAL_GUIDE.md`
- **Setup Instructions:** See `SETUP_GUIDE.md`
- **VBA Source Code:** `BloombergDataTransformer_v2.vba`

---

## 🆘 Support

### Bloomberg Terminal Help
```
HELP<GO>           - Context help
F1 F1              - Live chat with Bloomberg support
FLDS<GO>           - Search for data fields
DAPI<GO>           - Developer resources
```

### VBA Script Issues
Check the error message and refer to:
1. Error handling in the script shows line numbers
2. SETUP_GUIDE.md troubleshooting section
3. Verify Bloomberg Terminal is running and connected

---

## ✅ Summary

You now have a complete solution to:
- ✅ Transform NAV exports into structured portfolio reports
- ✅ Separate stocks and options with proper categorization
- ✅ Calculate option analytics (strike, expiry, moneyness, hedging)
- ✅ Integrate Bloomberg API for live data
- ✅ Automate daily workflow with one-click button

**All files are ready to use on your Windows machine!**
