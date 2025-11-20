# Bloomberg Data Transformer - Setup Guide

## Part 1: Installing the VBA Script (Windows)

### Step 1: Open Your Bloomberg Source File
1. On your Windows machine, open the Bloomberg export file (e.g., `Gain And Exposure_Custom_MOBIUS...xlsx`)
2. Ensure Bloomberg Terminal is running and you're logged in

### Step 2: Enable Developer Tools in Excel
1. Press `Alt + F11` to open VBA Editor
2. If that doesn't work:
   - File → Options → Customize Ribbon
   - Check "Developer" on the right panel
   - Click OK
   - Now click Developer tab → Visual Basic

### Step 3: Import the VBA Code
1. In VBA Editor (Alt + F11):
   - Insert → Module (this creates a new code module)
   - Copy the entire contents of `BloombergDataTransformer.vba`
   - Paste into the module window
   - File → Save (or Ctrl + S)

2. Save as Macro-Enabled Workbook:
   - Close VBA Editor (Alt + Q)
   - File → Save As
   - Change file type to "Excel Macro-Enabled Workbook (*.xlsm)"
   - Save to a location like: `C:\Bloomberg\PortfolioTransformer.xlsm`

### Step 4: Create a Button to Run the Macro

**Option A: Quick Access Button**
1. View → Macros → View Macros (or Alt + F8)
2. Select `TransformBloombergData`
3. Click "Options" and assign a shortcut key (e.g., Ctrl + Shift + T)
4. Click OK

**Option B: Button on Worksheet**
1. Developer tab → Insert → Button (Form Control)
2. Draw the button on your worksheet
3. In the "Assign Macro" dialog, select `TransformBloombergData`
4. Click OK
5. Right-click the button → Edit Text → Type "Transform Data"

### Step 5: Run the Transformation
1. Open a Bloomberg export file
2. Either:
   - Click your button, OR
   - Press Alt + F8, select `TransformBloombergData`, click Run, OR
   - Use your keyboard shortcut (Ctrl + Shift + T)

3. The script will:
   - Create a new workbook
   - Generate "Stocks" and "Options" tabs
   - Format the data
   - Save automatically with today's date

---

## Part 2: Understanding the Transformation

### What the Script Does

#### From Source File:
- Rows 1-5: Headers
- Row 6: USD Cash position
- Rows 7+: Mixed holdings (stocks, ETFs, and options)

#### To Output File:

**Stocks Tab:**
- Separates all equity holdings (stocks, ETFs)
- Columns:
  - Name, Ticker, Quantity
  - Unit Cost, Current Price (Bloomberg live data)
  - Total Cost, Market Value, P&L
  - Portfolio Weight %, Attribution %
- Cash positions at bottom

**Options Tab:**
- Separated into two sections: PUTS and CALLS
- For each option:
  - **Name**: Full option description
  - **Quantity**: Number of contracts
  - **Underlying Qty**: Shares of underlying stock you own
  - **% Hedged**: What % of your position is hedged
  - **Strike Px**: Extracted from option name
  - **Underlying Px**: Live price from Bloomberg
  - **% Moneyness**: How far in/out of the money
  - **Expiry**: Expiration date
  - **Unit Cost, Current Px, Mkt Value, P&L**

### Key Features

**1. Automatic Option Parsing**
The script reads option names like:
```
META 01/16/2026 PUT 700
```
And extracts:
- Ticker: META
- Expiry: 01/16/2026
- Type: PUT
- Strike: 700

**2. Bloomberg Integration**
Uses Bloomberg Excel formulas:
```vba
=BDP("AAPL US Equity","PX_LAST")
```
This pulls live data for:
- Stock prices
- Option underlying prices

**3. Automatic Calculations**
- **% Hedged** = (Option Contracts × 100) / Underlying Shares
- **% Moneyness** = (Underlying Price - Strike) / Strike
- **Market Value** = Current Price × Quantity

---

## Part 3: Troubleshooting

### Error: "User not logged into Bloomberg"
**Solution**: Ensure Bloomberg Terminal is running and you're logged in

### Error: "Cannot find Bloomberg Type Library"
**Solution**:
1. In VBA Editor: Tools → References
2. Check "Bloomberg API COM 3.0 Type Library"
3. If not visible, click Browse and navigate to:
   - `C:\blp\API\Office Tools\BloombergUI.xla`

### Error: "Macro security"
**Solution**:
1. File → Options → Trust Center → Trust Center Settings
2. Macro Settings → Enable all macros (or enable for trusted locations)
3. Add your Bloomberg folder to Trusted Locations

### Bloomberg Formulas Return #N/A
**Solution**:
- Check ticker format (e.g., "AAPL US Equity" not "AAPL")
- Ensure Bloomberg Terminal is connected
- Try manually typing `=BDP("AAPL US Equity","PX_LAST")` in a cell to test

### Options Not Separating Correctly
**Solution**:
- Check that option names contain " PUT " or " CALL " with spaces
- Ensure the format is: `TICKER DATE PUT/CALL STRIKE`

---

## Part 4: Customization

### Modify Column Mapping
In the `ProcessStock` subroutine, change these lines to map different columns:

```vba
wsTarget.Cells(targetRow, 2).Value = wsSource.Cells(sourceRow, 1).Value  ' Name
wsTarget.Cells(targetRow, 3).Value = wsSource.Cells(sourceRow, 2).Value  ' Ticker
' etc.
```

### Change Output File Location
In `TransformBloombergData` subroutine:

```vba
outputPath = "C:\Your\Custom\Path\Portfolio_" & todayDate & ".xlsx"
```

### Add Additional Bloomberg Fields
To pull more data from Bloomberg, add formulas like:

```vba
' Pull 52-week high
wsTarget.Cells(targetRow, 12).FormulaArray = "=BDP(""" & ticker & """,""HIGH_52WEEK"")"

' Pull dividend yield
wsTarget.Cells(targetRow, 13).FormulaArray = "=BDP(""" & ticker & """,""DVD_YIELD"")"

' Pull beta
wsTarget.Cells(targetRow, 14).FormulaArray = "=BDP(""" & ticker & """,""BETA_RAW_OVERRIDABLE"")"
```

---

## Part 5: Bloomberg Excel API Quick Reference

### Common Bloomberg Excel Functions

#### BDP (Bloomberg Data Point)
Get a single current data point:
```excel
=BDP("AAPL US Equity", "PX_LAST")           ' Last price
=BDP("AAPL US Equity", "VOLUME")            ' Volume
=BDP("AAPL US Equity", "PE_RATIO")          ' P/E Ratio
=BDP("AAPL US Equity", "DVD_YIELD")         ' Dividend Yield
=BDP("AAPL US Equity", "MARKET_CAP")        ' Market Cap
=BDP("AAPL US Equity", "52_WK_HIGH")        ' 52-week high
=BDP("AAPL US Equity", "BETA_RAW_OVERRIDABLE") ' Beta
```

#### BDH (Bloomberg Data History)
Get historical data:
```excel
=BDH("AAPL US Equity", "PX_LAST", "1/1/2024", "12/31/2024")
=BDH("SPY US Equity", "PX_LAST", TODAY()-365, TODAY())
```

#### BDS (Bloomberg Data Set)
Get multiple related data points:
```excel
=BDS("AAPL US Equity", "OPT_CHAIN")         ' Option chain
=BDS("AAPL US Equity", "DVD_HIST")          ' Dividend history
=BDS("AAPL US Equity", "EARN_ANN_EPS")      ' Earnings announcements
```

#### BDSV (Bloomberg Snapshot)
Real-time streaming data:
```excel
=BDSV("AAPL US Equity", "LAST_PRICE")       ' Live price
```

### Option-Specific Fields

For options (e.g., "AAPL 01/16/2026 C 250 Equity"):
```excel
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_STRIKE_PX")      ' Strike price
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_EXPIRE_DT")      ' Expiry date
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_IMPLIED_VOLATILITY") ' IV
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_DELTA")          ' Delta
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_GAMMA")          ' Gamma
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_VEGA")           ' Vega
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_THETA")          ' Theta
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_RHO")            ' Rho
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_UNDL_PX")        ' Underlying price
```

### Finding Field Names in Bloomberg Terminal

1. In Bloomberg Terminal, type: `FLDS<GO>`
2. Search for field names (e.g., "price", "dividend", "volume")
3. Copy the field mnemonic code (e.g., `PX_LAST`, `DVD_YIELD`)
4. Use in Excel: `=BDP("TICKER", "FIELD_CODE")`

### Using in VBA

```vba
' Single value
Dim lastPrice As Double
lastPrice = Application.Run("BDP", "AAPL US Equity", "PX_LAST")

' Array formula (for worksheet)
wsTarget.Cells(1, 1).FormulaArray = "=BDP(""AAPL US Equity"",""PX_LAST"")"

' Historical data
Dim histData As Variant
histData = Application.Run("BDH", "AAPL US Equity", "PX_LAST", "1/1/2024", "12/31/2024")
```

---

## Part 6: Enhanced Option Analytics

### Add These to Your Options Tab

You can extend the VBA script to include Greeks and IV:

```vba
' Add after line that sets Underlying Price
' Add Delta
wsTarget.Cells(targetRow, 16).FormulaArray = "=BDP(""" & ticker & " " & expiry & " " & optionType & " " & strike & " Equity"",""OPT_DELTA"")"

' Add Gamma
wsTarget.Cells(targetRow, 17).FormulaArray = "=BDP(""" & ticker & " " & expiry & " " & optionType & " " & strike & " Equity"",""OPT_GAMMA"")"

' Add IV
wsTarget.Cells(targetRow, 18).FormulaArray = "=BDP(""" & ticker & " " & expiry & " " & optionType & " " & strike & " Equity"",""OPT_IMPLIED_VOLATILITY"")"

' Add Theta
wsTarget.Cells(targetRow, 19).FormulaArray = "=BDP(""" & ticker & " " & expiry & " " & optionType & " " & strike & " Equity"",""OPT_THETA"")"
```

---

## Next Steps

1. ✅ VBA script created
2. ⏭️ Test on sample Bloomberg file
3. ⏭️ Customize for your specific needs
4. ⏭️ Learn Bloomberg Terminal navigation (see Part 7)
5. ⏭️ Build Black-Scholes pricer (separate guide)
6. ⏭️ Learn about call debit spreads (separate guide)
