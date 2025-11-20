# Bloomberg Terminal Navigation Guide

## Part 1: Bloomberg Terminal Basics

### Understanding the Bloomberg Command System

Bloomberg Terminal uses a unique command structure:
```
COMMAND<GO>
```

- Type the command
- Press `<GO>` (green button on Bloomberg keyboard, or typically F9/Enter)

### Essential Navigation Commands

| Command | Description | Example Use |
|---------|-------------|-------------|
| `HELP<GO>` | Help system | Get started guide |
| `BYE<GO>` | Log out | End session |
| `MENU<GO>` | Main menu | Navigate all functions |
| `MSG<GO>` | Bloomberg messenger | Chat with other users |
| `GRAB<GO>` | Screenshot tool | Capture terminal screens |
| `DOCS<GO>` | Bloomberg University | Tutorials and guides |

---

## Part 2: Security Lookup and Analysis

### Finding Securities

**1. Equity Ticker Search:**
```
AAPL US <Equity> <GO>
```
- `AAPL` = Ticker symbol
- `US` = Country/Exchange (US, JP, LN, etc.)
- `<Equity>` = Yellow key on Bloomberg keyboard

**Alternative method:**
```
AAPL <GO>
```
Bloomberg will show a list of matches; select the correct one.

**2. Quick Search:**
```
NAME<GO>
```
Type company name and search.

**3. ISIN/CUSIP Lookup:**
```
US0378331005 <GO>
```
Enter ISIN or CUSIP directly.

### Key Equity Analysis Functions

| Function | Command | Description |
|----------|---------|-------------|
| **Quote** | `AAPL <Equity> <GO>` | Current price, volume, market data |
| **Description** | `AAPL <Equity> DES<GO>` | Company profile, description |
| **Graph** | `AAPL <Equity> GP<GO>` | Price chart, technical analysis |
| **Fundamentals** | `AAPL <Equity> FA<GO>` | Financial statement analysis |
| **Ownership** | `AAPL <Equity> OWN<GO>` | Major shareholders |
| **Estimates** | `AAPL <Equity> EEO<GO>` | Analyst estimates, consensus |
| **News** | `AAPL <Equity> N<GO>` | Latest news |
| **Earnings** | `AAPL <Equity> ERN<GO>` | Earnings history |
| **Dividend** | `AAPL <Equity> DVD<GO>` | Dividend information |
| **Relative Value** | `AAPL <Equity> RV<GO>` | Compare to peers |

---

## Part 3: Options Analysis

### Finding Options

**1. Basic Option Lookup:**
```
AAPL <Equity> OMON<GO>
```
- Opens Options Monitor
- Shows all available options
- Real-time prices, Greeks, IV

**2. Option Chain:**
```
AAPL <Equity> CALL<GO>
```
Shows call options chain

```
AAPL <Equity> PUT<GO>
```
Shows put options chain

**3. Specific Option Contract:**
```
AAPL 01/16/2026 C 250 <Equity> <GO>
```
Format: `TICKER EXPIRY C/P STRIKE <Equity> <GO>`
- `C` = Call
- `P` = Put
- `250` = Strike price

### Key Options Functions

| Function | Command | Description |
|----------|---------|-------------|
| **Options Monitor** | `OMON<GO>` | Real-time option chains |
| **Option Valuation** | `OVME<GO>` | Theoretical value calculator |
| **Scenario Analysis** | `OSA<GO>` | What-if scenarios |
| **Implied Volatility** | `HIVG<GO>` | Historical vs implied vol |
| **Greeks** | `GRKS<GO>` | Delta, gamma, theta, vega, rho |
| **Option Strategy** | `OMON<GO>` then F9 | Build multi-leg strategies |
| **Skew Graph** | `SKEW<GO>` | Volatility skew |

### Options Monitor (OMON) - Detailed Guide

1. **Open Options Monitor:**
   ```
   AAPL <Equity> OMON<GO>
   ```

2. **Key Sections:**
   - **Left Panel**: Strike prices
   - **Middle**: Calls (bid, ask, last, volume, IV, Greeks)
   - **Right**: Puts (bid, ask, last, volume, IV, Greeks)

3. **Useful Shortcuts in OMON:**
   - `F9`: Build custom strategy
   - `F10`: Change expiration date
   - `Ctrl+E`: Export to Excel
   - `Ctrl+G`: Graph profit/loss

4. **Building a Strategy:**
   - Click on an option to select it
   - Specify quantity (negative for short)
   - Click "Add to Strategy"
   - Repeat for multi-leg strategies
   - View P&L diagram

---

## Part 4: Excel Integration

### Bloomberg Excel Add-In Functions

**1. Activate Excel with Bloomberg:**
```
XLTP<GO>
```
- Downloads Bloomberg Excel Add-In
- Restarts Excel with Bloomberg ribbon

**2. Key Excel Functions (type in cell):**

```excel
=BDP("AAPL US Equity", "PX_LAST")           ' Last price
=BDP("AAPL US Equity", "VOLUME")            ' Today's volume
=BDH("AAPL US Equity", "PX_LAST", "1/1/2024", TODAY())  ' Historical prices
```

**3. Option Data in Excel:**

```excel
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_DELTA")
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_IMPLIED_VOLATILITY")
=BDP("AAPL 01/16/2026 C 250 Equity", "OPT_STRIKE_PX")
```

**4. Portfolio Functions:**

```excel
=BDP("AAPL US Equity", "PX_LAST") * [Shares]      ' Market value
=BDP("AAPL US Equity", "TOT_RETURN_INDEX_GROSS")  ' Total return
```

### Exporting Data to Excel

**Method 1: Direct Export**
1. Open any Bloomberg screen
2. Press `Ctrl + E` (or use Actions menu → Export)
3. Select "Excel"
4. Choose what to export (table, chart, etc.)

**Method 2: Using Bloomberg Templates**
```
DAPI<GO>
```
- Download pre-built Excel templates
- Templates with Bloomberg formulas
- Equity analysis, fixed income, etc.

**Method 3: Using PORT (Portfolio)**
```
PORT<GO>
```
- Build your portfolio
- Add securities
- Export → Excel
- Auto-generates portfolio report

---

## Part 5: Portfolio Management

### Creating and Managing Portfolios

**1. Portfolio Manager:**
```
PORT<GO>
```

**2. Create New Portfolio:**
- Click "Create Portfolio"
- Name it (e.g., "MOBIUS EMERGING OPPORTUNITIES")
- Add securities:
  - Type ticker → Enter quantity → Add
  - Repeat for all holdings

**3. Portfolio Analytics:**
```
PRTU<GO>
```
- Upload portfolio positions
- Run analytics (risk, attribution, performance)

**4. Custom Reports:**
```
CUST<GO>
```
- Build custom portfolio reports
- Select metrics to display
- Save templates for daily use

**5. Export Portfolio Data:**
- In `PORT<GO>`, click Actions → Export
- Choose "Custom Export"
- Select fields:
  - Position quantity
  - Market value
  - Cost basis
  - P&L
  - Greeks (for options)
- Export to Excel

---

## Part 6: Market Data and Research

### Real-Time Market Data

| Function | Command | Description |
|----------|---------|-------------|
| **Market Overview** | `WEI<GO>` | World equity indices |
| **Sector Performance** | `HS<GO>` | Heat map by sector |
| **Top Movers** | `MOST<GO>` | Biggest gainers/losers |
| **Market Monitor** | `ALLQ<GO>` | Create custom monitor |
| **Economic Calendar** | `ECO<GO>` | Upcoming economic releases |
| **Earnings Calendar** | `EARN<GO>` | Upcoming earnings |

### Research and Analysis

| Function | Command | Description |
|----------|---------|-------------|
| **News** | `N<GO>` | All news |
| **Research** | `RES<GO>` | Analyst research reports |
| **Alerts** | `ALRT<GO>` | Set price/news alerts |
| **Comparable Companies** | `COMP<GO>` | Find peers |
| **Relative Valuation** | `RV<GO>` | Compare metrics to peers |

---

## Part 7: Advanced Functions for Your Use Case

### Daily Workflow for Portfolio Management

**Morning Routine:**

1. **Check Market Overview:**
   ```
   WEI<GO>
   ```
   See how global markets are performing.

2. **Check Your Portfolio:**
   ```
   PORT<GO>
   ```
   Review overnight changes.

3. **Update Positions:**
   - Add new trades
   - Update quantities
   - Note: You can automate this via Bloomberg API

4. **Export Daily Report:**
   ```
   CUST<GO>
   ```
   Use your saved template "Gain And Exposure_Custom"
   - File → Export to Excel
   - Save with date: `Portfolio_MMDDYYYY.xlsx`

5. **Run Your VBA Transformer:**
   - Open exported file
   - Run `TransformBloombergData` macro
   - Review Stocks and Options tabs

### Options-Specific Daily Workflow

1. **Check Options Positions:**
   ```
   OMON<GO>
   ```
   For each underlying (AAPL, META, GOOGL, etc.)

2. **Monitor Greeks:**
   - Focus on Delta (directional exposure)
   - Watch Theta (time decay)
   - Check Gamma (delta sensitivity)

3. **Review Implied Volatility:**
   ```
   HIVG<GO>
   ```
   Compare current IV to historical levels.

4. **Check Expiring Positions:**
   - Filter by expiration date
   - Decide: Roll, exercise, or close

5. **Scenario Analysis:**
   ```
   OSA<GO>
   ```
   Model potential outcomes:
   - Stock moves up 10%
   - Stock moves down 10%
   - 30 days pass (time decay)

---

## Part 8: Keyboard Shortcuts

### Essential Bloomberg Shortcuts

| Shortcut | Action |
|----------|--------|
| `<GO>` | Execute command (F9 or Enter) |
| `<HELP>` | Context-sensitive help |
| `Ctrl + N` | Open new window |
| `Ctrl + W` | Close current window |
| `Ctrl + E` | Export to Excel |
| `Ctrl + G` | Graph/Chart |
| `Ctrl + P` | Print |
| `Alt + 1` | Window 1 |
| `Alt + 2` | Window 2 |
| `Page Up/Down` | Scroll through data |
| `Home/End` | Go to start/end |

### Bloomberg Keyboard Keys

If you have a Bloomberg keyboard:

- **Yellow Key** = Equity
- **Green Key** = Corp Bonds
- **Red Key** = Index
- **Blue Key** = Government
- **Orange Key** = Commodities
- **White Key** = Currencies
- **Pink Key** = Preferred

Without Bloomberg keyboard, use:
```
<EQUITY>, <INDEX>, <GOVT>, etc.
```

---

## Part 9: Getting Help

### Bloomberg Support

**1. Help Function:**
```
HELP<GO>
```

**2. Instant Bloomberg (IB) - Live Chat:**
```
MSG<GO>
```
Then type "Bloomberg Help Desk" and ask questions.

Or press:
```
F1 F1
```
(Press F1 twice quickly)

**3. Bloomberg Terminal Training:**
```
DOCS<GO>
```
- Free tutorials
- Video guides
- Certification courses

**4. Field Finder:**
```
FLDS<GO>
```
Search for any data field to use in Excel.

**5. Function Finder:**
```
SECF<GO>
```
Search by keyword (e.g., "option", "dividend", "volatility").

---

## Part 10: Your Specific Use Cases

### Use Case 1: Daily Portfolio Export for VBA Transformation

**Setup (One Time):**
1. Create custom view:
   ```
   CUST<GO>
   ```
2. Select your fund: "MOBIUS EMERGING OPPORTUNITIES FUND LP"
3. Configure columns:
   - Product Name
   - Ticker
   - ISIN
   - Portfolio Weight %
   - Unit Cost USD
   - Today USD
   - % Daily Gain/Loss
   - Contribution to Performance
   - Total Cost USD
   - Market Value USD
   - Total Net Profit and Loss YTD
   - # of Shares
4. Save as template: "Gain And Exposure_Custom"

**Daily Export:**
1. Open: `CUST<GO>`
2. Load template: "Gain And Exposure_Custom"
3. Actions → Export to Excel
4. Save to designated folder
5. Run VBA macro `TransformBloombergData`

### Use Case 2: Option Greeks Monitoring

**For each underlying with options:**
```
[TICKER] <Equity> OMON<GO>
```

**Add to your watchlist:**
```
BDH<GO>
```
Create a custom page with all your option positions.

**Export Greeks to Excel:**
- Use `OMON<GO>` → Ctrl+E → Excel
- Or use BDP formulas in your transformed sheet

### Use Case 3: Risk Analysis

**Portfolio Risk:**
```
PORT<GO> → PRSK<GO>
```

**Options Risk (Portfolio Greeks):**
```
OMON<GO>
```
- View aggregate portfolio Greeks
- See total Delta, Gamma, Theta exposure

**Scenario Analysis:**
```
OSA<GO>
```
- Model market moves
- Stress test your portfolio

---

## Part 11: Tips and Best Practices

### Pro Tips

1. **Save Your Layouts:**
   - Bloomberg remembers your open windows
   - Create different "Desks" for different workflows
   - Desktop → Save Current Desktop

2. **Use Launchpad:**
   ```
   BLP<GO>
   ```
   - Add your most-used functions
   - Quick access icons

3. **Set Up Alerts:**
   ```
   ALRT<GO>
   ```
   - Price targets
   - News mentions
   - Earnings announcements
   - Option IV changes

4. **Create Custom Monitors:**
   ```
   ALLQ<GO>
   ```
   - Build watchlists
   - Real-time updating
   - Customizable columns

5. **Use Bloomberg Anywhere:**
   - Access terminal from web browser
   - Remote work capability
   - Same functionality as desktop

### Common Mistakes to Avoid

1. **Incorrect Ticker Format:**
   - ❌ `AAPL` → ✅ `AAPL US Equity`
   - ❌ `SPY` → ✅ `SPY US Equity`

2. **Wrong Date Format:**
   - Bloomberg uses: `MM/DD/YYYY`
   - Excel might use different format

3. **Forgetting <GO>:**
   - Always end commands with `<GO>` (Enter/F9)

4. **Not Saving Custom Views:**
   - Save templates to avoid rebuilding daily

5. **Mixing Field Codes:**
   - Use `FLDS<GO>` to verify exact field names

---

## Quick Reference Card

### Your Daily Commands

```
PORT<GO>                  → Portfolio
CUST<GO>                  → Custom report export
OMON<GO>                  → Options monitor
AAPL <Equity> GP<GO>      → Chart
XLTP<GO>                  → Excel add-in
FLDS<GO>                  → Field search
HELP<GO>                  → Help
MSG<GO>                   → Support chat
```

### Next Steps

1. ✅ Set up your custom export template in `CUST<GO>`
2. ⏭️ Test daily export workflow
3. ⏭️ Build option Greeks monitoring dashboard
4. ⏭️ Set up alerts for your positions
5. ⏭️ Integrate with VBA automation
