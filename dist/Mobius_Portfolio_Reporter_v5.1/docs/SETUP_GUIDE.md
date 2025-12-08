# Mobius Portfolio Reporter - Setup Guide v5.1

## Overview
This system automatically transforms NAV report emails into formatted portfolio reports with live Bloomberg data.

---

## Prerequisites
- Bloomberg Terminal installed and logged in
- Bloomberg Excel Add-in enabled
- Microsoft Outlook
- Microsoft Excel with macro support

---

## Step 1: Run the Installer (Automated)

1. Double-click `INSTALL.bat`
2. If prompted by Windows Security, click "Run anyway"
3. The script will:
   - Create `C:\Mobius Reports\` folder structure
   - Copy the Portfolio Transformer workbook

---

## Step 2: Set Up Outlook Email Monitor (Manual)

This step requires pasting code into Outlook's VBA editor.

### 2.1 Open Outlook VBA Editor
1. Open Microsoft Outlook
2. Press `Alt + F11` to open the VBA editor

### 2.2 Paste the Monitor Code
1. In the left panel, find `Microsoft Outlook Objects`
2. Double-click `ThisOutlookSession`
3. Open the file `files\OutlookMonitor.txt` from this package
4. Copy ALL the code from that file
5. Paste it into the `ThisOutlookSession` window

### 2.3 Enable Macros in Outlook
1. In Outlook, go to `File > Options > Trust Center`
2. Click `Trust Center Settings`
3. Select `Macro Settings`
4. Choose `Enable all macros` (or `Notifications for all macros`)
5. Click OK and restart Outlook

---

## Step 3: Enable Bloomberg Excel Add-in

### 3.1 Check if Already Enabled
1. Open Excel
2. Look for a `Bloomberg` tab in the ribbon
3. If present, skip to Step 4

### 3.2 Enable the Add-in
1. In Excel, go to `File > Options > Add-ins`
2. At the bottom, select `COM Add-ins` and click `Go`
3. Check the box for `Bloomberg Excel Tools`
4. Click OK

### 3.3 Troubleshooting
If the Bloomberg add-in shows an error:
1. Open Bloomberg Terminal
2. Type `DAPI <GO>`
3. Click `Install Office Add-ins`
4. Select your Excel version (32-bit or 64-bit)
5. Complete the installation and restart Excel

---

## Step 4: Test the System

### 4.1 Verify Folder Structure
Open File Explorer and confirm these folders exist:
```
C:\Mobius Reports\
  ├── Incoming\       (emails save attachments here)
  ├── Transformed\    (processed reports go here)
  └── Archive\        (processed inputs moved here)
```

### 4.2 Test with a Sample Email
1. Have someone send you a test NAV report email
2. Outlook should automatically:
   - Save the Excel attachment to `C:\Mobius Reports\Incoming\`
   - Open the Portfolio Transformer
   - Process the file
   - Save the result to `C:\Mobius Reports\Transformed\`

---

## Daily Workflow

Once set up, the system runs automatically:

1. **Receive Email** - NAV report arrives in your inbox
2. **Auto-Process** - Outlook detects attachment and triggers Excel
3. **Review Output** - Check `C:\Mobius Reports\Transformed\` for results
4. **Bloomberg Refresh** - Open the output file; Bloomberg formulas update automatically

---

## Troubleshooting

### "Macros have been disabled"
- Enable macros in Excel: `File > Options > Trust Center > Macro Settings`
- Enable macros in Outlook: Same path in Outlook options

### Bloomberg formulas show #NAME?
- Ensure Bloomberg Terminal is running
- Check Excel Add-ins for Bloomberg Excel Tools
- Run `DAPI <GO>` in Terminal to reinstall Office Tools

### Emails not being processed
- Verify the Outlook VBA code is in `ThisOutlookSession`
- Check that macros are enabled in Outlook
- Restart Outlook after making changes

### Japanese stocks showing wrong prices
- v5.1 includes automatic FX conversion for foreign tickers
- Prices display in USD regardless of local currency

---

## Version History

- **v5.1** - Added FX conversion for foreign tickers (JPY, GBP, EUR, etc.)
- **v5.0** - Added professional formatting (navy headers, zebra striping)
- **v4.0** - Initial Bloomberg integration

---

## Support

Contact Jacob for assistance.
