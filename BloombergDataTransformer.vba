' ====================================================================
' Bloomberg Portfolio Data Transformer - Version 5.1
' ====================================================================
' WHAT'S NEW IN V5.1:
'   - FX conversion for non-USD tickers (JP, LN, etc.)
'   - Japanese stocks now display prices in USD
'
' PREVIOUS (V5):
'   - Professional formatting matching input file styling
'   - Navy blue sub-headers with white text
'   - Alternating row colors (white/light gray zebra striping)
'   - Consistent font colors and borders
'
' PREVIOUS (V4):
'   - Integration with Outlook email monitor (automatic triggering)
'   - Reads YTD Fund Return from non-custom "Gain And Exposure" file (K94)
'   - Output saved to C:\Mobius Reports\Transformed\ folder
'   - Can be run manually or triggered by Outlook VBA
'
' DATA SOURCES:
'   1. Primary: Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP_[DATE].XLSX
'      - Contains: Position details, P&L, weights, # of shares
'   2. Performance: Gain And Exposure_MOBIUS EMERGING OPPORTUNITIES FUND LP_[DATE].XLSX
'      - Contains: K94 = YTD Fund Return (e.g., 3.74%)
'   3. Optional: DailyRor file for MTD/additional metrics
'
' USAGE (Manual):
'   1. Open the Custom NAV file
'   2. Run TransformBloombergData macro
'   3. Script finds matching non-custom file automatically
'
' USAGE (Automatic via Outlook):
'   1. Outlook monitor detects both emails
'   2. Saves attachments to C:\Mobius Reports\Incoming\
'   3. Calls SetDailyFilePath with the non-custom file path
'   4. Calls TransformBloombergData
'
' ====================================================================

Option Explicit

' ============================================
' CONFIGURATION
' ============================================
Private Const OUTPUT_FOLDER As String = "C:\Mobius Reports\Transformed\"
Private Const INCOMING_FOLDER As String = "C:\Mobius Reports\Incoming\"

' ============================================
' COLOR SCHEME (matching input file styling)
' ============================================
' Navy blue for sub-headers: #003366 = RGB(0, 51, 102)
Private Const COLOR_NAVY_BLUE As Long = 6697728    ' RGB(0, 51, 102) as Long
' White for sub-header text and alternating rows
Private Const COLOR_WHITE As Long = 16777215       ' RGB(255, 255, 255)
' Light gray for alternating rows: #F2F2F2
Private Const COLOR_LIGHT_GRAY As Long = 15921906  ' RGB(242, 242, 242)
' Dark gray for data text: #404040
Private Const COLOR_DARK_GRAY As Long = 4210752    ' RGB(64, 64, 64)
' Gray for header text: #595959
Private Const COLOR_HEADER_GRAY As Long = 5855577  ' RGB(89, 89, 89)

' ============================================
' GLOBAL VARIABLES
' ============================================
Dim stockPositions As Object      ' Dictionary: ticker -> shares
Dim ytdReturn As Double           ' YTD return from DailyRor
Dim mtdReturn As Double           ' MTD return from DailyRor
Dim totalEquity As Double         ' Total portfolio value
Dim navPerShare As Double         ' NAV per share
Dim performanceDataFound As Boolean
Dim dailyFilePath As String       ' Path to non-custom file (set by Outlook or found automatically)
Dim ytdFundReturn As Double       ' YTD Fund Return from K94 of non-custom file
Dim ytdFundReturnFound As Boolean ' Flag if K94 was found

' ============================================
' OUTLOOK INTEGRATION - Called by Outlook VBA
' ============================================
Public Sub SetDailyFilePath(filePath As String)
    ' Called by Outlook VBA to set the path to the non-custom file
    dailyFilePath = filePath
End Sub

' ====================================================================
' MAIN TRANSFORMATION PROCEDURE
' ====================================================================
Sub TransformBloombergData()
    Dim wsSource As Worksheet
    Dim wsStocks As Worksheet
    Dim wsOptions As Worksheet
    Dim wbOutput As Workbook
    Dim lastRow As Long
    Dim i As Long
    Dim productName As String
    Dim outputPath As String
    Dim todayDate As String
    Dim sourceFolder As String
    Dim reportDate As String

    On Error GoTo ErrorHandler

    ' Initialize
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set stockPositions = CreateObject("Scripting.Dictionary")
    performanceDataFound = False
    ytdFundReturnFound = False
    ytdReturn = 0
    mtdReturn = 0
    totalEquity = 0
    navPerShare = 0
    ytdFundReturn = 0

    ' Get the source worksheet (Custom file should be active)
    Set wsSource = ActiveSheet
    sourceFolder = ActiveWorkbook.Path & "\"

    ' Extract date from filename for finding matching files
    reportDate = ExtractDateFromFilename(ActiveWorkbook.Name)

    ' Try to find and read the non-custom file for K94 (YTD Fund Return)
    Call ReadYTDFundReturn(sourceFolder, reportDate)

    ' Try to find and read DailyRor file for additional metrics
    Call ReadDailyRorData(sourceFolder)

    ' Find last row with data
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' First pass: Build stock positions dictionary
    For i = 6 To lastRow
        productName = Trim(CStr(wsSource.Cells(i, 1).Value))
        If productName <> "" And productName <> "USD" And Not IsOption(productName) Then
            Dim ticker As String
            Dim shares As Variant
            ticker = Trim(CStr(wsSource.Cells(i, 2).Value))
            shares = wsSource.Cells(i, 12).Value

            Dim baseTicker As String
            baseTicker = ExtractBaseTicker(ticker)

            If baseTicker <> "" And IsNumeric(shares) Then
                stockPositions(baseTicker) = shares
            End If
        End If
    Next i

    ' Create new workbook for output
    Set wbOutput = Workbooks.Add
    wbOutput.Sheets(1).Name = "Stocks"
    Set wsStocks = wbOutput.Sheets("Stocks")

    Set wsOptions = wbOutput.Sheets.Add(After:=wsStocks)
    wsOptions.Name = "Options"

    ' Setup headers
    Call SetupStocksHeaders(wsStocks)
    Call SetupOptionsHeaders(wsOptions)

    ' Process rows
    Dim stockRow As Long
    Dim putRow As Long
    Dim callRow As Long
    Dim optionPutRows As Long

    stockRow = 4
    putRow = 5
    optionPutRows = 0

    ' Count puts first
    For i = 6 To lastRow
        productName = Trim(CStr(wsSource.Cells(i, 1).Value))
        If productName <> "" And productName <> "USD" Then
            If IsPutOption(productName) Then
                optionPutRows = optionPutRows + 1
            End If
        End If
    Next i

    ' Set starting row for calls
    callRow = putRow + optionPutRows + 2

    ' Add CALLS header
    wsOptions.Cells(callRow - 1, 2).Value = "CALLS"
    wsOptions.Cells(callRow - 1, 2).Font.Bold = True
    wsOptions.Cells(callRow - 1, 2).Font.Size = 12
    wsOptions.Range("J3:O4").Copy wsOptions.Range("J" & (callRow - 1))
    wsOptions.Range("B4:O4").Copy wsOptions.Range("B" & callRow)
    callRow = callRow + 1

    putRow = 5

    ' Process all rows
    For i = 6 To lastRow
        productName = Trim(CStr(wsSource.Cells(i, 1).Value))

        If productName <> "" And productName <> "USD" Then
            If IsOption(productName) Then
                If IsPutOption(productName) Then
                    Call ProcessOption(wsSource, i, wsOptions, putRow, "PUT")
                    putRow = putRow + 1
                ElseIf IsCallOption(productName) Then
                    Call ProcessOption(wsSource, i, wsOptions, callRow, "CALL")
                    callRow = callRow + 1
                End If
            Else
                Call ProcessStock(wsSource, i, wsStocks, stockRow)
                stockRow = stockRow + 1
            End If
        End If
    Next i

    ' Add cash positions
    Call AddCashPositions(wsSource, wsStocks, stockRow, lastRow)

    ' Add bottom totals section
    Call AddBottomTotals(wsStocks, stockRow + 4)

    ' Format sheets
    Call FormatStocksSheet(wsStocks, stockRow)
    Call FormatOptionsSheet(wsOptions, putRow, callRow)

    ' Determine output path
    ' Use OUTPUT_FOLDER if it exists, otherwise use source folder
    Dim savePath As String
    If Dir(OUTPUT_FOLDER, vbDirectory) <> "" Then
        savePath = OUTPUT_FOLDER
    Else
        savePath = sourceFolder
    End If

    todayDate = Format(Date, "DD MMMM YYYY")
    outputPath = savePath & "Transformed_Portfolio_" & todayDate & ".xlsx"

    If Dir(outputPath) <> "" Then
        outputPath = savePath & "Transformed_Portfolio_" & Format(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
    End If

    wbOutput.SaveAs outputPath

    ' Cleanup
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Summary message
    Dim msg As String
    msg = "Transformation complete!" & vbCrLf & vbCrLf
    msg = msg & "File saved: " & outputPath & vbCrLf & vbCrLf
    msg = msg & "Stocks processed: " & (stockRow - 4) & vbCrLf
    msg = msg & "Options processed: " & (putRow - 5 + callRow - (putRow + optionPutRows + 3)) & vbCrLf & vbCrLf

    msg = msg & "Performance Data:" & vbCrLf
    If ytdFundReturnFound Then
        msg = msg & "  YTD Fund Return (K94): " & Format(ytdFundReturn, "0.00%") & vbCrLf
    Else
        msg = msg & "  YTD Fund Return: Not found (non-custom file missing)" & vbCrLf
    End If

    If performanceDataFound Then
        msg = msg & "  MTD Return (DailyRor): " & Format(mtdReturn, "0.00%") & vbCrLf
        msg = msg & "  Total Equity: " & Format(totalEquity, "$#,##0")
    End If

    MsgBox msg, vbInformation, "Bloomberg Data Transformer v5.1"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical
End Sub

' ====================================================================
' READ YTD FUND RETURN FROM NON-CUSTOM FILE (K94)
' ====================================================================
Sub ReadYTDFundReturn(folderPath As String, reportDate As String)
    Dim fileName As String
    Dim filePath As String
    Dim wbDaily As Workbook
    Dim wsDaily As Worksheet
    Dim k94Value As Variant

    On Error GoTo NotFound

    ' If Outlook already set the path, use it
    If dailyFilePath <> "" And Dir(dailyFilePath) <> "" Then
        filePath = dailyFilePath
    Else
        ' Try to find the non-custom file in the same folder
        ' Pattern: Gain And Exposure_MOBIUS EMERGING OPPORTUNITIES FUND LP_MMDDYYYY.XLSX
        ' (Note: NO "Custom_" in the name)

        fileName = Dir(folderPath & "Gain And Exposure_MOBIUS EMERGING OPPORTUNITIES FUND LP_" & reportDate & ".XLSX")

        If fileName = "" Then
            ' Try incoming folder
            fileName = Dir(INCOMING_FOLDER & "Gain And Exposure_MOBIUS EMERGING OPPORTUNITIES FUND LP_" & reportDate & ".XLSX")
            If fileName <> "" Then
                filePath = INCOMING_FOLDER & fileName
            End If
        Else
            filePath = folderPath & fileName
        End If
    End If

    If filePath = "" Or Dir(filePath) = "" Then
        ytdFundReturnFound = False
        Exit Sub
    End If

    ' Open the non-custom file
    Set wbDaily = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set wsDaily = wbDaily.Sheets(1)

    ' Read K94 - YTD ROR
    k94Value = wsDaily.Cells(94, 11).Value  ' Column K = 11

    If IsNumeric(k94Value) Then
        ytdFundReturn = CDbl(k94Value)
        ytdFundReturnFound = True
    Else
        ytdFundReturnFound = False
    End If

    ' Also grab total equity from H94 if available
    If IsNumeric(wsDaily.Cells(94, 8).Value) Then
        totalEquity = CDbl(wsDaily.Cells(94, 8).Value)
    End If

    wbDaily.Close SaveChanges:=False

    Exit Sub

NotFound:
    ytdFundReturnFound = False
    On Error GoTo 0
End Sub

' ====================================================================
' EXTRACT DATE FROM FILENAME
' ====================================================================
Function ExtractDateFromFilename(fileName As String) As String
    ' Extract MMDDYYYY from filename like:
    ' "Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP_11262025.XLSX"

    Dim pos As Long
    Dim dateStr As String

    ' Find the underscore before the date
    pos = InStrRev(fileName, "_")

    If pos > 0 Then
        ' Extract 8 characters after the underscore (MMDDYYYY)
        dateStr = Mid(fileName, pos + 1, 8)

        If IsNumeric(dateStr) And Len(dateStr) = 8 Then
            ExtractDateFromFilename = dateStr
        Else
            ExtractDateFromFilename = ""
        End If
    Else
        ExtractDateFromFilename = ""
    End If
End Function

' ====================================================================
' READ DAILYROR FILE FOR ADDITIONAL PERFORMANCE DATA
' ====================================================================
Sub ReadDailyRorData(folderPath As String)
    Dim fileName As String
    Dim filePath As String
    Dim wbRor As Workbook
    Dim wsRor As Worksheet
    Dim i As Long

    On Error GoTo NotFound

    ' Look for DailyRor file (1003_DailyRor, not 1003_A)
    fileName = Dir(folderPath & "*_1003_DailyRor_*.xls")

    If fileName = "" Then
        ' Try incoming folder
        fileName = Dir(INCOMING_FOLDER & "*_1003_DailyRor_*.xls")
        If fileName <> "" Then
            filePath = INCOMING_FOLDER & fileName
        End If
    Else
        filePath = folderPath & fileName
    End If

    If fileName = "" Then
        ' Try alternative pattern
        fileName = Dir(folderPath & "*DailyRor*.xls")
        If fileName <> "" Then
            filePath = folderPath & fileName
        End If
    End If

    If fileName = "" Then
        performanceDataFound = False
        Exit Sub
    End If

    ' Make sure we don't get the _A version
    If InStr(fileName, "_A_DailyRor") > 0 Then
        fileName = Dir()
        Do While fileName <> ""
            If InStr(fileName, "_A_DailyRor") = 0 And InStr(fileName, "DailyRor") > 0 Then
                Exit Do
            End If
            fileName = Dir()
        Loop
    End If

    If fileName = "" Then
        performanceDataFound = False
        Exit Sub
    End If

    If filePath = "" Then filePath = folderPath & fileName

    ' Open DailyRor workbook
    Set wbRor = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set wsRor = wbRor.Sheets(1)

    ' Find data row (Row 13 or 14)
    For i = 13 To 20
        If IsNumeric(wsRor.Cells(i, 2).Value) And wsRor.Cells(i, 2).Value > 0 Then
            If totalEquity = 0 Then totalEquity = wsRor.Cells(i, 2).Value
            mtdReturn = wsRor.Cells(i, 4).Value
            If Not ytdFundReturnFound Then ytdReturn = wsRor.Cells(i, 6).Value
            navPerShare = wsRor.Cells(i, 9).Value
        End If
    Next i

    wbRor.Close SaveChanges:=False
    performanceDataFound = True

    Exit Sub

NotFound:
    performanceDataFound = False
    On Error GoTo 0
End Sub

' ====================================================================
' HELPER FUNCTIONS
' ====================================================================

Function IsOption(productName As String) As Boolean
    IsOption = (InStr(1, productName, " PUT ", vbTextCompare) > 0) Or _
               (InStr(1, productName, " CALL ", vbTextCompare) > 0)
End Function

Function IsPutOption(productName As String) As Boolean
    IsPutOption = InStr(1, productName, " PUT ", vbTextCompare) > 0
End Function

Function IsCallOption(productName As String) As Boolean
    IsCallOption = InStr(1, productName, " CALL ", vbTextCompare) > 0
End Function

Function ExtractBaseTicker(fullTicker As String) As String
    Dim spacePos As Long
    fullTicker = Trim(fullTicker)
    spacePos = InStr(fullTicker, " ")
    If spacePos > 0 Then
        ExtractBaseTicker = Left(fullTicker, spacePos - 1)
    Else
        ExtractBaseTicker = fullTicker
    End If
End Function

Function ExtractStrike(productName As String) As String
    Dim pos As Long
    Dim optionType As String

    If IsPutOption(productName) Then
        optionType = "PUT"
    Else
        optionType = "CALL"
    End If

    pos = InStr(1, productName, optionType, vbTextCompare)
    If pos > 0 Then
        ExtractStrike = Trim(Mid(productName, pos + Len(optionType)))
    Else
        ExtractStrike = ""
    End If
End Function

Function ExtractExpiry(productName As String) As String
    Dim parts() As String
    Dim i As Long

    parts = Split(productName, " ")

    For i = LBound(parts) To UBound(parts)
        If InStr(parts(i), "/") > 0 Then
            ExtractExpiry = parts(i)
            Exit Function
        End If
    Next i

    ExtractExpiry = ""
End Function

Function ExtractTickerFromOptionName(productName As String) As String
    Dim spacePos As Long
    spacePos = InStr(productName, " ")
    If spacePos > 0 Then
        ExtractTickerFromOptionName = Left(productName, spacePos - 1)
    Else
        ExtractTickerFromOptionName = productName
    End If
End Function

Function GetUnderlyingShares(optionTicker As String) As Variant
    Dim baseTicker As String
    baseTicker = ExtractTickerFromOptionName(optionTicker)

    If stockPositions.Exists(baseTicker) Then
        GetUnderlyingShares = stockPositions(baseTicker)
    Else
        GetUnderlyingShares = 0
    End If
End Function

' ====================================================================
' SETUP HEADERS
' ====================================================================

Sub SetupStocksHeaders(ws As Worksheet)
    ' Row 2: Column headers (bold gray text)
    ws.Cells(2, 2).Value = "Name"
    ws.Cells(2, 3).Value = "Ticker"
    ws.Cells(2, 4).Value = "Quantity"
    ws.Cells(2, 5).Value = "Unit Cost"
    ws.Cells(2, 6).Value = "Current Px"
    ws.Cells(2, 7).Value = "Total Cost"
    ws.Cells(2, 8).Value = "Mkt Value"
    ws.Cells(2, 9).Value = "% Diff (Cost)"
    ws.Cells(2, 10).Value = "Daily Chg %"
    ws.Cells(2, 11).Value = "P&L"
    ws.Cells(2, 12).Value = "Portfolio Wgt"
    ws.Cells(2, 13).Value = "Attribution"

    ' Row 3: Sub-headers (navy blue background, white text)
    ws.Cells(3, 5).Value = "USD"
    ws.Cells(3, 6).Value = "USD"
    ws.Cells(3, 7).Value = "USD"
    ws.Cells(3, 8).Value = "USD"
    ws.Cells(3, 11).Value = "YTD"
    ws.Cells(3, 12).Value = "%"
    ws.Cells(3, 13).Value = "%"

    ' Format header row (row 2) - bold gray text
    With ws.Range("B2:M2")
        .Font.Bold = True
        .Font.Color = COLOR_HEADER_GRAY
        .HorizontalAlignment = xlCenter
    End With

    ' Format sub-header row (row 3) - navy blue background, white text
    With ws.Range("B3:M3")
        .Font.Bold = True
        .Font.Color = COLOR_WHITE
        .Interior.Color = COLOR_NAVY_BLUE
        .HorizontalAlignment = xlCenter
    End With
End Sub

Sub SetupOptionsHeaders(ws As Worksheet)
    ' Row 2: Section title "PUTS"
    ws.Cells(2, 2).Value = "PUTS"
    ws.Cells(2, 2).Font.Bold = True
    ws.Cells(2, 2).Font.Size = 12
    ws.Cells(2, 2).Font.Color = COLOR_HEADER_GRAY

    ' Row 3: Column group headers (navy blue background, white text)
    ws.Cells(3, 10).Value = "Unit Cost"
    ws.Cells(3, 12).Value = "Total Cost"
    ws.Cells(3, 13).Value = "Current Px"
    ws.Cells(3, 14).Value = "Mkt Value"
    ws.Cells(3, 15).Value = "P&L"

    ' Row 4: Sub-headers with units
    ws.Cells(4, 2).Value = "Name"
    ws.Cells(4, 3).Value = "Quantity"
    ws.Cells(4, 4).Value = "Underlying Qty"
    ws.Cells(4, 5).Value = "% Hedged"
    ws.Cells(4, 6).Value = "Strike Px"
    ws.Cells(4, 7).Value = "Underlying Px"
    ws.Cells(4, 8).Value = "% Moneyness"
    ws.Cells(4, 9).Value = "Expiry"
    ws.Cells(4, 10).Value = "USD"
    ws.Cells(4, 11).Value = "% Yield"
    ws.Cells(4, 12).Value = "USD"
    ws.Cells(4, 13).Value = "USD"
    ws.Cells(4, 14).Value = "USD"
    ws.Cells(4, 15).Value = "YTD"

    ' Format row 3 - navy blue background, white text for column group headers
    With ws.Range("J3:O3")
        .Font.Bold = True
        .Font.Color = COLOR_WHITE
        .Interior.Color = COLOR_NAVY_BLUE
        .HorizontalAlignment = xlCenter
    End With

    ' Format row 4 - navy blue background, white text for sub-headers
    With ws.Range("B4:O4")
        .Font.Bold = True
        .Font.Color = COLOR_WHITE
        .Interior.Color = COLOR_NAVY_BLUE
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ====================================================================
' PROCESS DATA ROWS
' ====================================================================

Sub ProcessStock(wsSource As Worksheet, sourceRow As Long, wsTarget As Worksheet, targetRow As Long)
    wsTarget.Cells(targetRow, 2).Value = wsSource.Cells(sourceRow, 1).Value  ' Name
    wsTarget.Cells(targetRow, 3).Value = wsSource.Cells(sourceRow, 2).Value  ' Ticker
    wsTarget.Cells(targetRow, 4).Value = wsSource.Cells(sourceRow, 12).Value ' Quantity
    wsTarget.Cells(targetRow, 5).Value = wsSource.Cells(sourceRow, 5).Value  ' Unit Cost

    Dim ticker As String
    Dim fxCurrency As String
    ticker = Trim(CStr(wsSource.Cells(sourceRow, 2).Value))

    If ticker <> "" And InStr(ticker, " ") > 0 Then
        ' Check if foreign ticker needs FX conversion
        fxCurrency = GetFXCurrency(ticker)

        If fxCurrency <> "" Then
            ' Foreign ticker: multiply price by FX rate to convert to USD
            wsTarget.Cells(targetRow, 6).FormulaArray = "=BDP(""" & ticker & " Equity"",""PX_LAST"")*BDP(""" & fxCurrency & "USD Curncy"",""PX_LAST"")"
        Else
            ' USD ticker: use price directly
            wsTarget.Cells(targetRow, 6).FormulaArray = "=BDP(""" & ticker & " Equity"",""PX_LAST"")"
        End If
    Else
        wsTarget.Cells(targetRow, 6).Value = wsSource.Cells(sourceRow, 6).Value
    End If

    wsTarget.Cells(targetRow, 7).Value = wsSource.Cells(sourceRow, 9).Value  ' Total Cost
    wsTarget.Cells(targetRow, 8).Formula = "=F" & targetRow & "*D" & targetRow
    wsTarget.Cells(targetRow, 9).Formula = "=(F" & targetRow & "-E" & targetRow & ")/E" & targetRow
    wsTarget.Cells(targetRow, 10).Value = wsSource.Cells(sourceRow, 7).Value
    wsTarget.Cells(targetRow, 11).Value = wsSource.Cells(sourceRow, 11).Value ' P&L
    wsTarget.Cells(targetRow, 12).Value = wsSource.Cells(sourceRow, 4).Value  ' Portfolio Wgt
    wsTarget.Cells(targetRow, 13).Value = wsSource.Cells(sourceRow, 8).Value  ' Attribution
End Sub

' ====================================================================
' CURRENCY HELPER - Returns FX currency code for non-USD tickers
' ====================================================================
Function GetFXCurrency(ticker As String) As String
    Dim suffix As String
    Dim spacePos As Long

    ' Extract the exchange suffix (e.g., "JP" from "2644 JP")
    spacePos = InStr(ticker, " ")
    If spacePos > 0 Then
        suffix = UCase(Trim(Mid(ticker, spacePos + 1)))
    Else
        GetFXCurrency = ""
        Exit Function
    End If

    ' Map exchange suffix to currency
    Select Case suffix
        Case "JP"   ' Japan
            GetFXCurrency = "JPY"
        Case "LN"   ' London
            GetFXCurrency = "GBP"
        Case "GY", "GR"  ' Germany
            GetFXCurrency = "EUR"
        Case "FP"   ' France
            GetFXCurrency = "EUR"
        Case "IM"   ' Italy
            GetFXCurrency = "EUR"
        Case "SM"   ' Spain
            GetFXCurrency = "EUR"
        Case "NA"   ' Netherlands
            GetFXCurrency = "EUR"
        Case "AV"   ' Austria
            GetFXCurrency = "EUR"
        Case "SW"   ' Switzerland
            GetFXCurrency = "CHF"
        Case "CN"   ' Canada
            GetFXCurrency = "CAD"
        Case "AU"   ' Australia
            GetFXCurrency = "AUD"
        Case "HK"   ' Hong Kong
            GetFXCurrency = "HKD"
        Case "SP"   ' Singapore
            GetFXCurrency = "SGD"
        Case "KS"   ' Korea
            GetFXCurrency = "KRW"
        Case "TT"   ' Taiwan
            GetFXCurrency = "TWD"
        Case "US", "UN", "UA", "UQ", "UW"  ' US exchanges
            GetFXCurrency = ""  ' No conversion needed
        Case Else
            GetFXCurrency = ""  ' Default: assume USD or unknown
    End Select
End Function

Sub ProcessOption(wsSource As Worksheet, sourceRow As Long, wsTarget As Worksheet, targetRow As Long, optionType As String)
    Dim productName As String
    Dim occTicker As String
    Dim quantity As Variant
    Dim underlyingQty As Variant
    Dim strike As String
    Dim expiry As String
    Dim baseTicker As String

    productName = wsSource.Cells(sourceRow, 1).Value
    occTicker = Trim(CStr(wsSource.Cells(sourceRow, 2).Value))
    quantity = wsSource.Cells(sourceRow, 12).Value
    strike = ExtractStrike(productName)
    expiry = ExtractExpiry(productName)
    baseTicker = ExtractTickerFromOptionName(productName)

    underlyingQty = GetUnderlyingShares(baseTicker)

    wsTarget.Cells(targetRow, 2).Value = productName
    wsTarget.Cells(targetRow, 3).Value = quantity
    wsTarget.Cells(targetRow, 4).Value = underlyingQty

    If underlyingQty <> 0 Then
        If optionType = "PUT" Then
            wsTarget.Cells(targetRow, 5).Formula = "=$C" & targetRow & "*100/$D" & targetRow
        Else
            wsTarget.Cells(targetRow, 5).Formula = "=-$C" & targetRow & "*100/$D" & targetRow
        End If
    Else
        wsTarget.Cells(targetRow, 5).Value = "N/A"
    End If

    If IsNumeric(strike) Then
        wsTarget.Cells(targetRow, 6).Value = CDbl(strike)
    Else
        wsTarget.Cells(targetRow, 6).Value = strike
    End If

    If occTicker <> "" Then
        wsTarget.Cells(targetRow, 7).FormulaArray = "=BDP(""" & occTicker & " Equity"",""OPT_UNDL_PX"")"
    End If

    wsTarget.Cells(targetRow, 8).Formula = "=(G" & targetRow & "-F" & targetRow & ")/F" & targetRow
    wsTarget.Cells(targetRow, 9).Value = expiry
    wsTarget.Cells(targetRow, 10).Value = wsSource.Cells(sourceRow, 5).Value
    wsTarget.Cells(targetRow, 11).Formula = "=J" & targetRow & "/G" & targetRow
    wsTarget.Cells(targetRow, 12).Value = wsSource.Cells(sourceRow, 9).Value
    wsTarget.Cells(targetRow, 13).Value = wsSource.Cells(sourceRow, 6).Value
    wsTarget.Cells(targetRow, 14).Value = wsSource.Cells(sourceRow, 10).Value
    wsTarget.Cells(targetRow, 15).Value = wsSource.Cells(sourceRow, 11).Value
End Sub

Sub AddCashPositions(wsSource As Worksheet, wsTarget As Worksheet, startRow As Long, lastRow As Long)
    Dim i As Long
    Dim targetRow As Long
    Dim productName As String

    targetRow = startRow + 2

    For i = 6 To lastRow
        productName = Trim(CStr(wsSource.Cells(i, 1).Value))
        If productName = "USD" Or productName = "JPY" Or productName = "CAD" Or productName = "EUR" Or productName = "GBP" Then
            wsTarget.Cells(targetRow, 2).Value = productName & " "
            wsTarget.Cells(targetRow, 4).Value = wsSource.Cells(i, 12).Value
            targetRow = targetRow + 1
        End If
    Next i
End Sub

' ====================================================================
' ADD BOTTOM TOTALS SECTION - UPDATED FOR V4
' ====================================================================

Sub AddBottomTotals(ws As Worksheet, startRow As Long)
    Dim r As Long
    r = startRow + 1

    ' Section header
    ws.Cells(r, 2).Value = "FUND PERFORMANCE SUMMARY"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Size = 12
    r = r + 2

    ' Total Portfolio Value
    ws.Cells(r, 2).Value = "Total Portfolio Value:"
    ws.Cells(r, 2).Font.Bold = True
    If totalEquity > 0 Then
        ws.Cells(r, 4).Value = totalEquity
        ws.Cells(r, 4).NumberFormat = "$#,##0"
    Else
        ws.Cells(r, 4).Value = "(Not available)"
    End If
    r = r + 1

    ' NAV Per Share
    ws.Cells(r, 2).Value = "NAV Per Share:"
    ws.Cells(r, 2).Font.Bold = True
    If navPerShare > 0 Then
        ws.Cells(r, 4).Value = navPerShare
        ws.Cells(r, 4).NumberFormat = "$#,##0.00"
    Else
        ws.Cells(r, 4).Value = "(Not available)"
    End If
    r = r + 1

    ' Inception Date
    ws.Cells(r, 2).Value = "Fund Inception Date:"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 4).Value = "March 2025"
    r = r + 1

    ' YTD Fund Return (from K94 - PRIMARY SOURCE)
    ws.Cells(r, 2).Value = "YTD Fund Return:"
    ws.Cells(r, 2).Font.Bold = True
    If ytdFundReturnFound Then
        ws.Cells(r, 4).Value = ytdFundReturn
        ws.Cells(r, 4).NumberFormat = "0.00%"
        ws.Cells(r, 5).Value = "(from Gain & Exposure report)"
        ws.Cells(r, 5).Font.Italic = True
        ws.Cells(r, 5).Font.Color = RGB(128, 128, 128)
    ElseIf performanceDataFound And ytdReturn <> 0 Then
        ws.Cells(r, 4).Value = ytdReturn
        ws.Cells(r, 4).NumberFormat = "0.00%"
        ws.Cells(r, 5).Value = "(from DailyRor)"
        ws.Cells(r, 5).Font.Italic = True
        ws.Cells(r, 5).Font.Color = RGB(128, 128, 128)
    Else
        ws.Cells(r, 4).Value = "(Not available)"
    End If
    r = r + 1

    ' MTD Net Return
    ws.Cells(r, 2).Value = "MTD Net Return:"
    ws.Cells(r, 2).Font.Bold = True
    If performanceDataFound And mtdReturn <> 0 Then
        ws.Cells(r, 4).Value = mtdReturn
        ws.Cells(r, 4).NumberFormat = "0.00%"
    Else
        ws.Cells(r, 4).Value = "(Not available)"
    End If
    r = r + 1

    ' Data Source Note
    r = r + 1
    ws.Cells(r, 2).Value = "Report generated: " & Format(Now, "MMMM D, YYYY h:mm AM/PM")
    ws.Cells(r, 2).Font.Italic = True
    ws.Cells(r, 2).Font.Color = RGB(128, 128, 128)
End Sub

' ====================================================================
' FORMATTING
' ====================================================================

Sub FormatStocksSheet(ws As Worksheet, lastRow As Long)
    Dim i As Long

    If lastRow > 4 Then
        ' Number formats
        ws.Range("D4:D" & lastRow).NumberFormat = "#,##0"
        ws.Range("E4:H" & lastRow).NumberFormat = "$#,##0.00"
        ws.Range("I4:I" & lastRow).NumberFormat = "0.00%"
        ws.Range("J4:J" & lastRow).NumberFormat = "0.00%"
        ws.Range("K4:K" & lastRow).NumberFormat = "#,##0"
        ws.Range("L4:M" & lastRow).NumberFormat = "0.00%"

        ' Apply alternating row colors (zebra striping) and font color
        For i = 4 To lastRow
            With ws.Range("B" & i & ":M" & i)
                .Font.Color = COLOR_DARK_GRAY
                If (i Mod 2) = 0 Then
                    .Interior.Color = COLOR_WHITE
                Else
                    .Interior.Color = COLOR_LIGHT_GRAY
                End If
            End With
        Next i
    End If

    ws.Columns("B:M").AutoFit

    ' Add borders
    If lastRow > 4 Then
        With ws.Range("B2:M" & lastRow).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(200, 200, 200)
        End With
    End If
End Sub

Sub FormatOptionsSheet(ws As Worksheet, lastPutRow As Long, lastCallRow As Long)
    Dim lastRow As Long
    Dim i As Long
    Dim callHeaderRow As Long

    lastRow = Application.Max(lastPutRow, lastCallRow)

    ' Number formats
    ws.Range("C:C").NumberFormat = "#,##0"
    ws.Range("D:D").NumberFormat = "#,##0"
    ws.Range("E:E").NumberFormat = "0.00%"
    ws.Range("F:F").NumberFormat = "$#,##0.00"
    ws.Range("G:G").NumberFormat = "$#,##0.00"
    ws.Range("H:H").NumberFormat = "0.00%"
    ws.Range("J:J").NumberFormat = "$#,##0.00"
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("L:N").NumberFormat = "$#,##0"
    ws.Range("O:O").NumberFormat = "#,##0"

    ' Find CALLS header row (it's 2 rows after the last PUT)
    callHeaderRow = lastPutRow + 2

    ' Apply alternating row colors for PUTS section (starting row 5)
    For i = 5 To lastPutRow
        With ws.Range("B" & i & ":O" & i)
            .Font.Color = COLOR_DARK_GRAY
            If (i Mod 2) = 1 Then
                .Interior.Color = COLOR_WHITE
            Else
                .Interior.Color = COLOR_LIGHT_GRAY
            End If
        End With
    Next i

    ' Format CALLS header row with navy blue
    If callHeaderRow > 5 Then
        ' CALLS title
        ws.Cells(callHeaderRow - 1, 2).Font.Color = COLOR_HEADER_GRAY

        ' CALLS sub-header row - navy blue background
        With ws.Range("B" & callHeaderRow & ":O" & callHeaderRow)
            .Font.Bold = True
            .Font.Color = COLOR_WHITE
            .Interior.Color = COLOR_NAVY_BLUE
            .HorizontalAlignment = xlCenter
        End With

        ' Format column group headers for CALLS section
        With ws.Range("J" & (callHeaderRow - 1) & ":O" & (callHeaderRow - 1))
            .Font.Bold = True
            .Font.Color = COLOR_WHITE
            .Interior.Color = COLOR_NAVY_BLUE
            .HorizontalAlignment = xlCenter
        End With

        ' Apply alternating row colors for CALLS section
        For i = callHeaderRow + 1 To lastCallRow
            With ws.Range("B" & i & ":O" & i)
                .Font.Color = COLOR_DARK_GRAY
                If (i Mod 2) = 0 Then
                    .Interior.Color = COLOR_WHITE
                Else
                    .Interior.Color = COLOR_LIGHT_GRAY
                End If
            End With
        Next i
    End If

    ws.Columns("B:O").AutoFit

    ' Add borders
    If lastRow > 4 Then
        With ws.Range("B2:O" & lastRow).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(200, 200, 200)
        End With
    End If
End Sub
