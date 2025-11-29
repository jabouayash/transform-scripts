' ====================================================================
' Bloomberg Portfolio Data Transformer - Version 3
' ====================================================================
' WHAT'S NEW IN V3:
'   - Added "% Diff (Cost)" column: (Current Price - Unit Cost) / Unit Cost
'   - Added "Daily Chg %" column: From NAV source Column G
'   - Added bottom totals section with YTD/MTD performance from DailyRor file
'   - Multi-file support: Reads performance data from DailyRor report
'
' DATA SOURCES:
'   - Primary: Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP_[DATE].XLSX
'   - Performance: Mobius Emerging Opportunities Fund Ltd._United Overseas Bank Limited_1003_DailyRor_[DATE].xls
'     (Note: Using 1003_DailyRor, NOT 1003_A_DailyRor - the A version is a sub-account)
'
' USAGE:
'   1. Place both files in the same folder
'   2. Open the NAV file (Gain And Exposure_Custom_...)
'   3. Run TransformBloombergData macro
'   4. Script automatically finds and reads the DailyRor file
'
' COLUMN LAYOUT (Stocks Tab):
'   B: Name | C: Ticker | D: Quantity | E: Unit Cost | F: Current Px
'   G: Total Cost | H: Mkt Value | I: % Diff (Cost) | J: Daily Chg %
'   K: P&L | L: Portfolio Wgt | M: Attribution
'
' BOTTOM SECTION (Stocks Tab):
'   - Total Portfolio Value
'   - Fund Inception Date (March 2025)
'   - YTD Net Return (from DailyRor)
'   - MTD Net Return (from DailyRor)
' ====================================================================

Option Explicit

' Global variables
Dim stockPositions As Object ' Dictionary: ticker -> shares
Dim ytdReturn As Double      ' YTD return from DailyRor
Dim mtdReturn As Double      ' MTD return from DailyRor
Dim totalEquity As Double    ' Total portfolio value from DailyRor
Dim navPerShare As Double    ' NAV per share from DailyRor
Dim performanceDataFound As Boolean ' Flag if DailyRor was found

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

    On Error GoTo ErrorHandler

    ' Initialize
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set stockPositions = CreateObject("Scripting.Dictionary")
    performanceDataFound = False
    ytdReturn = 0
    mtdReturn = 0
    totalEquity = 0
    navPerShare = 0

    ' Get the source worksheet
    Set wsSource = ActiveSheet
    sourceFolder = ActiveWorkbook.Path & "\"

    ' Try to find and read DailyRor file
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

    ' Save output
    todayDate = Format(Date, "DD MMMM YYYY")
    outputPath = sourceFolder & "Transformed_Portfolio_" & todayDate & ".xlsx"

    If Dir(outputPath) <> "" Then
        outputPath = sourceFolder & "Transformed_Portfolio_" & Format(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
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

    If performanceDataFound Then
        msg = msg & "Performance data loaded from DailyRor file:" & vbCrLf
        msg = msg & "  YTD Return: " & Format(ytdReturn, "0.00%") & vbCrLf
        msg = msg & "  MTD Return: " & Format(mtdReturn, "0.00%") & vbCrLf
        msg = msg & "  Total Equity: " & Format(totalEquity, "$#,##0")
    Else
        msg = msg & "NOTE: DailyRor file not found in folder." & vbCrLf
        msg = msg & "Performance data not included." & vbCrLf
        msg = msg & "Expected file: *_1003_DailyRor_*.xls"
    End If

    MsgBox msg, vbInformation, "Bloomberg Data Transformer v3"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical
End Sub

' ====================================================================
' READ DAILYROR FILE FOR PERFORMANCE DATA
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
        ' Try alternative pattern
        fileName = Dir(folderPath & "*DailyRor*.xls")
    End If

    If fileName = "" Then
        performanceDataFound = False
        Exit Sub
    End If

    ' Make sure we don't get the _A version
    If InStr(fileName, "_A_DailyRor") > 0 Then
        ' Skip _A version, look for next match
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

    filePath = folderPath & fileName

    ' Open DailyRor workbook
    Set wbRor = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set wsRor = wbRor.Sheets(1)

    ' Find data row (Row 13 or 14 based on our analysis)
    ' Structure: Row 12 = Headers, Row 13/14 = Data
    ' Columns: A=Date, B=Ending Equity, D=MTD Return, F=YTD Return, I=NAV Per Share

    For i = 13 To 20
        If IsNumeric(wsRor.Cells(i, 2).Value) And wsRor.Cells(i, 2).Value > 0 Then
            ' Found data row - get the LAST row with data (most recent)
            totalEquity = wsRor.Cells(i, 2).Value
            mtdReturn = wsRor.Cells(i, 4).Value
            ytdReturn = wsRor.Cells(i, 6).Value
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
' SETUP HEADERS - UPDATED FOR V3
' ====================================================================

Sub SetupStocksHeaders(ws As Worksheet)
    ' Row 2 - Main headers
    ws.Cells(2, 2).Value = "Name"
    ws.Cells(2, 3).Value = "Ticker"
    ws.Cells(2, 4).Value = "Quantity"
    ws.Cells(2, 5).Value = "Unit Cost"
    ws.Cells(2, 6).Value = "Current Px"
    ws.Cells(2, 7).Value = "Total Cost"
    ws.Cells(2, 8).Value = "Mkt Value"
    ws.Cells(2, 9).Value = "% Diff (Cost)"    ' NEW
    ws.Cells(2, 10).Value = "Daily Chg %"     ' NEW
    ws.Cells(2, 11).Value = "P&L"
    ws.Cells(2, 12).Value = "Portfolio Wgt"
    ws.Cells(2, 13).Value = "Attribution"

    ' Row 3 - Sub headers
    ws.Cells(3, 5).Value = "USD"
    ws.Cells(3, 6).Value = "USD"
    ws.Cells(3, 7).Value = "USD"
    ws.Cells(3, 8).Value = "USD"
    ws.Cells(3, 9).Value = ""                 ' NEW
    ws.Cells(3, 10).Value = ""                ' NEW
    ws.Cells(3, 11).Value = "YTD"
    ws.Cells(3, 12).Value = "%"
    ws.Cells(3, 13).Value = "%"

    With ws.Range("B2:M3")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

Sub SetupOptionsHeaders(ws As Worksheet)
    ' PUTS Section
    ws.Cells(2, 2).Value = "PUTS"
    ws.Cells(2, 2).Font.Bold = True
    ws.Cells(2, 2).Font.Size = 12

    ws.Cells(3, 10).Value = "Unit Cost"
    ws.Cells(3, 12).Value = "Total Cost"
    ws.Cells(3, 13).Value = "Current Px"
    ws.Cells(3, 14).Value = "Mkt Value"
    ws.Cells(3, 15).Value = "P&L"

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

    With ws.Range("B3:O4")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ====================================================================
' PROCESS DATA ROWS - UPDATED FOR V3
' ====================================================================

Sub ProcessStock(wsSource As Worksheet, sourceRow As Long, wsTarget As Worksheet, targetRow As Long)
    ' Column mapping - UPDATED FOR V3:
    ' A: Product Name -> B: Name
    ' B: Ticker -> C: Ticker
    ' L: # of Shares -> D: Quantity
    ' E: Unit Cost USD -> E: Unit Cost USD
    ' F: Today USD -> F: Current Px USD (Bloomberg)
    ' I: Total Cost USD -> G: Total Cost USD
    ' Calculated -> H: Mkt Value (=F*D)
    ' Calculated -> I: % Diff (Cost) = (F-E)/E  *** NEW ***
    ' G: % Daily Gain/Loss -> J: Daily Chg %   *** NEW ***
    ' K: Total Net P&L YTD -> K: P&L
    ' D: Portfolio Weight % -> L: Portfolio Wgt
    ' H: Contribution -> M: Attribution

    wsTarget.Cells(targetRow, 2).Value = wsSource.Cells(sourceRow, 1).Value  ' Name
    wsTarget.Cells(targetRow, 3).Value = wsSource.Cells(sourceRow, 2).Value  ' Ticker
    wsTarget.Cells(targetRow, 4).Value = wsSource.Cells(sourceRow, 12).Value ' Quantity
    wsTarget.Cells(targetRow, 5).Value = wsSource.Cells(sourceRow, 5).Value  ' Unit Cost

    ' Current Price - Bloomberg
    Dim ticker As String
    ticker = Trim(CStr(wsSource.Cells(sourceRow, 2).Value))

    If ticker <> "" And InStr(ticker, " ") > 0 Then
        wsTarget.Cells(targetRow, 6).FormulaArray = "=BDP(""" & ticker & " Equity"",""PX_LAST"")"
    Else
        wsTarget.Cells(targetRow, 6).Value = wsSource.Cells(sourceRow, 6).Value
    End If

    wsTarget.Cells(targetRow, 7).Value = wsSource.Cells(sourceRow, 9).Value  ' Total Cost

    ' Market Value formula
    wsTarget.Cells(targetRow, 8).Formula = "=F" & targetRow & "*D" & targetRow

    ' *** NEW: % Diff (Cost) = (Current Price - Unit Cost) / Unit Cost ***
    wsTarget.Cells(targetRow, 9).Formula = "=(F" & targetRow & "-E" & targetRow & ")/E" & targetRow

    ' *** NEW: Daily Change % from source Column G ***
    wsTarget.Cells(targetRow, 10).Value = wsSource.Cells(sourceRow, 7).Value

    wsTarget.Cells(targetRow, 11).Value = wsSource.Cells(sourceRow, 11).Value ' P&L
    wsTarget.Cells(targetRow, 12).Value = wsSource.Cells(sourceRow, 4).Value  ' Portfolio Wgt
    wsTarget.Cells(targetRow, 13).Value = wsSource.Cells(sourceRow, 8).Value  ' Attribution
End Sub

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
' ADD BOTTOM TOTALS SECTION - NEW IN V3
' ====================================================================

Sub AddBottomTotals(ws As Worksheet, startRow As Long)
    Dim r As Long
    r = startRow

    ' Add some spacing
    r = r + 1

    ' Section header
    ws.Cells(r, 2).Value = "FUND PERFORMANCE SUMMARY"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Size = 12
    r = r + 2

    ' Total Portfolio Value
    ws.Cells(r, 2).Value = "Total Portfolio Value:"
    ws.Cells(r, 2).Font.Bold = True
    If performanceDataFound And totalEquity > 0 Then
        ws.Cells(r, 4).Value = totalEquity
        ws.Cells(r, 4).NumberFormat = "$#,##0"
    Else
        ws.Cells(r, 4).Value = "(See DailyRor report)"
    End If
    r = r + 1

    ' NAV Per Share
    ws.Cells(r, 2).Value = "NAV Per Share:"
    ws.Cells(r, 2).Font.Bold = True
    If performanceDataFound And navPerShare > 0 Then
        ws.Cells(r, 4).Value = navPerShare
        ws.Cells(r, 4).NumberFormat = "$#,##0.00"
    Else
        ws.Cells(r, 4).Value = "(See DailyRor report)"
    End If
    r = r + 1

    ' Inception Date
    ws.Cells(r, 2).Value = "Fund Inception Date:"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 4).Value = "March 2025"
    r = r + 1

    ' YTD Net Return
    ws.Cells(r, 2).Value = "YTD Net Return:"
    ws.Cells(r, 2).Font.Bold = True
    If performanceDataFound Then
        ws.Cells(r, 4).Value = ytdReturn
        ws.Cells(r, 4).NumberFormat = "0.00%"
    Else
        ws.Cells(r, 4).Value = "(DailyRor file not found)"
    End If
    r = r + 1

    ' MTD Net Return
    ws.Cells(r, 2).Value = "MTD Net Return:"
    ws.Cells(r, 2).Font.Bold = True
    If performanceDataFound Then
        ws.Cells(r, 4).Value = mtdReturn
        ws.Cells(r, 4).NumberFormat = "0.00%"
    Else
        ws.Cells(r, 4).Value = "(DailyRor file not found)"
    End If
    r = r + 1

    ' Data Source Note
    r = r + 1
    ws.Cells(r, 2).Value = "Performance data source:"
    ws.Cells(r, 3).Value = "DailyRor_1003 report (not _A version)"
    ws.Cells(r, 2).Font.Italic = True
    ws.Cells(r, 3).Font.Italic = True
End Sub

' ====================================================================
' FORMATTING - UPDATED FOR V3
' ====================================================================

Sub FormatStocksSheet(ws As Worksheet, lastRow As Long)
    If lastRow > 4 Then
        ws.Range("D4:D" & lastRow).NumberFormat = "#,##0"
        ws.Range("E4:H" & lastRow).NumberFormat = "$#,##0.00"
        ws.Range("I4:I" & lastRow).NumberFormat = "0.00%"      ' % Diff (Cost)
        ws.Range("J4:J" & lastRow).NumberFormat = "0.00%"      ' Daily Chg %
        ws.Range("K4:K" & lastRow).NumberFormat = "#,##0"
        ws.Range("L4:M" & lastRow).NumberFormat = "0.00%"
    End If

    ws.Columns("B:M").AutoFit

    If lastRow > 4 Then
        ws.Range("B2:M" & lastRow).Borders.LineStyle = xlContinuous
    End If
End Sub

Sub FormatOptionsSheet(ws As Worksheet, lastPutRow As Long, lastCallRow As Long)
    Dim lastRow As Long
    lastRow = Application.Max(lastPutRow, lastCallRow)

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

    ws.Columns("B:O").AutoFit

    If lastRow > 4 Then
        ws.Range("B2:O" & lastRow).Borders.LineStyle = xlContinuous
    End If
End Sub
