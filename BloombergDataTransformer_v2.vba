' ====================================================================
' Bloomberg Portfolio Data Transformer - Production Version
' ====================================================================
' Transforms NAV calculation files into structured format with separate
' tabs for Stocks and Options, enhanced with Bloomberg API data
'
' Source: Daily NAV calculation file
' Output: Excel file with Stocks and Options tabs
' Bloomberg API: Live prices, option Greeks, implied volatility
' ====================================================================

Option Explicit

' Global variables to store stock positions for option matching
Dim stockPositions As Object ' Dictionary: ticker -> shares

' Main transformation procedure - attach this to a button
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

    On Error GoTo ErrorHandler

    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Initialize stock positions dictionary
    Set stockPositions = CreateObject("Scripting.Dictionary")

    ' Get the source worksheet (assuming active sheet)
    Set wsSource = ActiveSheet

    ' Find last row with data
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' First pass: Build stock positions dictionary for option matching
    For i = 6 To lastRow
        productName = Trim(CStr(wsSource.Cells(i, 1).Value))
        If productName <> "" And productName <> "USD" And Not IsOption(productName) Then
            Dim ticker As String
            Dim shares As Variant
            ticker = Trim(CStr(wsSource.Cells(i, 2).Value))
            shares = wsSource.Cells(i, 12).Value ' Column L: # of Shares

            ' Extract base ticker (e.g., "META US" -> "META")
            Dim baseTicker As String
            baseTicker = ExtractBaseTicker(ticker)

            If baseTicker <> "" And IsNumeric(shares) Then
                stockPositions(baseTicker) = shares
            End If
        End If
    Next i

    ' Create new workbook for output
    Set wbOutput = Workbooks.Add

    ' Rename default sheet to Stocks
    wbOutput.Sheets(1).Name = "Stocks"
    Set wsStocks = wbOutput.Sheets("Stocks")

    ' Add Options sheet
    Set wsOptions = wbOutput.Sheets.Add(After:=wsStocks)
    wsOptions.Name = "Options"

    ' Setup headers for Stocks sheet
    Call SetupStocksHeaders(wsStocks)

    ' Setup headers for Options sheet
    Call SetupOptionsHeaders(wsOptions)

    ' Process each row from source
    Dim stockRow As Long
    Dim putRow As Long
    Dim callRow As Long

    stockRow = 4  ' Starting row for stocks (after headers)
    putRow = 5    ' Starting row for puts (after headers)

    ' First, process all rows to count puts
    Dim optionPutRows As Long
    optionPutRows = 0

    For i = 6 To lastRow
        productName = Trim(CStr(wsSource.Cells(i, 1).Value))
        If productName <> "" And productName <> "USD" Then
            If IsPutOption(productName) Then
                optionPutRows = optionPutRows + 1
            End If
        End If
    Next i

    ' Set starting row for calls (after all puts + headers + blank row)
    callRow = putRow + optionPutRows + 2

    ' Add CALLS header
    wsOptions.Cells(callRow - 1, 2).Value = "CALLS"
    wsOptions.Cells(callRow - 1, 2).Font.Bold = True
    wsOptions.Cells(callRow - 1, 2).Font.Size = 12

    ' Copy headers for CALLS section
    wsOptions.Range("J3:O4").Copy wsOptions.Range("J" & (callRow - 1))
    wsOptions.Range("B4:O4").Copy wsOptions.Range("B" & callRow)
    callRow = callRow + 1

    ' Reset put counter
    putRow = 5

    ' Second pass: Process all rows and populate output
    For i = 6 To lastRow
        productName = Trim(CStr(wsSource.Cells(i, 1).Value))

        ' Skip empty rows and USD cash row (we'll add cash at the end)
        If productName <> "" And productName <> "USD" Then
            If IsOption(productName) Then
                ' Process as option
                If IsPutOption(productName) Then
                    Call ProcessOption(wsSource, i, wsOptions, putRow, "PUT")
                    putRow = putRow + 1
                ElseIf IsCallOption(productName) Then
                    Call ProcessOption(wsSource, i, wsOptions, callRow, "CALL")
                    callRow = callRow + 1
                End If
            Else
                ' Process as stock/ETF
                Call ProcessStock(wsSource, i, wsStocks, stockRow)
                stockRow = stockRow + 1
            End If
        End If
    Next i

    ' Add cash positions at the bottom of Stocks sheet
    Call AddCashPositions(wsSource, wsStocks, stockRow, lastRow)

    ' Format the output sheets
    Call FormatStocksSheet(wsStocks, stockRow)
    Call FormatOptionsSheet(wsOptions, putRow, callRow)

    ' Save the output file
    todayDate = Format(Date, "DD MMMM YYYY")
    outputPath = Application.ActiveWorkbook.Path & "\Transformed_Portfolio_" & todayDate & ".xlsx"

    ' If file exists, add timestamp
    If Dir(outputPath) <> "" Then
        outputPath = Application.ActiveWorkbook.Path & "\Transformed_Portfolio_" & Format(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
    End If

    wbOutput.SaveAs outputPath

    ' Re-enable screen updating
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Transformation complete!" & vbCrLf & vbCrLf & _
           "File saved: " & outputPath & vbCrLf & vbCrLf & _
           "Stocks processed: " & (stockRow - 4) & vbCrLf & _
           "Options processed: " & (putRow - 5 + callRow - (putRow + optionPutRows + 3)), _
           vbInformation, "Bloomberg Data Transformer"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description & vbCrLf & "Line: " & Erl, vbCritical
End Sub

' ====================================================================
' Helper Functions
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

' Extract base ticker from full ticker
' Examples: "META US" -> "META", "AAPL US" -> "AAPL", "2644 JP" -> "2644"
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

' Extract strike price from option name
' Example: "META 01/16/2026 PUT 700" -> "700"
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

' Extract expiry date from option name
' Example: "META 01/16/2026 PUT 700" -> "01/16/2026"
Function ExtractExpiry(productName As String) As String
    Dim parts() As String
    Dim i As Long

    ' Split by space
    parts = Split(productName, " ")

    ' Find the date part (contains /)
    For i = LBound(parts) To UBound(parts)
        If InStr(parts(i), "/") > 0 Then
            ExtractExpiry = parts(i)
            Exit Function
        End If
    Next i

    ExtractExpiry = ""
End Function

' Extract ticker from option name
' Example: "META 01/16/2026 PUT 700" -> "META"
Function ExtractTickerFromOptionName(productName As String) As String
    Dim spacePos As Long
    spacePos = InStr(productName, " ")
    If spacePos > 0 Then
        ExtractTickerFromOptionName = Left(productName, spacePos - 1)
    Else
        ExtractTickerFromOptionName = productName
    End If
End Function

' Get underlying shares for an option
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
' Setup Headers
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
    ws.Cells(2, 9).Value = "P&L"
    ws.Cells(2, 10).Value = "Portfolio Wgt"
    ws.Cells(2, 11).Value = "Attribution"

    ' Row 3 - Sub headers (units)
    ws.Cells(3, 5).Value = "USD"
    ws.Cells(3, 6).Value = "USD"
    ws.Cells(3, 7).Value = "USD"
    ws.Cells(3, 8).Value = "USD"
    ws.Cells(3, 9).Value = "YTD"
    ws.Cells(3, 10).Value = "%"
    ws.Cells(3, 11).Value = "%"

    ' Format headers
    With ws.Range("B2:K3")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

Sub SetupOptionsHeaders(ws As Worksheet)
    ' PUTS Section
    ws.Cells(2, 2).Value = "PUTS"
    ws.Cells(2, 2).Font.Bold = True
    ws.Cells(2, 2).Font.Size = 12

    ' PUTS Headers - Row 3-4
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

    ' Format headers
    With ws.Range("B3:O4")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ====================================================================
' Process Data Rows
' ====================================================================

Sub ProcessStock(wsSource As Worksheet, sourceRow As Long, wsTarget As Worksheet, targetRow As Long)
    ' Column mapping from NAV source to output:
    ' A: Product Name -> B: Name
    ' B: Ticker -> C: Ticker
    ' L: # of Shares -> D: Quantity
    ' E: Unit Cost USD -> E: Unit Cost USD
    ' F: Today USD -> F: Current Px USD (with Bloomberg refresh option)
    ' I: Total Cost USD -> G: Total Cost USD
    ' J: Market Value USD -> H: Mkt Value (formula)
    ' K: Total Net P&L YTD -> I: P&L YTD
    ' D: Portfolio Weight % -> J: Portfolio Wgt %
    ' H: Contribution to Performance -> K: Attribution %

    wsTarget.Cells(targetRow, 2).Value = wsSource.Cells(sourceRow, 1).Value  ' Name
    wsTarget.Cells(targetRow, 3).Value = wsSource.Cells(sourceRow, 2).Value  ' Ticker
    wsTarget.Cells(targetRow, 4).Value = wsSource.Cells(sourceRow, 12).Value ' Quantity
    wsTarget.Cells(targetRow, 5).Value = wsSource.Cells(sourceRow, 5).Value  ' Unit Cost

    ' Current Price - use Bloomberg formula for live data
    Dim ticker As String
    ticker = Trim(CStr(wsSource.Cells(sourceRow, 2).Value))

    If ticker <> "" And InStr(ticker, " ") > 0 Then
        ' Has exchange code - use Bloomberg
        wsTarget.Cells(targetRow, 6).FormulaArray = "=BDP(""" & ticker & " Equity"",""PX_LAST"")"
    Else
        ' No exchange code - use NAV value
        wsTarget.Cells(targetRow, 6).Value = wsSource.Cells(sourceRow, 6).Value
    End If

    wsTarget.Cells(targetRow, 7).Value = wsSource.Cells(sourceRow, 9).Value  ' Total Cost

    ' Market Value - formula
    wsTarget.Cells(targetRow, 8).Formula = "=F" & targetRow & "*D" & targetRow

    wsTarget.Cells(targetRow, 9).Value = wsSource.Cells(sourceRow, 11).Value ' P&L YTD
    wsTarget.Cells(targetRow, 10).Value = wsSource.Cells(sourceRow, 4).Value ' Portfolio Wgt
    wsTarget.Cells(targetRow, 11).Value = wsSource.Cells(sourceRow, 8).Value ' Attribution
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
    occTicker = Trim(CStr(wsSource.Cells(sourceRow, 2).Value)) ' Column B: OCC format
    quantity = wsSource.Cells(sourceRow, 12).Value ' # of contracts
    strike = ExtractStrike(productName)
    expiry = ExtractExpiry(productName)
    baseTicker = ExtractTickerFromOptionName(productName)

    ' Get underlying quantity from stock positions
    underlyingQty = GetUnderlyingShares(baseTicker)

    wsTarget.Cells(targetRow, 2).Value = productName  ' Name
    wsTarget.Cells(targetRow, 3).Value = quantity     ' Quantity (contracts)
    wsTarget.Cells(targetRow, 4).Value = underlyingQty ' Underlying Qty (shares)

    ' % Hedged formula - handle positive (long) and negative (short) positions
    If underlyingQty <> 0 Then
        If optionType = "PUT" Then
            wsTarget.Cells(targetRow, 5).Formula = "=$C" & targetRow & "*100/$D" & targetRow
        Else ' CALL
            wsTarget.Cells(targetRow, 5).Formula = "=-$C" & targetRow & "*100/$D" & targetRow
        End If
    Else
        wsTarget.Cells(targetRow, 5).Value = "N/A"
    End If

    ' Strike Price - extracted from name
    If IsNumeric(strike) Then
        wsTarget.Cells(targetRow, 6).Value = CDbl(strike)
    Else
        wsTarget.Cells(targetRow, 6).Value = strike
    End If

    ' Underlying Price - Bloomberg formula using OCC ticker
    If occTicker <> "" Then
        wsTarget.Cells(targetRow, 7).FormulaArray = "=BDP(""" & occTicker & " Equity"",""OPT_UNDL_PX"")"
    End If

    ' % Moneyness formula
    wsTarget.Cells(targetRow, 8).Formula = "=(G" & targetRow & "-F" & targetRow & ")/F" & targetRow

    ' Expiry
    wsTarget.Cells(targetRow, 9).Value = expiry

    ' Unit Cost
    wsTarget.Cells(targetRow, 10).Value = wsSource.Cells(sourceRow, 5).Value

    ' % Yield formula
    wsTarget.Cells(targetRow, 11).Formula = "=J" & targetRow & "/G" & targetRow

    ' Total Cost
    wsTarget.Cells(targetRow, 12).Value = wsSource.Cells(sourceRow, 9).Value

    ' Current Px
    wsTarget.Cells(targetRow, 13).Value = wsSource.Cells(sourceRow, 6).Value

    ' Mkt Value
    wsTarget.Cells(targetRow, 14).Value = wsSource.Cells(sourceRow, 10).Value

    ' P&L YTD
    wsTarget.Cells(targetRow, 15).Value = wsSource.Cells(sourceRow, 11).Value
End Sub

Sub AddCashPositions(wsSource As Worksheet, wsTarget As Worksheet, startRow As Long, lastRow As Long)
    Dim i As Long
    Dim targetRow As Long
    Dim productName As String

    targetRow = startRow + 2 ' Leave a blank row

    ' Look for cash positions (USD, JPY, CAD, etc.)
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
' Formatting
' ====================================================================

Sub FormatStocksSheet(ws As Worksheet, lastRow As Long)
    ' Number formats
    If lastRow > 4 Then
        ws.Range("D4:D" & lastRow).NumberFormat = "#,##0"
        ws.Range("E4:H" & lastRow).NumberFormat = "$#,##0.00"
        ws.Range("I4:I" & lastRow).NumberFormat = "#,##0"
        ws.Range("J4:K" & lastRow).NumberFormat = "0.0000%"
    End If

    ' Auto-fit columns
    ws.Columns("B:K").AutoFit

    ' Add borders
    If lastRow > 4 Then
        ws.Range("B2:K" & lastRow).Borders.LineStyle = xlContinuous
    End If
End Sub

Sub FormatOptionsSheet(ws As Worksheet, lastPutRow As Long, lastCallRow As Long)
    Dim lastRow As Long
    lastRow = Application.Max(lastPutRow, lastCallRow)

    ' Number formats
    ws.Range("C:C").NumberFormat = "#,##0"          ' Quantity
    ws.Range("D:D").NumberFormat = "#,##0"          ' Underlying Qty
    ws.Range("E:E").NumberFormat = "0.00%"          ' % Hedged
    ws.Range("F:F").NumberFormat = "$#,##0.00"      ' Strike
    ws.Range("G:G").NumberFormat = "$#,##0.00"      ' Underlying Px
    ws.Range("H:H").NumberFormat = "0.00%"          ' % Moneyness
    ws.Range("J:J").NumberFormat = "$#,##0.00"      ' Unit Cost
    ws.Range("K:K").NumberFormat = "0.00%"          ' % Yield
    ws.Range("L:N").NumberFormat = "$#,##0"         ' Costs and Values
    ws.Range("O:O").NumberFormat = "#,##0"          ' P&L

    ' Auto-fit columns
    ws.Columns("B:O").AutoFit

    ' Add borders
    If lastRow > 4 Then
        ws.Range("B2:O" & lastRow).Borders.LineStyle = xlContinuous
    End If
End Sub
