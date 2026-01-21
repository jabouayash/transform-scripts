' ====================================================================
' Mobius Portfolio Report - Outlook Email Monitor
' ====================================================================
' PURPOSE:
'   Automatically monitors incoming emails for daily NAV reports,
'   saves attachments, and triggers Excel transformation when both
'   required emails arrive.
'
' SETUP:
'   1. Open Outlook
'   2. Press Alt+F11 to open VBA Editor
'   3. In Project Explorer, find "ThisOutlookSession"
'   4. Copy ALL code from this file into ThisOutlookSession
'   5. Save and restart Outlook
'   6. When prompted about macros, click "Enable Macros"
'
' TESTING:
'   - Forward the two daily emails to yourself
'   - The script handles "FW:" and "Fwd:" prefixes automatically
'   - Or use: Tools > Macros > RunManualTest
'
' EMAILS MONITORED:
'   1. "Mobius Emerging Opportunities Fund LP| Custom daily portfolio report MMDDYYYY"
'   2. "Mobius Emerging Opportunities Fund LP| Daily Reports MMDDYYYY"
'
' ====================================================================

Option Explicit

' ============================================
' CONFIGURATION - EDIT THESE IF NEEDED
' ============================================
Private Const BASE_FOLDER As String = "C:\Mobius Reports"
Private Const INCOMING_FOLDER As String = "C:\Mobius Reports\Incoming"
Private Const TRANSFORMED_FOLDER As String = "C:\Mobius Reports\Transformed"
Private Const ARCHIVE_FOLDER As String = "C:\Mobius Reports\Archive"

' Email subject patterns (without date)
Private Const SUBJECT_CUSTOM As String = "Mobius Emerging Opportunities Fund LP| Custom daily portfolio report"
Private Const SUBJECT_DAILY As String = "Mobius Emerging Opportunities Fund LP| Daily Reports"

' File name patterns to save
Private Const FILE_CUSTOM As String = "Gain And Exposure_Custom_MOBIUS EMERGING OPPORTUNITIES FUND LP"
Private Const FILE_DAILY As String = "Gain And Exposure_MOBIUS EMERGING OPPORTUNITIES FUND LP"

' Path to your Excel transformer workbook (with the VBA macro)
Private Const EXCEL_TRANSFORMER_PATH As String = "C:\Mobius Reports\Portfolio Transformer.xlsm"

' Logging and persistence files
Private Const LOG_FILE As String = "C:\Mobius Reports\monitor_log.txt"
Private Const TRACKER_FILE As String = "C:\Mobius Reports\email_tracker.txt"

' ============================================
' TRACKING VARIABLES
' ============================================
' Dictionary to track which emails have arrived for each date
' Key = date string (MMDDYYYY), Value = "CUSTOM", "DAILY", or "BOTH"
Private emailTracker As Object

' ============================================
' OUTLOOK EVENT - RUNS ON STARTUP
' ============================================
Private WithEvents InboxItems As Outlook.Items

Private Sub Application_Startup()
    ' Initialize when Outlook starts (silent - no popup)
    Call InitializeMonitor
End Sub

Public Sub InitializeMonitor()
    ' Set up the inbox monitor
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    Set InboxItems = ns.GetDefaultFolder(olFolderInbox).Items

    ' Initialize tracker
    Set emailTracker = CreateObject("Scripting.Dictionary")

    ' Create folders if they don't exist
    Call EnsureFoldersExist

    ' Load any saved tracker state (survives Outlook restart)
    Call LoadTracker

    Call WriteLog("Monitor initialized")
End Sub

' ============================================
' MAIN EVENT - TRIGGERED ON NEW EMAIL
' ============================================
Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler

    If TypeOf Item Is Outlook.MailItem Then
        Dim mail As Outlook.MailItem
        Set mail = Item

        ' Check if this is one of our target emails
        Call ProcessIncomingEmail(mail)
    End If

    Exit Sub
ErrorHandler:
    ' Silent fail - don't interrupt user for errors
    Call WriteLog("ERROR in ItemAdd: " & Err.Description)
End Sub

' ============================================
' EMAIL PROCESSING
' ============================================
Private Sub ProcessIncomingEmail(mail As Outlook.MailItem)
    Dim subject As String
    Dim cleanedSubject As String
    Dim reportDate As String
    Dim emailType As String

    subject = mail.subject

    ' Remove FW:/Fwd: prefixes for testing with forwarded emails
    cleanedSubject = StripForwardPrefixes(subject)

    ' Check if this matches our patterns
    If InStr(1, cleanedSubject, SUBJECT_CUSTOM, vbTextCompare) > 0 Then
        emailType = "CUSTOM"
        reportDate = ExtractDateFromSubject(cleanedSubject, SUBJECT_CUSTOM)
    ElseIf InStr(1, cleanedSubject, SUBJECT_DAILY, vbTextCompare) > 0 Then
        emailType = "DAILY"
        reportDate = ExtractDateFromSubject(cleanedSubject, SUBJECT_DAILY)
    Else
        ' Not a target email, ignore
        Exit Sub
    End If

    If reportDate = "" Then
        Call WriteLog("Could not extract date from subject: " & subject)
        Exit Sub
    End If

    ' Save attachments
    Call SaveAttachments(mail, emailType, reportDate)

    ' Update tracker
    Call UpdateTracker(reportDate, emailType)

    ' Check if we have both emails now
    If emailTracker(reportDate) = "BOTH" Then
        Call TriggerTransformation(reportDate)
    End If
End Sub

Private Function StripForwardPrefixes(subject As String) As String
    ' Remove forward/reply prefixes - handles multiple like "FW: FW: FW: ..."
    Dim clean As String
    Dim previousClean As String

    clean = Trim(subject)

    ' Loop until no more prefixes are found
    Do
        previousClean = clean

        ' Remove common prefixes (case-insensitive check)
        If Len(clean) >= 4 And UCase(Left(clean, 4)) = "FW: " Then
            clean = Trim(Mid(clean, 5))
        ElseIf Len(clean) >= 5 And UCase(Left(clean, 5)) = "FWD: " Then
            clean = Trim(Mid(clean, 6))
        ElseIf Len(clean) >= 4 And UCase(Left(clean, 4)) = "RE: " Then
            clean = Trim(Mid(clean, 5))
        End If

    Loop While clean <> previousClean  ' Keep going until nothing changes

    StripForwardPrefixes = clean
End Function

Private Function ExtractDateFromSubject(subject As String, pattern As String) As String
    ' Extract the MMDDYYYY date from the end of the subject
    Dim dateStart As Long
    Dim dateStr As String

    dateStart = Len(pattern) + 2  ' +2 for space after pattern

    If Len(subject) >= dateStart + 7 Then
        dateStr = Mid(subject, dateStart, 8)

        ' Validate it looks like a date (8 digits)
        If IsNumeric(dateStr) And Len(dateStr) = 8 Then
            ExtractDateFromSubject = dateStr
        Else
            ExtractDateFromSubject = ""
        End If
    Else
        ExtractDateFromSubject = ""
    End If
End Function

Private Sub UpdateTracker(reportDate As String, emailType As String)
    If Not emailTracker.Exists(reportDate) Then
        emailTracker(reportDate) = emailType
        Call WriteLog("Received " & emailType & " report for " & FormatReportDate(reportDate))
    ElseIf emailTracker(reportDate) <> emailType Then
        ' We now have both types
        emailTracker(reportDate) = "BOTH"
        Call WriteLog("Received " & emailType & " report for " & FormatReportDate(reportDate) & " - BOTH now received")
    End If
    ' If same type arrives twice, just keep the existing value

    ' Save tracker state to file (survives Outlook restart)
    Call SaveTracker
End Sub

' ============================================
' ATTACHMENT SAVING
' ============================================
Private Sub SaveAttachments(mail As Outlook.MailItem, emailType As String, reportDate As String)
    Dim att As Outlook.Attachment
    Dim savePath As String
    Dim targetPattern As String

    ' Determine which file pattern to look for
    If emailType = "CUSTOM" Then
        targetPattern = FILE_CUSTOM
    Else
        targetPattern = FILE_DAILY
    End If

    For Each att In mail.Attachments
        ' Check if this attachment matches our target file
        If InStr(1, att.FileName, targetPattern, vbTextCompare) > 0 Then
            ' Check if it's an Excel file
            If Right(LCase(att.FileName), 5) = ".xlsx" Or Right(LCase(att.FileName), 4) = ".xls" Then
                savePath = INCOMING_FOLDER & "\" & att.FileName

                ' Delete existing file if present (in case of re-processing)
                If Dir(savePath) <> "" Then
                    Kill savePath
                End If

                att.SaveAsFile savePath
                Call WriteLog("Saved attachment: " & att.FileName)
            End If
        End If
    Next att
End Sub

' ============================================
' TRANSFORMATION TRIGGER
' ============================================
Private Sub TriggerTransformation(reportDate As String)
    Dim customFile As String
    Dim dailyFile As String
    Dim msg As String

    ' Build expected file paths
    customFile = INCOMING_FOLDER & "\" & FILE_CUSTOM & "_" & reportDate & ".XLSX"
    dailyFile = INCOMING_FOLDER & "\" & FILE_DAILY & "_" & reportDate & ".XLSX"

    ' Verify both files exist
    If Dir(customFile) = "" Then
        Call WriteLog("ERROR: Custom file not found: " & customFile)
        MsgBox "Custom file not found: " & customFile, vbExclamation, "File Missing"
        Exit Sub
    End If

    If Dir(dailyFile) = "" Then
        Call WriteLog("ERROR: Daily file not found: " & dailyFile)
        MsgBox "Daily file not found: " & dailyFile, vbExclamation, "File Missing"
        Exit Sub
    End If

    ' Notify user
    msg = "Both daily reports received for " & FormatReportDate(reportDate) & "!" & vbCrLf & vbCrLf
    msg = msg & "Starting transformation..." & vbCrLf & vbCrLf
    msg = msg & "Files:" & vbCrLf
    msg = msg & "- Custom: " & Dir(customFile) & vbCrLf
    msg = msg & "- Daily: " & Dir(dailyFile)

    MsgBox msg, vbInformation, "Processing Reports"

    ' Launch Excel and run the transformation
    Call RunExcelTransformation(customFile, dailyFile, reportDate)

    ' Archive processed files
    Call ArchiveProcessedFiles(reportDate)
    Call WriteLog("Transformation completed for " & FormatReportDate(reportDate))
End Sub

Private Sub RunExcelTransformation(customFile As String, dailyFile As String, reportDate As String)
    Dim xlApp As Object
    Dim xlWb As Object
    Dim alreadyOpen As Boolean

    On Error Resume Next

    ' Try to get existing Excel instance
    Set xlApp = GetObject(, "Excel.Application")

    If xlApp Is Nothing Then
        ' Start new Excel instance
        Set xlApp = CreateObject("Excel.Application")
        alreadyOpen = False
    Else
        alreadyOpen = True
    End If

    On Error GoTo ErrorHandler

    xlApp.Visible = True

    ' Open the transformer workbook (contains the macro)
    If Dir(EXCEL_TRANSFORMER_PATH) = "" Then
        ' If transformer workbook doesn't exist, open the custom file directly
        Set xlWb = xlApp.Workbooks.Open(customFile)
        MsgBox "Transformer workbook not found at:" & vbCrLf & EXCEL_TRANSFORMER_PATH & vbCrLf & vbCrLf & _
               "Please run the TransformBloombergData macro manually.", vbExclamation, "Manual Step Required"
    Else
        ' Open transformer first, then the data file
        Set xlWb = xlApp.Workbooks.Open(EXCEL_TRANSFORMER_PATH)
        xlApp.Workbooks.Open customFile

        ' Store the daily file path for the macro to read K94
        ' We'll use a named range or environment variable approach
        xlApp.Run "'" & xlWb.Name & "'!SetDailyFilePath", dailyFile

        ' Run the transformation macro
        xlApp.Run "'" & xlWb.Name & "'!TransformBloombergData"

        MsgBox "Transformation complete!" & vbCrLf & vbCrLf & _
               "Output saved to: " & TRANSFORMED_FOLDER, vbInformation, "Done"
    End If

    Exit Sub

ErrorHandler:
    Call WriteLog("ERROR in RunExcelTransformation: " & Err.Description)
    MsgBox "Error during transformation: " & Err.Description & vbCrLf & vbCrLf & _
           "Please run the macro manually.", vbExclamation, "Error"
End Sub

Private Function FormatReportDate(dateStr As String) As String
    ' Convert MMDDYYYY to readable format
    Dim m As String, d As String, y As String

    If Len(dateStr) = 8 Then
        m = Left(dateStr, 2)
        d = Mid(dateStr, 3, 2)
        y = Right(dateStr, 4)
        FormatReportDate = m & "/" & d & "/" & y
    Else
        FormatReportDate = dateStr
    End If
End Function

' ============================================
' FOLDER MANAGEMENT
' ============================================
Private Sub EnsureFoldersExist()
    ' Create folder structure if it doesn't exist
    If Dir(BASE_FOLDER, vbDirectory) = "" Then MkDir BASE_FOLDER
    If Dir(INCOMING_FOLDER, vbDirectory) = "" Then MkDir INCOMING_FOLDER
    If Dir(TRANSFORMED_FOLDER, vbDirectory) = "" Then MkDir TRANSFORMED_FOLDER
    If Dir(ARCHIVE_FOLDER, vbDirectory) = "" Then MkDir ARCHIVE_FOLDER
End Sub

' ============================================
' MANUAL TESTING FUNCTIONS
' ============================================
Public Sub RunManualTest()
    ' Use this to test the setup without waiting for emails
    ' Run from: Tools > Macros > RunManualTest

    Dim msg As String

    ' Initialize if not already done
    If emailTracker Is Nothing Then
        Call InitializeMonitor
    End If

    msg = "=== Mobius Report Monitor Test ===" & vbCrLf & vbCrLf
    msg = msg & "Status: ACTIVE" & vbCrLf & vbCrLf
    msg = msg & "Watching for emails with subjects:" & vbCrLf
    msg = msg & "1. " & SUBJECT_CUSTOM & " [DATE]" & vbCrLf
    msg = msg & "2. " & SUBJECT_DAILY & " [DATE]" & vbCrLf & vbCrLf
    msg = msg & "Folders:" & vbCrLf
    msg = msg & "- Incoming: " & INCOMING_FOLDER & vbCrLf
    msg = msg & "- Output: " & TRANSFORMED_FOLDER & vbCrLf & vbCrLf
    msg = msg & "To test: Forward both daily emails to yourself."

    MsgBox msg, vbInformation, "Monitor Status"
End Sub

Public Sub ProcessSelectedEmail()
    ' Manually process a selected email in Outlook
    ' Select an email, then run this macro

    Dim sel As Outlook.Selection
    Dim mail As Outlook.MailItem

    Set sel = Application.ActiveExplorer.Selection

    If sel.Count = 0 Then
        MsgBox "Please select an email first.", vbExclamation, "No Selection"
        Exit Sub
    End If

    If TypeOf sel.Item(1) Is Outlook.MailItem Then
        Set mail = sel.Item(1)

        ' Initialize if needed
        If emailTracker Is Nothing Then
            Call InitializeMonitor
        End If

        Call ProcessIncomingEmail(mail)
        MsgBox "Processed: " & mail.subject, vbInformation, "Done"
    Else
        MsgBox "Selected item is not an email.", vbExclamation, "Invalid Selection"
    End If
End Sub

Public Sub CheckFolderContents()
    ' Show what files are in the Incoming folder

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim msg As String
    Dim fileCount As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(INCOMING_FOLDER) Then
        MsgBox "Incoming folder does not exist: " & INCOMING_FOLDER, vbExclamation, "Folder Missing"
        Exit Sub
    End If

    Set folder = fso.GetFolder(INCOMING_FOLDER)

    msg = "Files in Incoming folder:" & vbCrLf & vbCrLf
    fileCount = 0

    For Each file In folder.Files
        msg = msg & "- " & file.Name & vbCrLf
        fileCount = fileCount + 1
    Next file

    If fileCount = 0 Then
        msg = msg & "(empty)"
    End If

    MsgBox msg, vbInformation, "Incoming Folder (" & fileCount & " files)"
End Sub

Public Sub ClearIncomingFolder()
    ' Clear all files from the Incoming folder

    Dim result As VbMsgBoxResult
    result = MsgBox("Delete all files in the Incoming folder?", vbQuestion + vbYesNo, "Confirm")

    If result = vbYes Then
        On Error Resume Next
        Kill INCOMING_FOLDER & "\*.*"
        On Error GoTo 0
        MsgBox "Incoming folder cleared.", vbInformation, "Done"
    End If
End Sub

Public Sub ResetTracker()
    ' Reset the email tracker (for testing)
    Set emailTracker = CreateObject("Scripting.Dictionary")
    Call SaveTracker ' Clear the tracker file too
    Call WriteLog("Tracker reset by user")
    MsgBox "Email tracker has been reset.", vbInformation, "Reset Complete"
End Sub

' ============================================
' HELPER: Show current tracker state
' ============================================
Public Sub ShowTrackerState()
    Dim msg As String
    Dim key As Variant

    If emailTracker Is Nothing Or emailTracker.Count = 0 Then
        msg = "No emails tracked yet."
    Else
        msg = "Tracked emails:" & vbCrLf & vbCrLf
        For Each key In emailTracker.Keys
            msg = msg & FormatReportDate(CStr(key)) & ": " & emailTracker(key) & vbCrLf
        Next key
    End If

    MsgBox msg, vbInformation, "Email Tracker State"
End Sub

' ============================================
' LOGGING - Persistent log to file
' ============================================
Private Sub WriteLog(message As String)
    ' Writes timestamped log entry to file
    Dim fso As Object
    Dim logFile As Object
    Dim timestamp As String

    On Error Resume Next

    timestamp = Format(Now, "yyyy-mm-dd hh:mm:ss")

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.OpenTextFile(LOG_FILE, 8, True) ' 8 = append, True = create if missing
    logFile.WriteLine timestamp & " - " & message
    logFile.Close

    ' Also write to debug for immediate window
    Debug.Print timestamp & " - " & message

    On Error GoTo 0
End Sub

' ============================================
' TRACKER PERSISTENCE - Survives Outlook restart
' ============================================
Private Sub SaveTracker()
    ' Saves tracker state to file
    Dim fso As Object
    Dim trackerFile As Object
    Dim key As Variant

    On Error Resume Next

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set trackerFile = fso.CreateTextFile(TRACKER_FILE, True) ' True = overwrite

    If Not emailTracker Is Nothing Then
        For Each key In emailTracker.Keys
            trackerFile.WriteLine CStr(key) & "=" & emailTracker(key)
        Next key
    End If

    trackerFile.Close

    On Error GoTo 0
End Sub

Private Sub LoadTracker()
    ' Loads tracker state from file
    Dim fso As Object
    Dim trackerFile As Object
    Dim line As String
    Dim parts() As String

    On Error Resume Next

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(TRACKER_FILE) Then
        Set trackerFile = fso.OpenTextFile(TRACKER_FILE, 1) ' 1 = read

        Do While Not trackerFile.AtEndOfStream
            line = trackerFile.ReadLine
            If InStr(line, "=") > 0 Then
                parts = Split(line, "=")
                If UBound(parts) >= 1 Then
                    ' Only load if not already "BOTH" (already processed)
                    If parts(1) <> "BOTH" Then
                        emailTracker(parts(0)) = parts(1)
                    End If
                End If
            End If
        Loop

        trackerFile.Close
        Call WriteLog("Loaded tracker state: " & emailTracker.Count & " pending dates")
    End If

    On Error GoTo 0
End Sub

' ============================================
' FILE ARCHIVING - Move processed files
' ============================================
Private Sub ArchiveProcessedFiles(reportDate As String)
    ' Moves processed files from Incoming to Archive
    Dim fso As Object
    Dim customFile As String
    Dim dailyFile As String
    Dim archiveCustom As String
    Dim archiveDaily As String

    On Error Resume Next

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Build file paths
    customFile = INCOMING_FOLDER & "\" & FILE_CUSTOM & "_" & reportDate & ".XLSX"
    dailyFile = INCOMING_FOLDER & "\" & FILE_DAILY & "_" & reportDate & ".XLSX"
    archiveCustom = ARCHIVE_FOLDER & "\" & FILE_CUSTOM & "_" & reportDate & ".XLSX"
    archiveDaily = ARCHIVE_FOLDER & "\" & FILE_DAILY & "_" & reportDate & ".XLSX"

    ' Move custom file
    If fso.FileExists(customFile) Then
        If fso.FileExists(archiveCustom) Then fso.DeleteFile archiveCustom
        fso.MoveFile customFile, archiveCustom
        Call WriteLog("Archived: " & FILE_CUSTOM & "_" & reportDate & ".XLSX")
    End If

    ' Move daily file
    If fso.FileExists(dailyFile) Then
        If fso.FileExists(archiveDaily) Then fso.DeleteFile archiveDaily
        fso.MoveFile dailyFile, archiveDaily
        Call WriteLog("Archived: " & FILE_DAILY & "_" & reportDate & ".XLSX")
    End If

    On Error GoTo 0
End Sub
