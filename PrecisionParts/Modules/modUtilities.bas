Attribute VB_Name = "modUtilities"
' ============================================================================
' Module: modUtilities
' Description: String, date, number utility functions for Precision Parts System
' Created: January 2001
' Last Modified: March 2024
' ============================================================================

Option Explicit

' Windows API declarations for miscellaneous utilities
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' ============================================================================
' Function: SafeString
' Description: Returns an empty string if the value is Null, otherwise CStr
' ============================================================================
Public Function SafeString(ByVal vValue As Variant) As String
    On Error Resume Next
    If IsNull(vValue) Or IsEmpty(vValue) Then
        SafeString = ""
    Else
        SafeString = CStr(vValue)
    End If
End Function

' ============================================================================
' Function: SafeNumber
' Description: Returns 0 if the value is Null/Empty/non-numeric, otherwise CDbl
' ============================================================================
Public Function SafeNumber(ByVal vValue As Variant) As Double
    On Error Resume Next
    If IsNull(vValue) Or IsEmpty(vValue) Then
        SafeNumber = 0
    ElseIf IsNumeric(vValue) Then
        SafeNumber = CDbl(vValue)
    Else
        SafeNumber = 0
    End If
End Function

' ============================================================================
' Function: SafeLong
' Description: Returns 0 if the value is Null/Empty, otherwise CLng
' ============================================================================
Public Function SafeLong(ByVal vValue As Variant) As Long
    On Error Resume Next
    If IsNull(vValue) Or IsEmpty(vValue) Then
        SafeLong = 0
    ElseIf IsNumeric(vValue) Then
        SafeLong = CLng(vValue)
    Else
        SafeLong = 0
    End If
End Function

' ============================================================================
' Function: SafeDate
' Description: Returns a default date if value is Null, otherwise CDate
' ============================================================================
Public Function SafeDate(ByVal vValue As Variant, Optional ByVal dtDefault As Date = #1/1/1900#) As Date
    On Error Resume Next
    If IsNull(vValue) Or IsEmpty(vValue) Then
        SafeDate = dtDefault
    ElseIf IsDate(vValue) Then
        SafeDate = CDate(vValue)
    Else
        SafeDate = dtDefault
    End If
End Function

' ============================================================================
' Function: SafeBool
' Description: Returns False if value is Null, otherwise CBool
' ============================================================================
Public Function SafeBool(ByVal vValue As Variant) As Boolean
    On Error Resume Next
    If IsNull(vValue) Or IsEmpty(vValue) Then
        SafeBool = False
    Else
        SafeBool = CBool(vValue)
    End If
End Function

' ============================================================================
' Function: FormatCurrencyValue
' Description: Formats a number as currency string
' ============================================================================
Public Function FormatCurrencyValue(ByVal dAmount As Double) As String
    On Error GoTo ErrHandler
    
    FormatCurrencyValue = Format$(dAmount, "$#,##0.00")
    
    Exit Function
    
ErrHandler:
    FormatCurrencyValue = "$0.00"
End Function

' ============================================================================
' Function: FormatQuantity
' Description: Formats an integer quantity with thousand separators
' ============================================================================
Public Function FormatQuantity(ByVal lQty As Long) As String
    On Error GoTo ErrHandler
    
    FormatQuantity = Format$(lQty, "#,##0")
    
    Exit Function
    
ErrHandler:
    FormatQuantity = "0"
End Function

' ============================================================================
' Function: IsValidPartNumber
' Description: Validates a part number format (PP-XXXX-NNN)
' Parameters: sPartNumber - Part number to validate
' Returns: True if format is valid
' ============================================================================
Public Function IsValidPartNumber(ByVal sPartNumber As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim sParts() As String
    
    IsValidPartNumber = False
    
    ' Must not be empty
    If Len(Trim$(sPartNumber)) = 0 Then Exit Function
    
    ' Must match pattern PP-XXXX-NNN or similar
    ' Split on hyphens
    sParts = Split(sPartNumber, "-")
    
    ' Must have at least 2 parts
    If UBound(sParts) < 1 Then Exit Function
    
    ' First part should be 2-4 alpha characters
    If Len(sParts(0)) < 2 Or Len(sParts(0)) > 4 Then Exit Function
    
    ' Check first part is all alpha
    Dim i As Integer
    Dim sChar As String
    For i = 1 To Len(sParts(0))
        sChar = Mid$(sParts(0), i, 1)
        If sChar < "A" Or sChar > "Z" Then
            If sChar < "a" Or sChar > "z" Then
                Exit Function
            End If
        End If
    Next i
    
    ' Total length should be reasonable
    If Len(sPartNumber) < 5 Or Len(sPartNumber) > 20 Then Exit Function
    
    IsValidPartNumber = True
    
    Exit Function
    
ErrHandler:
    IsValidPartNumber = False
End Function

' ============================================================================
' Function: SQLSafe
' Description: Escapes single quotes in a string for safe SQL insertion
' Parameters: sValue - String to escape
' Returns: Escaped string (replaces ' with '')
' ============================================================================
Public Function SQLSafe(ByVal sValue As String) As String
    On Error Resume Next
    SQLSafe = Replace(sValue, "'", "''")
End Function

' ============================================================================
' Function: SQLDate
' Description: Formats a date for use in SQL statements (Access format)
' Parameters: dtValue - Date to format
' Returns: Date string wrapped in # delimiters
' ============================================================================
Public Function SQLDate(ByVal dtValue As Date) As String
    On Error Resume Next
    SQLDate = "#" & Format$(dtValue, "mm/dd/yyyy") & "#"
End Function

' ============================================================================
' Function: SQLDateTime
' Description: Formats a date/time for use in SQL statements
' ============================================================================
Public Function SQLDateTime(ByVal dtValue As Date) As String
    On Error Resume Next
    SQLDateTime = "#" & Format$(dtValue, "mm/dd/yyyy hh:nn:ss") & "#"
End Function

' ============================================================================
' Sub: LogMessage
' Description: Writes a message to the application log file
' Parameters: sMessage - Message to log
' ============================================================================
Public Sub LogMessage(ByVal sMessage As String)
    On Error Resume Next
    
    If Not g_EnableLogging Then Exit Sub
    
    Dim iFile As Integer
    Dim sLogFile As String
    Dim sLogEntry As String
    
    sLogFile = g_LogFilePath & "PrecisionParts_" & Format$(Date, "yyyymmdd") & ".log"
    
    ' Create log directory if it doesn't exist
    If Dir(g_LogFilePath, vbDirectory) = "" Then
        MkDir g_LogFilePath
    End If
    
    sLogEntry = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
                g_CurrentUsername & " | " & sMessage
    
    iFile = FreeFile
    Open sLogFile For Append As #iFile
    Print #iFile, sLogEntry
    Close #iFile
End Sub

' ============================================================================
' Sub: ShowError
' Description: Displays an error message box with consistent formatting
' Parameters: sSource - Where the error occurred
'             sMessage - Error description
'             Optional lErrNum - Error number
' ============================================================================
Public Sub ShowError(ByVal sSource As String, ByVal sMessage As String, Optional ByVal lErrNum As Long = 0)
    Dim sFullMessage As String
    
    sFullMessage = "An error occurred in: " & sSource & vbCrLf & vbCrLf & _
                   "Error: " & sMessage
    
    If lErrNum <> 0 Then
        sFullMessage = sFullMessage & vbCrLf & "Error Number: " & lErrNum
    End If
    
    sFullMessage = sFullMessage & vbCrLf & vbCrLf & _
                   "Please contact your system administrator if this problem persists."
    
    LogMessage "ERROR in " & sSource & ": " & sMessage & " (#" & lErrNum & ")"
    
    MsgBox sFullMessage, vbCritical + vbOKOnly, APP_TITLE
End Sub

' ============================================================================
' Function: ConfirmAction
' Description: Shows a Yes/No confirmation dialog
' Parameters: sMessage - Confirmation message
'             Optional sTitle - Dialog title
' Returns: True if user clicked Yes
' ============================================================================
Public Function ConfirmAction(ByVal sMessage As String, Optional ByVal sTitle As String = "") As Boolean
    If sTitle = "" Then sTitle = APP_TITLE
    
    ConfirmAction = (MsgBox(sMessage, vbQuestion + vbYesNo, sTitle) = vbYes)
End Function

' ============================================================================
' Sub: ExportToCSV
' Description: Exports an ADO recordset to a CSV file
' Parameters: rs - Recordset to export
'             sFilePath - Full path for the CSV file
'             bIncludeHeaders - Whether to include column headers
' ============================================================================
Public Sub ExportToCSV(ByVal rs As ADODB.Recordset, ByVal sFilePath As String, _
                       Optional ByVal bIncludeHeaders As Boolean = True)
    On Error GoTo ErrHandler
    
    If rs Is Nothing Then
        MsgBox "No data to export.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    If rs.EOF And rs.BOF Then
        MsgBox "No records to export.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    Dim iFile As Integer
    Dim iCol As Integer
    Dim sLine As String
    Dim lRowCount As Long
    
    iFile = FreeFile
    Open sFilePath For Output As #iFile
    
    ' Write header row
    If bIncludeHeaders Then
        sLine = ""
        For iCol = 0 To rs.Fields.Count - 1
            If iCol > 0 Then sLine = sLine & ","
            sLine = sLine & """" & rs.Fields(iCol).Name & """"
        Next iCol
        Print #iFile, sLine
    End If
    
    ' Write data rows
    rs.MoveFirst
    lRowCount = 0
    Do While Not rs.EOF
        sLine = ""
        For iCol = 0 To rs.Fields.Count - 1
            If iCol > 0 Then sLine = sLine & ","
            sLine = sLine & """" & SafeString(rs.Fields(iCol).Value) & """"
        Next iCol
        Print #iFile, sLine
        
        lRowCount = lRowCount + 1
        
        ' Allow UI to update during long exports (VB6 anti-pattern)
        If lRowCount Mod 100 = 0 Then
            DoEvents
        End If
        
        rs.MoveNext
    Loop
    
    Close #iFile
    
    MsgBox "Export completed successfully." & vbCrLf & _
           lRowCount & " records exported to:" & vbCrLf & sFilePath, _
           vbInformation, APP_TITLE
    
    LogMessage "CSV Export: " & lRowCount & " records to " & sFilePath
    
    Exit Sub
    
ErrHandler:
    On Error Resume Next
    Close #iFile
    
    ShowError "ExportToCSV", Err.Description, Err.Number
End Sub

' ============================================================================
' Function: BrowseForFile
' Description: Shows a file save dialog and returns the selected path
' Parameters: sFilter - File filter string
'             sTitle - Dialog title
'             sDefaultName - Default file name
' Returns: Selected file path or empty string if cancelled
' ============================================================================
Public Function BrowseForFile(Optional ByVal sFilter As String = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*", _
                              Optional ByVal sTitle As String = "Save As", _
                              Optional ByVal sDefaultName As String = "") As String
    On Error GoTo ErrHandler
    
    Dim dlg As Object
    Set dlg = CreateObject("MSComDlg.CommonDialog")
    
    dlg.Filter = sFilter
    dlg.DialogTitle = sTitle
    dlg.FileName = sDefaultName
    dlg.CancelError = True
    dlg.ShowSave
    
    BrowseForFile = dlg.FileName
    
    Set dlg = Nothing
    
    Exit Function
    
ErrHandler:
    ' User cancelled
    BrowseForFile = ""
    Set dlg = Nothing
End Function

' ============================================================================
' Function: PadLeft
' Description: Pads a string on the left to a specified length
' ============================================================================
Public Function PadLeft(ByVal sText As String, ByVal iLength As Integer, Optional ByVal sPadChar As String = " ") As String
    If Len(sText) >= iLength Then
        PadLeft = Left$(sText, iLength)
    Else
        PadLeft = String$(iLength - Len(sText), sPadChar) & sText
    End If
End Function

' ============================================================================
' Function: PadRight
' Description: Pads a string on the right to a specified length
' ============================================================================
Public Function PadRight(ByVal sText As String, ByVal iLength As Integer, Optional ByVal sPadChar As String = " ") As String
    If Len(sText) >= iLength Then
        PadRight = Left$(sText, iLength)
    Else
        PadRight = sText & String$(iLength - Len(sText), sPadChar)
    End If
End Function

' ============================================================================
' Function: TrimAll
' Description: Removes leading, trailing, and excess internal whitespace
' ============================================================================
Public Function TrimAll(ByVal sText As String) As String
    On Error Resume Next
    
    Dim sResult As String
    sResult = Trim$(sText)
    
    ' Remove double spaces
    Do While InStr(sResult, "  ") > 0
        sResult = Replace(sResult, "  ", " ")
    Loop
    
    TrimAll = sResult
End Function

' ============================================================================
' Function: IsValidEmail
' Description: Basic email validation
' ============================================================================
Public Function IsValidEmail(ByVal sEmail As String) As Boolean
    On Error Resume Next
    
    IsValidEmail = False
    
    If Len(sEmail) < 5 Then Exit Function
    If InStr(sEmail, "@") = 0 Then Exit Function
    If InStr(sEmail, ".") = 0 Then Exit Function
    If InStr(sEmail, " ") > 0 Then Exit Function
    
    Dim iAtPos As Integer
    iAtPos = InStr(sEmail, "@")
    
    If iAtPos < 2 Then Exit Function
    If iAtPos >= Len(sEmail) - 1 Then Exit Function
    
    IsValidEmail = True
End Function

' ============================================================================
' Function: GenerateWorkOrderNumber
' Description: Generates a new work order number in format WO-YYYYMMDD-NNN
' ============================================================================
Public Function GenerateWorkOrderNumber() As String
    On Error GoTo ErrHandler
    
    Dim sDate As String
    Dim lSeq As Long
    Dim sSQL As String
    
    sDate = Format$(Date, "yyyymmdd")
    
    ' Get the next sequence number for today
    sSQL = "SELECT COUNT(*) FROM WorkOrders WHERE WorkOrderNumber LIKE 'WO-" & sDate & "-%'"
    lSeq = SafeLong(GetScalarValue(sSQL, 0)) + 1
    
    GenerateWorkOrderNumber = "WO-" & sDate & "-" & Format$(lSeq, "000")
    
    Exit Function
    
ErrHandler:
    GenerateWorkOrderNumber = "WO-" & Format$(Now, "yyyymmddhhnnss")
End Function

' ============================================================================
' Sub: CenterForm
' Description: Centers a form on the screen
' ============================================================================
Public Sub CenterForm(frm As Form)
    On Error Resume Next
    frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2
End Sub

' ============================================================================
' Sub: SetWaitCursor
' Description: Sets/resets the hourglass cursor
' ============================================================================
Public Sub SetWaitCursor(ByVal bWait As Boolean)
    On Error Resume Next
    If bWait Then
        Screen.MousePointer = vbHourglass
        g_SystemBusy = True
    Else
        Screen.MousePointer = vbDefault
        g_SystemBusy = False
    End If
    DoEvents
End Sub
