Attribute VB_Name = "modDatabase"
' ============================================================================
' Module: modDatabase
' Description: ADO database connection management for Precision Parts System
' Created: January 2001
' Last Modified: March 2024
' Notes: Uses global connection object - classic VB6 pattern
' ============================================================================

Option Explicit

' Global ADO connection - shared across entire application
Public g_Connection As ADODB.Connection

' Connection state tracking
Private m_ConnectionRetries As Integer
Private m_LastError As String

' ============================================================================
' Function: OpenDatabaseConnection
' Description: Opens a global ADO connection to the Access database
' Returns: True if connection was successful
' ============================================================================
Public Function OpenDatabaseConnection() As Boolean
    On Error GoTo ErrHandler
    
    Dim sConnString As String
    
    ' Close existing connection if open
    If Not g_Connection Is Nothing Then
        If g_Connection.State = adStateOpen Then
            g_Connection.Close
        End If
        Set g_Connection = Nothing
    End If
    
    Set g_Connection = New ADODB.Connection
    
    ' Build connection string - hardcoded provider (VB6 anti-pattern)
    sConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & g_DatabasePath & ";" & _
                  "Persist Security Info=False;" & _
                  "Jet OLEDB:Database Password=pp2001mfg;"
    
    g_Connection.ConnectionTimeout = DB_TIMEOUT
    g_Connection.CommandTimeout = DB_TIMEOUT
    g_Connection.CursorLocation = adUseClient
    g_Connection.Mode = adModeShareDenyNone
    
    g_Connection.Open sConnString
    
    If g_Connection.State = adStateOpen Then
        OpenDatabaseConnection = True
        m_ConnectionRetries = 0
        If g_EnableLogging Then
            LogMessage "Database connection opened successfully: " & g_DatabasePath
        End If
    Else
        OpenDatabaseConnection = False
        LogMessage "ERROR: Database connection failed - state not open"
    End If
    
    Exit Function
    
ErrHandler:
    m_LastError = "OpenDatabaseConnection Error #" & Err.Number & ": " & Err.Description
    LogMessage m_LastError
    
    m_ConnectionRetries = m_ConnectionRetries + 1
    If m_ConnectionRetries < MAX_RETRY_COUNT Then
        DoEvents
        ' Wait and retry (blocking - VB6 anti-pattern)
        Dim lStart As Long
        lStart = Timer
        Do While Timer - lStart < 2
            DoEvents
        Loop
        Resume
    End If
    
    MsgBox "Unable to connect to database." & vbCrLf & vbCrLf & _
           "Path: " & g_DatabasePath & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Please contact your system administrator.", _
           vbCritical + vbOKOnly, APP_TITLE
    
    OpenDatabaseConnection = False
    Set g_Connection = Nothing
End Function

' ============================================================================
' Sub: CloseDatabaseConnection
' Description: Closes the global database connection
' ============================================================================
Public Sub CloseDatabaseConnection()
    On Error Resume Next
    
    If Not g_Connection Is Nothing Then
        If g_Connection.State = adStateOpen Then
            g_Connection.Close
            LogMessage "Database connection closed"
        End If
        Set g_Connection = Nothing
    End If
    
    On Error GoTo 0
End Sub

' ============================================================================
' Function: IsDatabaseConnected
' Description: Check if database connection is alive
' Returns: True if connected
' ============================================================================
Public Function IsDatabaseConnected() As Boolean
    On Error GoTo ErrHandler
    
    If g_Connection Is Nothing Then
        IsDatabaseConnected = False
        Exit Function
    End If
    
    IsDatabaseConnected = (g_Connection.State = adStateOpen)
    
    Exit Function
    
ErrHandler:
    IsDatabaseConnected = False
End Function

' ============================================================================
' Function: EnsureConnection
' Description: Makes sure database is connected, reconnects if needed
' Returns: True if connected
' ============================================================================
Public Function EnsureConnection() As Boolean
    On Error GoTo ErrHandler
    
    If IsDatabaseConnected() Then
        EnsureConnection = True
        Exit Function
    End If
    
    ' Try to reconnect
    LogMessage "Database connection lost - attempting to reconnect..."
    m_ConnectionRetries = 0
    EnsureConnection = OpenDatabaseConnection()
    
    Exit Function
    
ErrHandler:
    LogMessage "EnsureConnection Error: " & Err.Description
    EnsureConnection = False
End Function

' ============================================================================
' Function: GetRecordset
' Description: Returns an ADO recordset for the given SQL query
' Parameters: sSQL - SQL query string
'             eCursorType - Optional cursor type (default: adOpenStatic)
'             eLockType - Optional lock type (default: adLockReadOnly)
' Returns: ADODB.Recordset or Nothing on error
' ============================================================================
Public Function GetRecordset(ByVal sSQL As String, _
                             Optional ByVal eCursorType As ADODB.CursorTypeEnum = adOpenStatic, _
                             Optional ByVal eLockType As ADODB.LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    On Error GoTo ErrHandler
    
    If Not EnsureConnection() Then
        Set GetRecordset = Nothing
        Exit Function
    End If
    
    Dim rs As New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open sSQL, g_Connection, eCursorType, eLockType
    
    Set GetRecordset = rs
    
    If g_EnableLogging Then
        LogMessage "GetRecordset: " & Left$(sSQL, 200)
    End If
    
    Exit Function
    
ErrHandler:
    LogMessage "GetRecordset Error #" & Err.Number & ": " & Err.Description & vbCrLf & "SQL: " & sSQL
    
    Set GetRecordset = Nothing
    
    ' Show error to user (VB6 anti-pattern - mixing UI in data layer)
    If Err.Number = -2147217900 Then ' Syntax error in SQL
        MsgBox "A database query error occurred." & vbCrLf & _
               "Error: " & Err.Description, vbExclamation, APP_TITLE
    ElseIf Err.Number = -2147467259 Then ' General DB error
        MsgBox "Database error occurred. The connection may have been lost." & vbCrLf & _
               "Error: " & Err.Description, vbExclamation, APP_TITLE
    End If
End Function

' ============================================================================
' Function: GetEditableRecordset
' Description: Returns an editable ADO recordset for updates
' Parameters: sSQL - SQL query string
' Returns: ADODB.Recordset with optimistic locking
' ============================================================================
Public Function GetEditableRecordset(ByVal sSQL As String) As ADODB.Recordset
    On Error GoTo ErrHandler
    
    Set GetEditableRecordset = GetRecordset(sSQL, adOpenKeyset, adLockOptimistic)
    
    Exit Function
    
ErrHandler:
    LogMessage "GetEditableRecordset Error: " & Err.Description
    Set GetEditableRecordset = Nothing
End Function

' ============================================================================
' Function: ExecuteSQL
' Description: Executes a non-query SQL statement (INSERT, UPDATE, DELETE)
' Parameters: sSQL - SQL statement to execute
' Returns: Number of records affected, or -1 on error
' ============================================================================
Public Function ExecuteSQL(ByVal sSQL As String) As Long
    On Error GoTo ErrHandler
    
    If Not EnsureConnection() Then
        ExecuteSQL = -1
        Exit Function
    End If
    
    Dim lRecordsAffected As Long
    
    g_Connection.Execute sSQL, lRecordsAffected, adCmdText
    
    ExecuteSQL = lRecordsAffected
    
    If g_EnableLogging Then
        LogMessage "ExecuteSQL (" & lRecordsAffected & " rows): " & Left$(sSQL, 200)
    End If
    
    Exit Function
    
ErrHandler:
    LogMessage "ExecuteSQL Error #" & Err.Number & ": " & Err.Description & vbCrLf & "SQL: " & sSQL
    
    MsgBox "Database error executing statement." & vbCrLf & _
           "Error: " & Err.Description, vbCritical, APP_TITLE
    
    ExecuteSQL = -1
End Function

' ============================================================================
' Function: GetScalarValue
' Description: Returns a single value from a SQL query
' Parameters: sSQL - SQL query that returns one value
'             vDefault - Default value if query returns nothing
' Returns: Variant containing the scalar value
' ============================================================================
Public Function GetScalarValue(ByVal sSQL As String, Optional ByVal vDefault As Variant = Null) As Variant
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    
    Set rs = GetRecordset(sSQL)
    
    If rs Is Nothing Then
        GetScalarValue = vDefault
        Exit Function
    End If
    
    If rs.EOF Then
        GetScalarValue = vDefault
    Else
        If IsNull(rs.Fields(0).Value) Then
            GetScalarValue = vDefault
        Else
            GetScalarValue = rs.Fields(0).Value
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
ErrHandler:
    LogMessage "GetScalarValue Error: " & Err.Description & " SQL: " & sSQL
    GetScalarValue = vDefault
    
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Function

' ============================================================================
' Function: GetNextID
' Description: Gets the next available ID for a table (auto-increment workaround)
' Parameters: sTable - Table name
'             sIDField - ID field name
' Returns: Next available ID number
' ============================================================================
Public Function GetNextID(ByVal sTable As String, ByVal sIDField As String) As Long
    On Error GoTo ErrHandler
    
    Dim vMaxID As Variant
    
    vMaxID = GetScalarValue("SELECT MAX(" & sIDField & ") FROM " & sTable, 0)
    
    If IsNull(vMaxID) Or IsEmpty(vMaxID) Then
        GetNextID = 1
    Else
        GetNextID = CLng(vMaxID) + 1
    End If
    
    Exit Function
    
ErrHandler:
    LogMessage "GetNextID Error: " & Err.Description
    GetNextID = -1
End Function

' ============================================================================
' Function: RecordExists
' Description: Checks if a record exists in a table
' Parameters: sTable - Table name
'             sWhereClause - WHERE clause without the WHERE keyword
' Returns: True if at least one record matches
' ============================================================================
Public Function RecordExists(ByVal sTable As String, ByVal sWhereClause As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim vCount As Variant
    
    vCount = GetScalarValue("SELECT COUNT(*) FROM " & sTable & " WHERE " & sWhereClause, 0)
    
    RecordExists = (CLng(vCount) > 0)
    
    Exit Function
    
ErrHandler:
    LogMessage "RecordExists Error: " & Err.Description
    RecordExists = False
End Function

' ============================================================================
' Sub: BeginTransaction
' Description: Begins a database transaction
' ============================================================================
Public Sub BeginTransaction()
    On Error GoTo ErrHandler
    
    If EnsureConnection() Then
        g_Connection.BeginTrans
        LogMessage "Transaction started"
    End If
    
    Exit Sub
    
ErrHandler:
    LogMessage "BeginTransaction Error: " & Err.Description
End Sub

' ============================================================================
' Sub: CommitTransaction
' Description: Commits the current transaction
' ============================================================================
Public Sub CommitTransaction()
    On Error GoTo ErrHandler
    
    If Not g_Connection Is Nothing Then
        g_Connection.CommitTrans
        LogMessage "Transaction committed"
    End If
    
    Exit Sub
    
ErrHandler:
    LogMessage "CommitTransaction Error: " & Err.Description
End Sub

' ============================================================================
' Sub: RollbackTransaction
' Description: Rolls back the current transaction
' ============================================================================
Public Sub RollbackTransaction()
    On Error GoTo ErrHandler
    
    If Not g_Connection Is Nothing Then
        g_Connection.RollbackTrans
        LogMessage "Transaction rolled back"
    End If
    
    Exit Sub
    
ErrHandler:
    LogMessage "RollbackTransaction Error: " & Err.Description
End Sub

' ============================================================================
' Sub: BackupDatabase
' Description: Creates a backup copy of the database file
' ============================================================================
Public Sub BackupDatabase()
    On Error GoTo ErrHandler
    
    Dim sBackupFile As String
    Dim sTimestamp As String
    
    sTimestamp = Format$(Now, "yyyymmdd_hhnnss")
    sBackupFile = g_BackupPath & "PrecisionParts_" & sTimestamp & ".mdb"
    
    ' Make sure backup folder exists
    If Dir(g_BackupPath, vbDirectory) = "" Then
        MkDir g_BackupPath
    End If
    
    ' Close connection before copying
    CloseDatabaseConnection
    
    ' Copy the file
    FileCopy g_DatabasePath, sBackupFile
    
    ' Reopen connection
    OpenDatabaseConnection
    
    MsgBox "Database backup created successfully:" & vbCrLf & sBackupFile, _
           vbInformation, APP_TITLE
    
    LogMessage "Database backup created: " & sBackupFile
    
    Exit Sub
    
ErrHandler:
    LogMessage "BackupDatabase Error: " & Err.Description
    MsgBox "Failed to create database backup." & vbCrLf & _
           "Error: " & Err.Description, vbCritical, APP_TITLE
    
    ' Try to reopen connection
    On Error Resume Next
    OpenDatabaseConnection
End Sub

' ============================================================================
' Sub: CompactDatabase
' Description: Compacts and repairs the database
' ============================================================================
Public Sub CompactDatabase()
    On Error GoTo ErrHandler
    
    Dim sTempFile As String
    
    sTempFile = g_DatabasePath & ".tmp"
    
    ' Close connection
    CloseDatabaseConnection
    
    ' Use JRO to compact
    Dim jro As Object
    Set jro = CreateObject("JRO.JetEngine")
    
    Dim sSrc As String
    Dim sDst As String
    sSrc = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_DatabasePath
    sDst = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sTempFile
    
    jro.CompactDatabase sSrc, sDst
    
    Set jro = Nothing
    
    ' Replace original with compacted
    Kill g_DatabasePath
    Name sTempFile As g_DatabasePath
    
    ' Reopen
    OpenDatabaseConnection
    
    MsgBox "Database compacted successfully.", vbInformation, APP_TITLE
    LogMessage "Database compacted"
    
    Exit Sub
    
ErrHandler:
    LogMessage "CompactDatabase Error: " & Err.Description
    MsgBox "Failed to compact database." & vbCrLf & _
           "Error: " & Err.Description, vbCritical, APP_TITLE
    
    On Error Resume Next
    If Dir(sTempFile) <> "" Then Kill sTempFile
    OpenDatabaseConnection
End Sub
