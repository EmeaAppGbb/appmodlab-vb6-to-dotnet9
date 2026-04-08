VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' Class: clsPart
' Description: Part number validation and data access
' Created: February 2001
' Last Modified: March 2024
' ============================================================================

Option Explicit

' Private member variables
Private m_PartID As Long
Private m_PartNumber As String
Private m_Description As String
Private m_Material As String
Private m_MaterialSpec As String
Private m_UnitCost As Double
Private m_Weight As Double
Private m_DrawingNumber As String
Private m_RevisionLevel As String
Private m_Category As String
Private m_IsActive As Boolean
Private m_CreatedDate As Date
Private m_ModifiedDate As Date
Private m_IsNew As Boolean
Private m_IsDirty As Boolean
Private m_LastError As String

' ============================================================================
' Properties
' ============================================================================
Public Property Get PartID() As Long
    PartID = m_PartID
End Property

Public Property Let PartID(ByVal lValue As Long)
    m_PartID = lValue
    m_IsDirty = True
End Property

Public Property Get PartNumber() As String
    PartNumber = m_PartNumber
End Property

Public Property Let PartNumber(ByVal sValue As String)
    m_PartNumber = UCase$(Trim$(sValue))
    m_IsDirty = True
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(ByVal sValue As String)
    m_Description = sValue
    m_IsDirty = True
End Property

Public Property Get Material() As String
    Material = m_Material
End Property

Public Property Let Material(ByVal sValue As String)
    m_Material = sValue
    m_IsDirty = True
End Property

Public Property Get MaterialSpec() As String
    MaterialSpec = m_MaterialSpec
End Property

Public Property Let MaterialSpec(ByVal sValue As String)
    m_MaterialSpec = sValue
    m_IsDirty = True
End Property

Public Property Get UnitCost() As Double
    UnitCost = m_UnitCost
End Property

Public Property Let UnitCost(ByVal dValue As Double)
    m_UnitCost = dValue
    m_IsDirty = True
End Property

Public Property Get Weight() As Double
    Weight = m_Weight
End Property

Public Property Let Weight(ByVal dValue As Double)
    m_Weight = dValue
    m_IsDirty = True
End Property

Public Property Get DrawingNumber() As String
    DrawingNumber = m_DrawingNumber
End Property

Public Property Let DrawingNumber(ByVal sValue As String)
    m_DrawingNumber = sValue
    m_IsDirty = True
End Property

Public Property Get RevisionLevel() As String
    RevisionLevel = m_RevisionLevel
End Property

Public Property Let RevisionLevel(ByVal sValue As String)
    m_RevisionLevel = sValue
    m_IsDirty = True
End Property

Public Property Get Category() As String
    Category = m_Category
End Property

Public Property Let Category(ByVal sValue As String)
    m_Category = sValue
    m_IsDirty = True
End Property

Public Property Get IsActive() As Boolean
    IsActive = m_IsActive
End Property

Public Property Let IsActive(ByVal bValue As Boolean)
    m_IsActive = bValue
    m_IsDirty = True
End Property

Public Property Get IsNew() As Boolean
    IsNew = m_IsNew
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = m_IsDirty
End Property

Public Property Get LastError() As String
    LastError = m_LastError
End Property

' ============================================================================
' Sub: Class_Initialize
' ============================================================================
Private Sub Class_Initialize()
    m_PartID = 0
    m_PartNumber = ""
    m_Description = ""
    m_Material = ""
    m_MaterialSpec = ""
    m_UnitCost = 0
    m_Weight = 0
    m_DrawingNumber = ""
    m_RevisionLevel = "A"
    m_Category = ""
    m_IsActive = True
    m_IsNew = True
    m_IsDirty = False
    m_LastError = ""
End Sub

' ============================================================================
' Function: Load
' Description: Loads a part from the database
' Parameters: sPartNumber - Part number to load
' Returns: True if loaded successfully
' ============================================================================
Public Function Load(ByVal sPartNumber As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT * FROM Parts WHERE PartNumber = '" & SQLSafe(sPartNumber) & "'"
    
    Set rs = GetRecordset(sSQL)
    
    If rs Is Nothing Then
        m_LastError = "Failed to query database for Part '" & sPartNumber & "'"
        Load = False
        Exit Function
    End If
    
    If rs.EOF Then
        m_LastError = "Part '" & sPartNumber & "' was not found."
        rs.Close
        Set rs = Nothing
        Load = False
        Exit Function
    End If
    
    m_PartID = SafeLong(rs!PartID)
    m_PartNumber = SafeString(rs!PartNumber)
    m_Description = SafeString(rs!Description)
    m_Material = SafeString(rs!Material)
    m_MaterialSpec = SafeString(rs!MaterialSpec)
    m_UnitCost = SafeNumber(rs!UnitCost)
    m_Weight = SafeNumber(rs!Weight)
    m_DrawingNumber = SafeString(rs!DrawingNumber)
    m_RevisionLevel = SafeString(rs!RevisionLevel)
    m_Category = SafeString(rs!Category)
    m_IsActive = SafeBool(rs!IsActive)
    m_CreatedDate = SafeDate(rs!CreatedDate)
    m_ModifiedDate = SafeDate(rs!ModifiedDate)
    
    rs.Close
    Set rs = Nothing
    
    m_IsNew = False
    m_IsDirty = False
    
    Load = True
    
    Exit Function
    
ErrHandler:
    m_LastError = "Error loading Part: " & Err.Description
    LogMessage m_LastError
    
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
    Load = False
End Function

' ============================================================================
' Function: LoadByID
' Description: Loads a part from the database by ID
' Parameters: lPartID - ID of the part to load
' Returns: True if loaded successfully
' ============================================================================
Public Function LoadByID(ByVal lPartID As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT * FROM Parts WHERE PartID = " & lPartID
    
    Set rs = GetRecordset(sSQL)
    
    If rs Is Nothing Or rs.EOF Then
        m_LastError = "Part ID #" & lPartID & " was not found."
        If Not rs Is Nothing Then rs.Close
        Set rs = Nothing
        LoadByID = False
        Exit Function
    End If
    
    m_PartID = SafeLong(rs!PartID)
    m_PartNumber = SafeString(rs!PartNumber)
    m_Description = SafeString(rs!Description)
    m_Material = SafeString(rs!Material)
    m_MaterialSpec = SafeString(rs!MaterialSpec)
    m_UnitCost = SafeNumber(rs!UnitCost)
    m_Weight = SafeNumber(rs!Weight)
    m_DrawingNumber = SafeString(rs!DrawingNumber)
    m_RevisionLevel = SafeString(rs!RevisionLevel)
    m_Category = SafeString(rs!Category)
    m_IsActive = SafeBool(rs!IsActive)
    m_CreatedDate = SafeDate(rs!CreatedDate)
    m_ModifiedDate = SafeDate(rs!ModifiedDate)
    
    rs.Close
    Set rs = Nothing
    
    m_IsNew = False
    m_IsDirty = False
    
    LoadByID = True
    
    Exit Function
    
ErrHandler:
    m_LastError = "Error loading Part by ID: " & Err.Description
    LogMessage m_LastError
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    LoadByID = False
End Function

' ============================================================================
' Function: Save
' Description: Saves the part to the database
' Returns: True if saved successfully
' ============================================================================
Public Function Save() As Boolean
    On Error GoTo ErrHandler
    
    If Not ValidatePartNumber(m_PartNumber) Then
        Save = False
        Exit Function
    End If
    
    If Len(Trim$(m_Description)) = 0 Then
        m_LastError = "Part Description is required."
        Save = False
        Exit Function
    End If
    
    m_ModifiedDate = Now
    
    Dim sSQL As String
    
    If m_IsNew Then
        ' Check for duplicate part number
        If RecordExists("Parts", "PartNumber = '" & SQLSafe(m_PartNumber) & "'") Then
            m_LastError = "Part Number '" & m_PartNumber & "' already exists."
            Save = False
            Exit Function
        End If
        
        m_PartID = GetNextID("Parts", "PartID")
        m_CreatedDate = Now
        
        sSQL = "INSERT INTO Parts (PartID, PartNumber, Description, Material, " & _
               "MaterialSpec, UnitCost, Weight, DrawingNumber, RevisionLevel, " & _
               "Category, IsActive, CreatedDate, ModifiedDate) VALUES (" & _
               m_PartID & ", " & _
               "'" & SQLSafe(m_PartNumber) & "', " & _
               "'" & SQLSafe(m_Description) & "', " & _
               "'" & SQLSafe(m_Material) & "', " & _
               "'" & SQLSafe(m_MaterialSpec) & "', " & _
               m_UnitCost & ", " & _
               m_Weight & ", " & _
               "'" & SQLSafe(m_DrawingNumber) & "', " & _
               "'" & SQLSafe(m_RevisionLevel) & "', " & _
               "'" & SQLSafe(m_Category) & "', " & _
               IIf(m_IsActive, -1, 0) & ", " & _
               SQLDateTime(m_CreatedDate) & ", " & _
               SQLDateTime(m_ModifiedDate) & ")"
    Else
        sSQL = "UPDATE Parts SET " & _
               "PartNumber = '" & SQLSafe(m_PartNumber) & "', " & _
               "Description = '" & SQLSafe(m_Description) & "', " & _
               "Material = '" & SQLSafe(m_Material) & "', " & _
               "MaterialSpec = '" & SQLSafe(m_MaterialSpec) & "', " & _
               "UnitCost = " & m_UnitCost & ", " & _
               "Weight = " & m_Weight & ", " & _
               "DrawingNumber = '" & SQLSafe(m_DrawingNumber) & "', " & _
               "RevisionLevel = '" & SQLSafe(m_RevisionLevel) & "', " & _
               "Category = '" & SQLSafe(m_Category) & "', " & _
               "IsActive = " & IIf(m_IsActive, -1, 0) & ", " & _
               "ModifiedDate = " & SQLDateTime(m_ModifiedDate) & " " & _
               "WHERE PartID = " & m_PartID
    End If
    
    Dim lResult As Long
    lResult = ExecuteSQL(sSQL)
    
    If lResult < 0 Then
        m_LastError = "Failed to save Part to database."
        Save = False
        Exit Function
    End If
    
    m_IsNew = False
    m_IsDirty = False
    
    LogMessage "Part '" & m_PartNumber & "' saved"
    
    Save = True
    
    Exit Function
    
ErrHandler:
    m_LastError = "Error saving Part: " & Err.Description
    LogMessage m_LastError
    Save = False
End Function

' ============================================================================
' Function: ValidatePartNumber
' Description: Validates a part number format and checks business rules
' Parameters: sPartNumber - Part number to validate
' Returns: True if valid
' ============================================================================
Public Function ValidatePartNumber(ByVal sPartNumber As String) As Boolean
    On Error GoTo ErrHandler
    
    m_LastError = ""
    
    ' Must not be empty
    If Len(Trim$(sPartNumber)) = 0 Then
        m_LastError = "Part Number is required."
        ValidatePartNumber = False
        Exit Function
    End If
    
    ' Convert to uppercase for validation
    sPartNumber = UCase$(Trim$(sPartNumber))
    
    ' Must contain at least one hyphen
    If InStr(sPartNumber, "-") = 0 Then
        m_LastError = "Part Number must contain at least one hyphen (e.g., PP-1234-001)."
        ValidatePartNumber = False
        Exit Function
    End If
    
    ' Split into segments
    Dim sParts() As String
    sParts = Split(sPartNumber, "-")
    
    ' Must have 2-4 segments
    If UBound(sParts) < 1 Or UBound(sParts) > 3 Then
        m_LastError = "Part Number must have 2 to 4 segments separated by hyphens."
        ValidatePartNumber = False
        Exit Function
    End If
    
    ' First segment must be 2-4 alpha characters (category prefix)
    Dim sPrefix As String
    sPrefix = sParts(0)
    
    If Len(sPrefix) < 2 Or Len(sPrefix) > 4 Then
        m_LastError = "Part Number prefix must be 2-4 characters."
        ValidatePartNumber = False
        Exit Function
    End If
    
    Dim i As Integer
    Dim sChar As String
    For i = 1 To Len(sPrefix)
        sChar = Mid$(sPrefix, i, 1)
        If sChar < "A" Or sChar > "Z" Then
            m_LastError = "Part Number prefix must contain only letters."
            ValidatePartNumber = False
            Exit Function
        End If
    Next i
    
    ' Validate known prefixes (business rule - hardcoded)
    Select Case sPrefix
        Case "PP", "PPM", "PPA", "MFG", "ASM", "RAW", "HDW", "ELC"
            ' Valid prefix
        Case Else
            ' Allow but warn
            LogMessage "WARNING: Unknown part number prefix: " & sPrefix
    End Select
    
    ' Second segment should be numeric (4-6 digits)
    Dim sNumber As String
    sNumber = sParts(1)
    
    If Len(sNumber) < 1 Or Len(sNumber) > 8 Then
        m_LastError = "Part Number numeric segment must be 1-8 characters."
        ValidatePartNumber = False
        Exit Function
    End If
    
    ' Total length check
    If Len(sPartNumber) > 20 Then
        m_LastError = "Part Number cannot exceed 20 characters."
        ValidatePartNumber = False
        Exit Function
    End If
    
    ValidatePartNumber = True
    
    Exit Function
    
ErrHandler:
    m_LastError = "Part Number validation error: " & Err.Description
    ValidatePartNumber = False
End Function

' ============================================================================
' Function: Delete
' Description: Deletes this part from the database
' Returns: True if deleted
' ============================================================================
Public Function Delete() As Boolean
    On Error GoTo ErrHandler
    
    If m_IsNew Then
        m_LastError = "Cannot delete a part that has not been saved."
        Delete = False
        Exit Function
    End If
    
    ' Check for work orders using this part
    If RecordExists("WorkOrders", "PartNumber = '" & SQLSafe(m_PartNumber) & "'") Then
        m_LastError = "Cannot delete - Work Orders exist for this Part Number."
        MsgBox m_LastError, vbExclamation, APP_TITLE
        Delete = False
        Exit Function
    End If
    
    Dim sSQL As String
    sSQL = "DELETE FROM Parts WHERE PartID = " & m_PartID
    
    Dim lResult As Long
    lResult = ExecuteSQL(sSQL)
    
    If lResult < 0 Then
        m_LastError = "Failed to delete Part from database."
        Delete = False
        Exit Function
    End If
    
    LogMessage "Part '" & m_PartNumber & "' deleted"
    
    Class_Initialize
    
    Delete = True
    
    Exit Function
    
ErrHandler:
    m_LastError = "Error deleting Part: " & Err.Description
    LogMessage m_LastError
    Delete = False
End Function

' ============================================================================
' Function: GetDisplayString
' Description: Returns a formatted string for display in lists
' ============================================================================
Public Function GetDisplayString() As String
    GetDisplayString = m_PartNumber & " - " & m_Description
    If Len(m_Material) > 0 Then
        GetDisplayString = GetDisplayString & " [" & m_Material & "]"
    End If
End Function
