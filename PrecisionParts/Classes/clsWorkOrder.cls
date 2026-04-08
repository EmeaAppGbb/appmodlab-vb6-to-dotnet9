VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' Class: clsWorkOrder
' Description: Work order business logic and data access
' Created: January 2001
' Last Modified: March 2024
' Notes: Combines business logic with data access (VB6 anti-pattern)
' ============================================================================

Option Explicit

' Private member variables
Private m_WorkOrderID As Long
Private m_WorkOrderNumber As String
Private m_PartNumber As String
Private m_PartDescription As String
Private m_Quantity As Long
Private m_QuantityCompleted As Long
Private m_Status As String
Private m_Priority As String
Private m_DueDate As Date
Private m_StartDate As Date
Private m_CompletionDate As Date
Private m_CustomerPO As String
Private m_CustomerName As String
Private m_UnitCost As Double
Private m_MaterialCost As Double
Private m_LaborCost As Double
Private m_TotalCost As Double
Private m_Notes As String
Private m_CreatedBy As Long
Private m_CreatedDate As Date
Private m_ModifiedBy As Long
Private m_ModifiedDate As Date
Private m_IsNew As Boolean
Private m_IsDirty As Boolean
Private m_LastError As String

' ============================================================================
' Properties
' ============================================================================
Public Property Get WorkOrderID() As Long
    WorkOrderID = m_WorkOrderID
End Property

Public Property Let WorkOrderID(ByVal lValue As Long)
    m_WorkOrderID = lValue
    m_IsDirty = True
End Property

Public Property Get WorkOrderNumber() As String
    WorkOrderNumber = m_WorkOrderNumber
End Property

Public Property Let WorkOrderNumber(ByVal sValue As String)
    m_WorkOrderNumber = sValue
    m_IsDirty = True
End Property

Public Property Get PartNumber() As String
    PartNumber = m_PartNumber
End Property

Public Property Let PartNumber(ByVal sValue As String)
    m_PartNumber = Trim$(sValue)
    m_IsDirty = True
End Property

Public Property Get PartDescription() As String
    PartDescription = m_PartDescription
End Property

Public Property Let PartDescription(ByVal sValue As String)
    m_PartDescription = sValue
    m_IsDirty = True
End Property

Public Property Get Quantity() As Long
    Quantity = m_Quantity
End Property

Public Property Let Quantity(ByVal lValue As Long)
    m_Quantity = lValue
    m_IsDirty = True
End Property

Public Property Get QuantityCompleted() As Long
    QuantityCompleted = m_QuantityCompleted
End Property

Public Property Let QuantityCompleted(ByVal lValue As Long)
    m_QuantityCompleted = lValue
    m_IsDirty = True
End Property

Public Property Get Status() As String
    Status = m_Status
End Property

Public Property Let Status(ByVal sValue As String)
    m_Status = sValue
    m_IsDirty = True
End Property

Public Property Get Priority() As String
    Priority = m_Priority
End Property

Public Property Let Priority(ByVal sValue As String)
    m_Priority = sValue
    m_IsDirty = True
End Property

Public Property Get DueDate() As Date
    DueDate = m_DueDate
End Property

Public Property Let DueDate(ByVal dtValue As Date)
    m_DueDate = dtValue
    m_IsDirty = True
End Property

Public Property Get StartDate() As Date
    StartDate = m_StartDate
End Property

Public Property Let StartDate(ByVal dtValue As Date)
    m_StartDate = dtValue
    m_IsDirty = True
End Property

Public Property Get CompletionDate() As Date
    CompletionDate = m_CompletionDate
End Property

Public Property Let CompletionDate(ByVal dtValue As Date)
    m_CompletionDate = dtValue
    m_IsDirty = True
End Property

Public Property Get CustomerPO() As String
    CustomerPO = m_CustomerPO
End Property

Public Property Let CustomerPO(ByVal sValue As String)
    m_CustomerPO = sValue
    m_IsDirty = True
End Property

Public Property Get CustomerName() As String
    CustomerName = m_CustomerName
End Property

Public Property Let CustomerName(ByVal sValue As String)
    m_CustomerName = sValue
    m_IsDirty = True
End Property

Public Property Get UnitCost() As Double
    UnitCost = m_UnitCost
End Property

Public Property Let UnitCost(ByVal dValue As Double)
    m_UnitCost = dValue
    m_IsDirty = True
End Property

Public Property Get MaterialCost() As Double
    MaterialCost = m_MaterialCost
End Property

Public Property Let MaterialCost(ByVal dValue As Double)
    m_MaterialCost = dValue
    m_IsDirty = True
End Property

Public Property Get LaborCost() As Double
    LaborCost = m_LaborCost
End Property

Public Property Let LaborCost(ByVal dValue As Double)
    m_LaborCost = dValue
    m_IsDirty = True
End Property

Public Property Get TotalCost() As Double
    TotalCost = m_TotalCost
End Property

Public Property Get Notes() As String
    Notes = m_Notes
End Property

Public Property Let Notes(ByVal sValue As String)
    m_Notes = sValue
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
' Description: Set default values when class is created
' ============================================================================
Private Sub Class_Initialize()
    m_WorkOrderID = 0
    m_WorkOrderNumber = ""
    m_PartNumber = ""
    m_Quantity = 0
    m_QuantityCompleted = 0
    m_Status = WO_STATUS_NEW
    m_Priority = "Normal"
    m_DueDate = DateAdd("d", 14, Date)
    m_StartDate = Date
    m_CompletionDate = #1/1/1900#
    m_UnitCost = 0
    m_MaterialCost = 0
    m_LaborCost = 0
    m_TotalCost = 0
    m_IsNew = True
    m_IsDirty = False
    m_LastError = ""
End Sub

' ============================================================================
' Function: Load
' Description: Loads a work order from the database by ID
' Parameters: lWorkOrderID - ID of the work order to load
' Returns: True if loaded successfully
' ============================================================================
Public Function Load(ByVal lWorkOrderID As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT w.*, p.Description AS PartDesc " & _
           "FROM WorkOrders w " & _
           "LEFT JOIN Parts p ON w.PartNumber = p.PartNumber " & _
           "WHERE w.WorkOrderID = " & lWorkOrderID
    
    Set rs = GetRecordset(sSQL)
    
    If rs Is Nothing Then
        m_LastError = "Failed to query database for Work Order #" & lWorkOrderID
        Load = False
        Exit Function
    End If
    
    If rs.EOF Then
        m_LastError = "Work Order #" & lWorkOrderID & " was not found."
        rs.Close
        Set rs = Nothing
        Load = False
        Exit Function
    End If
    
    ' Populate properties from recordset
    m_WorkOrderID = SafeLong(rs!WorkOrderID)
    m_WorkOrderNumber = SafeString(rs!WorkOrderNumber)
    m_PartNumber = SafeString(rs!PartNumber)
    m_PartDescription = SafeString(rs!PartDesc)
    m_Quantity = SafeLong(rs!Quantity)
    m_QuantityCompleted = SafeLong(rs!QuantityCompleted)
    m_Status = SafeString(rs!Status)
    m_Priority = SafeString(rs!Priority)
    m_DueDate = SafeDate(rs!DueDate)
    m_StartDate = SafeDate(rs!StartDate)
    m_CompletionDate = SafeDate(rs!CompletionDate)
    m_CustomerPO = SafeString(rs!CustomerPO)
    m_CustomerName = SafeString(rs!CustomerName)
    m_UnitCost = SafeNumber(rs!UnitCost)
    m_MaterialCost = SafeNumber(rs!MaterialCost)
    m_LaborCost = SafeNumber(rs!LaborCost)
    m_Notes = SafeString(rs!Notes)
    m_CreatedBy = SafeLong(rs!CreatedBy)
    m_CreatedDate = SafeDate(rs!CreatedDate)
    m_ModifiedBy = SafeLong(rs!ModifiedBy)
    m_ModifiedDate = SafeDate(rs!ModifiedDate)
    
    ' Calculate total cost
    CalculateCost
    
    rs.Close
    Set rs = Nothing
    
    m_IsNew = False
    m_IsDirty = False
    
    Load = True
    
    LogMessage "Work Order #" & lWorkOrderID & " loaded"
    
    Exit Function
    
ErrHandler:
    m_LastError = "Error loading Work Order: " & Err.Description
    LogMessage m_LastError
    
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
    Load = False
End Function

' ============================================================================
' Function: Save
' Description: Saves the work order to the database (INSERT or UPDATE)
' Returns: True if saved successfully
' ============================================================================
Public Function Save() As Boolean
    On Error GoTo ErrHandler
    
    ' Validate before saving
    If Not Validate() Then
        Save = False
        Exit Function
    End If
    
    ' Calculate costs before saving
    CalculateCost
    
    Dim sSQL As String
    
    If m_IsNew Then
        ' Generate new ID and work order number
        m_WorkOrderID = GetNextID("WorkOrders", "WorkOrderID")
        
        If Len(m_WorkOrderNumber) = 0 Then
            m_WorkOrderNumber = GenerateWorkOrderNumber()
        End If
        
        m_CreatedBy = g_CurrentUserID
        m_CreatedDate = Now
        m_ModifiedBy = g_CurrentUserID
        m_ModifiedDate = Now
        
        sSQL = "INSERT INTO WorkOrders (WorkOrderID, WorkOrderNumber, PartNumber, " & _
               "Quantity, QuantityCompleted, Status, Priority, DueDate, StartDate, " & _
               "CompletionDate, CustomerPO, CustomerName, UnitCost, MaterialCost, " & _
               "LaborCost, Notes, CreatedBy, CreatedDate, ModifiedBy, ModifiedDate) " & _
               "VALUES (" & m_WorkOrderID & ", " & _
               "'" & SQLSafe(m_WorkOrderNumber) & "', " & _
               "'" & SQLSafe(m_PartNumber) & "', " & _
               m_Quantity & ", " & _
               m_QuantityCompleted & ", " & _
               "'" & SQLSafe(m_Status) & "', " & _
               "'" & SQLSafe(m_Priority) & "', " & _
               SQLDate(m_DueDate) & ", " & _
               SQLDate(m_StartDate) & ", " & _
               SQLDate(m_CompletionDate) & ", " & _
               "'" & SQLSafe(m_CustomerPO) & "', " & _
               "'" & SQLSafe(m_CustomerName) & "', " & _
               m_UnitCost & ", " & _
               m_MaterialCost & ", " & _
               m_LaborCost & ", " & _
               "'" & SQLSafe(m_Notes) & "', " & _
               m_CreatedBy & ", " & _
               SQLDateTime(m_CreatedDate) & ", " & _
               m_ModifiedBy & ", " & _
               SQLDateTime(m_ModifiedDate) & ")"
    Else
        m_ModifiedBy = g_CurrentUserID
        m_ModifiedDate = Now
        
        sSQL = "UPDATE WorkOrders SET " & _
               "WorkOrderNumber = '" & SQLSafe(m_WorkOrderNumber) & "', " & _
               "PartNumber = '" & SQLSafe(m_PartNumber) & "', " & _
               "Quantity = " & m_Quantity & ", " & _
               "QuantityCompleted = " & m_QuantityCompleted & ", " & _
               "Status = '" & SQLSafe(m_Status) & "', " & _
               "Priority = '" & SQLSafe(m_Priority) & "', " & _
               "DueDate = " & SQLDate(m_DueDate) & ", " & _
               "StartDate = " & SQLDate(m_StartDate) & ", " & _
               "CompletionDate = " & SQLDate(m_CompletionDate) & ", " & _
               "CustomerPO = '" & SQLSafe(m_CustomerPO) & "', " & _
               "CustomerName = '" & SQLSafe(m_CustomerName) & "', " & _
               "UnitCost = " & m_UnitCost & ", " & _
               "MaterialCost = " & m_MaterialCost & ", " & _
               "LaborCost = " & m_LaborCost & ", " & _
               "Notes = '" & SQLSafe(m_Notes) & "', " & _
               "ModifiedBy = " & m_ModifiedBy & ", " & _
               "ModifiedDate = " & SQLDateTime(m_ModifiedDate) & " " & _
               "WHERE WorkOrderID = " & m_WorkOrderID
    End If
    
    Dim lResult As Long
    lResult = ExecuteSQL(sSQL)
    
    If lResult < 0 Then
        m_LastError = "Failed to save Work Order to database."
        Save = False
        Exit Function
    End If
    
    m_IsNew = False
    m_IsDirty = False
    
    LogMessage "Work Order #" & m_WorkOrderID & " (" & m_WorkOrderNumber & ") saved"
    
    Save = True
    
    Exit Function
    
ErrHandler:
    m_LastError = "Error saving Work Order: " & Err.Description
    LogMessage m_LastError
    Save = False
End Function

' ============================================================================
' Function: Delete
' Description: Deletes this work order from the database
' Returns: True if deleted successfully
' ============================================================================
Public Function Delete() As Boolean
    On Error GoTo ErrHandler
    
    If m_IsNew Then
        m_LastError = "Cannot delete a work order that has not been saved."
        Delete = False
        Exit Function
    End If
    
    ' Check for related records
    If RecordExists("QualityChecks", "WorkOrderID = " & m_WorkOrderID) Then
        m_LastError = "Cannot delete - Quality Check records exist for this Work Order."
        MsgBox m_LastError, vbExclamation, APP_TITLE
        Delete = False
        Exit Function
    End If
    
    If RecordExists("ShippingManifests", "WorkOrderID = " & m_WorkOrderID) Then
        m_LastError = "Cannot delete - Shipping records exist for this Work Order."
        MsgBox m_LastError, vbExclamation, APP_TITLE
        Delete = False
        Exit Function
    End If
    
    Dim sSQL As String
    sSQL = "DELETE FROM WorkOrders WHERE WorkOrderID = " & m_WorkOrderID
    
    Dim lResult As Long
    lResult = ExecuteSQL(sSQL)
    
    If lResult < 0 Then
        m_LastError = "Failed to delete Work Order from database."
        Delete = False
        Exit Function
    End If
    
    LogMessage "Work Order #" & m_WorkOrderID & " deleted"
    
    ' Reset the object
    Class_Initialize
    
    Delete = True
    
    Exit Function
    
ErrHandler:
    m_LastError = "Error deleting Work Order: " & Err.Description
    LogMessage m_LastError
    Delete = False
End Function

' ============================================================================
' Function: Validate
' Description: Validates work order data before saving
' Returns: True if all data is valid
' ============================================================================
Public Function Validate() As Boolean
    On Error GoTo ErrHandler
    
    m_LastError = ""
    
    ' Part number is required
    If Len(Trim$(m_PartNumber)) = 0 Then
        m_LastError = "Part Number is required."
        Validate = False
        Exit Function
    End If
    
    ' Validate part number format
    If Not IsValidPartNumber(m_PartNumber) Then
        m_LastError = "Part Number format is invalid. Expected format: PP-XXXX-NNN"
        Validate = False
        Exit Function
    End If
    
    ' Quantity must be positive
    If m_Quantity <= 0 Then
        m_LastError = "Quantity must be greater than zero."
        Validate = False
        Exit Function
    End If
    
    ' Quantity completed cannot exceed quantity ordered
    If m_QuantityCompleted > m_Quantity Then
        m_LastError = "Quantity Completed cannot exceed Quantity Ordered."
        Validate = False
        Exit Function
    End If
    
    ' Due date must be set
    If m_DueDate < #1/1/2000# Then
        m_LastError = "Due Date must be a valid date."
        Validate = False
        Exit Function
    End If
    
    ' Start date must be before due date
    If m_StartDate > m_DueDate Then
        m_LastError = "Start Date cannot be after Due Date."
        Validate = False
        Exit Function
    End If
    
    ' Status must be valid
    Select Case m_Status
        Case WO_STATUS_NEW, WO_STATUS_INPROGRESS, WO_STATUS_COMPLETED, _
             WO_STATUS_CANCELLED, WO_STATUS_HOLD
            ' Valid
        Case Else
            m_LastError = "Invalid Work Order Status: " & m_Status
            Validate = False
            Exit Function
    End Select
    
    ' If completed, must have completion date
    If m_Status = WO_STATUS_COMPLETED Then
        If m_CompletionDate < #1/1/2000# Then
            m_CompletionDate = Now
        End If
    End If
    
    ' Check part exists in Parts table
    If Not RecordExists("Parts", "PartNumber = '" & SQLSafe(m_PartNumber) & "'") Then
        m_LastError = "Part Number '" & m_PartNumber & "' does not exist in the Parts table."
        Validate = False
        Exit Function
    End If
    
    Validate = True
    
    Exit Function
    
ErrHandler:
    m_LastError = "Validation error: " & Err.Description
    Validate = False
End Function

' ============================================================================
' Sub: CalculateCost
' Description: Calculates the total cost for this work order
' ============================================================================
Public Sub CalculateCost()
    On Error GoTo ErrHandler
    
    ' Look up material cost from Parts table
    Dim vUnitCost As Variant
    vUnitCost = GetScalarValue( _
        "SELECT UnitCost FROM Parts WHERE PartNumber = '" & SQLSafe(m_PartNumber) & "'", 0)
    
    If m_UnitCost = 0 Then
        m_UnitCost = SafeNumber(vUnitCost)
    End If
    
    ' Material cost = unit cost * quantity
    m_MaterialCost = m_UnitCost * m_Quantity
    
    ' Labor cost estimate: $45/hr, assume 0.5 hrs per part (hardcoded - VB6 anti-pattern)
    Dim dLaborRate As Double
    Dim dHoursPerPart As Double
    dLaborRate = 45#
    dHoursPerPart = 0.5
    
    m_LaborCost = m_Quantity * dHoursPerPart * dLaborRate
    
    ' Total cost
    m_TotalCost = m_MaterialCost + m_LaborCost
    
    ' Add 15% overhead (hardcoded - VB6 anti-pattern)
    m_TotalCost = m_TotalCost * 1.15
    
    Exit Sub
    
ErrHandler:
    m_TotalCost = 0
    LogMessage "CalculateCost Error: " & Err.Description
End Sub

' ============================================================================
' Function: GetPercentComplete
' Description: Returns the completion percentage
' ============================================================================
Public Function GetPercentComplete() As Double
    On Error Resume Next
    
    If m_Quantity = 0 Then
        GetPercentComplete = 0
    Else
        GetPercentComplete = (CDbl(m_QuantityCompleted) / CDbl(m_Quantity)) * 100
    End If
End Function

' ============================================================================
' Function: IsOverdue
' Description: Checks if the work order is past its due date
' ============================================================================
Public Function IsOverdue() As Boolean
    On Error Resume Next
    
    If m_Status = WO_STATUS_COMPLETED Or m_Status = WO_STATUS_CANCELLED Then
        IsOverdue = False
    Else
        IsOverdue = (Date > m_DueDate)
    End If
End Function

' ============================================================================
' Function: GetDaysRemaining
' Description: Returns number of days until due date
' ============================================================================
Public Function GetDaysRemaining() As Long
    On Error Resume Next
    GetDaysRemaining = DateDiff("d", Date, m_DueDate)
End Function
