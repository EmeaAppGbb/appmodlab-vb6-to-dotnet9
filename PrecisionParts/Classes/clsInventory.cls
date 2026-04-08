VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' Class: clsInventory
' Description: Inventory calculations and reorder point logic
' Created: February 2001
' Last Modified: March 2024
' Notes: Tightly coupled to database schema (VB6 anti-pattern)
' ============================================================================

Option Explicit

' Private member variables
Private m_MaterialID As Long
Private m_MaterialName As String
Private m_Supplier As String
Private m_QuantityOnHand As Long
Private m_QuantityOnOrder As Long
Private m_QuantityReserved As Long
Private m_ReorderPoint As Long
Private m_ReorderQuantity As Long
Private m_UnitCost As Double
Private m_UnitOfMeasure As String
Private m_Location As String
Private m_LastOrderDate As Date
Private m_LastReceivedDate As Date
Private m_LeadTimeDays As Integer
Private m_IsActive As Boolean
Private m_Notes As String
Private m_CreatedDate As Date
Private m_ModifiedDate As Date
Private m_IsNew As Boolean
Private m_IsDirty As Boolean
Private m_LastError As String

' Reorder status constants
Private Const STOCK_OK = "OK"
Private Const STOCK_LOW = "Low Stock"
Private Const STOCK_REORDER = "Reorder Required"
Private Const STOCK_CRITICAL = "CRITICAL - Out of Stock"

' ============================================================================
' Properties
' ============================================================================
Public Property Get MaterialID() As Long
    MaterialID = m_MaterialID
End Property

Public Property Let MaterialID(ByVal lValue As Long)
    m_MaterialID = lValue
    m_IsDirty = True
End Property

Public Property Get MaterialName() As String
    MaterialName = m_MaterialName
End Property

Public Property Let MaterialName(ByVal sValue As String)
    m_MaterialName = Trim$(sValue)
    m_IsDirty = True
End Property

Public Property Get Supplier() As String
    Supplier = m_Supplier
End Property

Public Property Let Supplier(ByVal sValue As String)
    m_Supplier = sValue
    m_IsDirty = True
End Property

Public Property Get QuantityOnHand() As Long
    QuantityOnHand = m_QuantityOnHand
End Property

Public Property Let QuantityOnHand(ByVal lValue As Long)
    m_QuantityOnHand = lValue
    m_IsDirty = True
End Property

Public Property Get QuantityOnOrder() As Long
    QuantityOnOrder = m_QuantityOnOrder
End Property

Public Property Let QuantityOnOrder(ByVal lValue As Long)
    m_QuantityOnOrder = lValue
    m_IsDirty = True
End Property

Public Property Get QuantityReserved() As Long
    QuantityReserved = m_QuantityReserved
End Property

Public Property Let QuantityReserved(ByVal lValue As Long)
    m_QuantityReserved = lValue
    m_IsDirty = True
End Property

Public Property Get ReorderPoint() As Long
    ReorderPoint = m_ReorderPoint
End Property

Public Property Let ReorderPoint(ByVal lValue As Long)
    m_ReorderPoint = lValue
    m_IsDirty = True
End Property

Public Property Get ReorderQuantity() As Long
    ReorderQuantity = m_ReorderQuantity
End Property

Public Property Let ReorderQuantity(ByVal lValue As Long)
    m_ReorderQuantity = lValue
    m_IsDirty = True
End Property

Public Property Get UnitCost() As Double
    UnitCost = m_UnitCost
End Property

Public Property Let UnitCost(ByVal dValue As Double)
    m_UnitCost = dValue
    m_IsDirty = True
End Property

Public Property Get UnitOfMeasure() As String
    UnitOfMeasure = m_UnitOfMeasure
End Property

Public Property Let UnitOfMeasure(ByVal sValue As String)
    m_UnitOfMeasure = sValue
    m_IsDirty = True
End Property

Public Property Get Location() As String
    Location = m_Location
End Property

Public Property Let Location(ByVal sValue As String)
    m_Location = sValue
    m_IsDirty = True
End Property

Public Property Get LastOrderDate() As Date
    LastOrderDate = m_LastOrderDate
End Property

Public Property Let LastOrderDate(ByVal dtValue As Date)
    m_LastOrderDate = dtValue
    m_IsDirty = True
End Property

Public Property Get LastReceivedDate() As Date
    LastReceivedDate = m_LastReceivedDate
End Property

Public Property Let LastReceivedDate(ByVal dtValue As Date)
    m_LastReceivedDate = dtValue
    m_IsDirty = True
End Property

Public Property Get LeadTimeDays() As Integer
    LeadTimeDays = m_LeadTimeDays
End Property

Public Property Let LeadTimeDays(ByVal iValue As Integer)
    m_LeadTimeDays = iValue
    m_IsDirty = True
End Property

Public Property Get IsActive() As Boolean
    IsActive = m_IsActive
End Property

Public Property Let IsActive(ByVal bValue As Boolean)
    m_IsActive = bValue
    m_IsDirty = True
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
' ============================================================================
Private Sub Class_Initialize()
    m_MaterialID = 0
    m_MaterialName = ""
    m_Supplier = ""
    m_QuantityOnHand = 0
    m_QuantityOnOrder = 0
    m_QuantityReserved = 0
    m_ReorderPoint = 100
    m_ReorderQuantity = 500
    m_UnitCost = 0
    m_UnitOfMeasure = "EA"
    m_Location = ""
    m_LastOrderDate = #1/1/1900#
    m_LastReceivedDate = #1/1/1900#
    m_LeadTimeDays = 7
    m_IsActive = True
    m_Notes = ""
    m_IsNew = True
    m_IsDirty = False
    m_LastError = ""
End Sub

' ============================================================================
' Function: Load
' Description: Loads an inventory item from the database
' Parameters: lMaterialID - ID of the material to load
' Returns: True if loaded successfully
' ============================================================================
Public Function Load(ByVal lMaterialID As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT * FROM RawMaterials WHERE MaterialID = " & lMaterialID
    
    Set rs = GetRecordset(sSQL)
    
    If rs Is Nothing Then
        m_LastError = "Failed to query database for Material #" & lMaterialID
        Load = False
        Exit Function
    End If
    
    If rs.EOF Then
        m_LastError = "Material #" & lMaterialID & " was not found."
        rs.Close
        Set rs = Nothing
        Load = False
        Exit Function
    End If
    
    ' Populate from recordset
    m_MaterialID = SafeLong(rs!MaterialID)
    m_MaterialName = SafeString(rs!MaterialName)
    m_Supplier = SafeString(rs!Supplier)
    m_QuantityOnHand = SafeLong(rs!QuantityOnHand)
    m_QuantityOnOrder = SafeLong(rs!QuantityOnOrder)
    m_QuantityReserved = SafeLong(rs!QuantityReserved)
    m_ReorderPoint = SafeLong(rs!ReorderPoint)
    m_ReorderQuantity = SafeLong(rs!ReorderQuantity)
    m_UnitCost = SafeNumber(rs!UnitCost)
    m_UnitOfMeasure = SafeString(rs!UnitOfMeasure)
    m_Location = SafeString(rs!Location)
    m_LastOrderDate = SafeDate(rs!LastOrderDate)
    m_LastReceivedDate = SafeDate(rs!LastReceivedDate)
    m_LeadTimeDays = CInt(SafeLong(rs!LeadTimeDays))
    m_IsActive = SafeBool(rs!IsActive)
    m_Notes = SafeString(rs!Notes)
    m_CreatedDate = SafeDate(rs!CreatedDate)
    m_ModifiedDate = SafeDate(rs!ModifiedDate)
    
    rs.Close
    Set rs = Nothing
    
    m_IsNew = False
    m_IsDirty = False
    
    Load = True
    
    Exit Function
    
ErrHandler:
    m_LastError = "Error loading Material: " & Err.Description
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
' Description: Saves the inventory item to the database
' Returns: True if saved successfully
' ============================================================================
Public Function Save() As Boolean
    On Error GoTo ErrHandler
    
    ' Validate
    If Len(Trim$(m_MaterialName)) = 0 Then
        m_LastError = "Material Name is required."
        Save = False
        Exit Function
    End If
    
    If m_ReorderPoint < 0 Then
        m_LastError = "Reorder Point cannot be negative."
        Save = False
        Exit Function
    End If
    
    If m_UnitCost < 0 Then
        m_LastError = "Unit Cost cannot be negative."
        Save = False
        Exit Function
    End If
    
    m_ModifiedDate = Now
    
    Dim sSQL As String
    
    If m_IsNew Then
        m_MaterialID = GetNextID("RawMaterials", "MaterialID")
        m_CreatedDate = Now
        
        sSQL = "INSERT INTO RawMaterials (MaterialID, MaterialName, Supplier, " & _
               "QuantityOnHand, QuantityOnOrder, QuantityReserved, ReorderPoint, " & _
               "ReorderQuantity, UnitCost, UnitOfMeasure, Location, " & _
               "LastOrderDate, LastReceivedDate, LeadTimeDays, IsActive, Notes, " & _
               "CreatedDate, ModifiedDate) VALUES (" & _
               m_MaterialID & ", " & _
               "'" & SQLSafe(m_MaterialName) & "', " & _
               "'" & SQLSafe(m_Supplier) & "', " & _
               m_QuantityOnHand & ", " & _
               m_QuantityOnOrder & ", " & _
               m_QuantityReserved & ", " & _
               m_ReorderPoint & ", " & _
               m_ReorderQuantity & ", " & _
               m_UnitCost & ", " & _
               "'" & SQLSafe(m_UnitOfMeasure) & "', " & _
               "'" & SQLSafe(m_Location) & "', " & _
               SQLDate(m_LastOrderDate) & ", " & _
               SQLDate(m_LastReceivedDate) & ", " & _
               m_LeadTimeDays & ", " & _
               IIf(m_IsActive, -1, 0) & ", " & _
               "'" & SQLSafe(m_Notes) & "', " & _
               SQLDateTime(m_CreatedDate) & ", " & _
               SQLDateTime(m_ModifiedDate) & ")"
    Else
        sSQL = "UPDATE RawMaterials SET " & _
               "MaterialName = '" & SQLSafe(m_MaterialName) & "', " & _
               "Supplier = '" & SQLSafe(m_Supplier) & "', " & _
               "QuantityOnHand = " & m_QuantityOnHand & ", " & _
               "QuantityOnOrder = " & m_QuantityOnOrder & ", " & _
               "QuantityReserved = " & m_QuantityReserved & ", " & _
               "ReorderPoint = " & m_ReorderPoint & ", " & _
               "ReorderQuantity = " & m_ReorderQuantity & ", " & _
               "UnitCost = " & m_UnitCost & ", " & _
               "UnitOfMeasure = '" & SQLSafe(m_UnitOfMeasure) & "', " & _
               "Location = '" & SQLSafe(m_Location) & "', " & _
               "LastOrderDate = " & SQLDate(m_LastOrderDate) & ", " & _
               "LastReceivedDate = " & SQLDate(m_LastReceivedDate) & ", " & _
               "LeadTimeDays = " & m_LeadTimeDays & ", " & _
               "IsActive = " & IIf(m_IsActive, -1, 0) & ", " & _
               "Notes = '" & SQLSafe(m_Notes) & "', " & _
               "ModifiedDate = " & SQLDateTime(m_ModifiedDate) & " " & _
               "WHERE MaterialID = " & m_MaterialID
    End If
    
    Dim lResult As Long
    lResult = ExecuteSQL(sSQL)
    
    If lResult < 0 Then
        m_LastError = "Failed to save material to database."
        Save = False
        Exit Function
    End If
    
    m_IsNew = False
    m_IsDirty = False
    
    LogMessage "Material #" & m_MaterialID & " (" & m_MaterialName & ") saved"
    
    Save = True
    
    Exit Function
    
ErrHandler:
    m_LastError = "Error saving Material: " & Err.Description
    LogMessage m_LastError
    Save = False
End Function

' ============================================================================
' Function: CheckReorderStatus
' Description: Checks if the material needs to be reordered
' Returns: Status string (OK, Low Stock, Reorder Required, CRITICAL)
' ============================================================================
Public Function CheckReorderStatus() As String
    On Error GoTo ErrHandler
    
    Dim lAvailable As Long
    
    ' Available = On Hand + On Order - Reserved
    lAvailable = m_QuantityOnHand + m_QuantityOnOrder - m_QuantityReserved
    
    If m_QuantityOnHand <= 0 Then
        CheckReorderStatus = STOCK_CRITICAL
    ElseIf lAvailable <= m_ReorderPoint Then
        CheckReorderStatus = STOCK_REORDER
    ElseIf lAvailable <= (m_ReorderPoint * 1.5) Then
        CheckReorderStatus = STOCK_LOW
    Else
        CheckReorderStatus = STOCK_OK
    End If
    
    Exit Function
    
ErrHandler:
    CheckReorderStatus = "Error"
    LogMessage "CheckReorderStatus Error: " & Err.Description
End Function

' ============================================================================
' Function: UpdateQuantity
' Description: Updates the quantity on hand (add or subtract)
' Parameters: lQuantityChange - Amount to add (positive) or subtract (negative)
'             sReason - Reason for the change
' Returns: True if updated successfully
' ============================================================================
Public Function UpdateQuantity(ByVal lQuantityChange As Long, ByVal sReason As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim lNewQty As Long
    lNewQty = m_QuantityOnHand + lQuantityChange
    
    ' Don't allow negative inventory
    If lNewQty < 0 Then
        m_LastError = "Insufficient quantity on hand. Available: " & m_QuantityOnHand & _
                      ", Requested: " & Abs(lQuantityChange)
        MsgBox m_LastError, vbExclamation, APP_TITLE
        UpdateQuantity = False
        Exit Function
    End If
    
    m_QuantityOnHand = lNewQty
    m_ModifiedDate = Now
    
    ' Update database directly (tight coupling - VB6 anti-pattern)
    Dim sSQL As String
    sSQL = "UPDATE RawMaterials SET QuantityOnHand = " & m_QuantityOnHand & _
           ", ModifiedDate = " & SQLDateTime(Now) & _
           " WHERE MaterialID = " & m_MaterialID
    
    Dim lResult As Long
    lResult = ExecuteSQL(sSQL)
    
    If lResult < 0 Then
        m_LastError = "Failed to update quantity in database."
        UpdateQuantity = False
        Exit Function
    End If
    
    ' Log the inventory transaction
    LogInventoryTransaction lQuantityChange, sReason
    
    ' Check if reorder is needed after update
    Dim sStatus As String
    sStatus = CheckReorderStatus()
    
    If sStatus = STOCK_REORDER Or sStatus = STOCK_CRITICAL Then
        MsgBox "WARNING: " & m_MaterialName & " - " & sStatus & vbCrLf & _
               "Current Qty: " & m_QuantityOnHand & vbCrLf & _
               "Reorder Point: " & m_ReorderPoint & vbCrLf & vbCrLf & _
               "Please place a purchase order.", vbExclamation, APP_TITLE
    End If
    
    UpdateQuantity = True
    
    Exit Function
    
ErrHandler:
    m_LastError = "Error updating quantity: " & Err.Description
    LogMessage m_LastError
    UpdateQuantity = False
End Function

' ============================================================================
' Sub: LogInventoryTransaction
' Description: Logs an inventory change to the transaction history
' ============================================================================
Private Sub LogInventoryTransaction(ByVal lQuantityChange As Long, ByVal sReason As String)
    On Error Resume Next
    
    Dim sSQL As String
    Dim lTransID As Long
    
    lTransID = GetNextID("InventoryTransactions", "TransactionID")
    
    sSQL = "INSERT INTO InventoryTransactions (TransactionID, MaterialID, " & _
           "QuantityChange, QuantityAfter, Reason, TransactionDate, UserID) " & _
           "VALUES (" & lTransID & ", " & m_MaterialID & ", " & _
           lQuantityChange & ", " & m_QuantityOnHand & ", " & _
           "'" & SQLSafe(sReason) & "', " & SQLDateTime(Now) & ", " & _
           g_CurrentUserID & ")"
    
    ExecuteSQL sSQL
End Sub

' ============================================================================
' Function: GetAvailableQuantity
' Description: Returns the available quantity (on hand - reserved)
' ============================================================================
Public Function GetAvailableQuantity() As Long
    On Error Resume Next
    GetAvailableQuantity = m_QuantityOnHand - m_QuantityReserved
    If GetAvailableQuantity < 0 Then GetAvailableQuantity = 0
End Function

' ============================================================================
' Function: GetTotalValue
' Description: Returns total inventory value (qty * unit cost)
' ============================================================================
Public Function GetTotalValue() As Double
    On Error Resume Next
    GetTotalValue = m_QuantityOnHand * m_UnitCost
End Function

' ============================================================================
' Function: GetExpectedDeliveryDate
' Description: Returns expected delivery date based on lead time
' ============================================================================
Public Function GetExpectedDeliveryDate() As Date
    On Error Resume Next
    If m_LastOrderDate > #1/1/2000# Then
        GetExpectedDeliveryDate = DateAdd("d", m_LeadTimeDays, m_LastOrderDate)
    Else
        GetExpectedDeliveryDate = DateAdd("d", m_LeadTimeDays, Date)
    End If
End Function

' ============================================================================
' Function: NeedsReorder
' Description: Simple boolean check for reorder
' ============================================================================
Public Function NeedsReorder() As Boolean
    On Error Resume Next
    Dim sStatus As String
    sStatus = CheckReorderStatus()
    NeedsReorder = (sStatus = STOCK_REORDER Or sStatus = STOCK_CRITICAL)
End Function
