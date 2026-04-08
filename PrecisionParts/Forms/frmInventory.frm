VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInventory 
   Caption         =   "Raw Material Inventory"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   11520
   Begin VB.ComboBox cboSupplier 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "E&xport"
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdInventory 
      Height          =   6375
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0   'None
      ScrollBars      =   2
      SelectionMode   =   1   'By Row
      AllowUserResizing=   1   'Columns
   End
   Begin VB.Label lblFilter 
      Caption         =   "Supplier:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   160
      Width           =   975
   End
   Begin VB.Label lblRecordCount 
      Alignment       =   2  'Center
      Caption         =   "0 Records"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   11175
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' Form: frmInventory
' Description: Raw material inventory management grid
' Created: February 2001
' Last Modified: March 2024
' ============================================================================

Option Explicit

Private m_SortColumn As Integer
Private m_SortAscending As Boolean

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    m_SortColumn = 0
    m_SortAscending = True
    
    ' Setup grid columns
    SetupGrid
    
    ' Load supplier list for filter
    LoadSupplierList
    
    ' Load inventory data
    LoadInventoryData
    
    Exit Sub
    
ErrHandler:
    ShowError "frmInventory.Form_Load", Err.Description, Err.Number
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' Resize grid to fill form
    grdInventory.Width = Me.ScaleWidth - 240
    grdInventory.Height = Me.ScaleHeight - grdInventory.Top - 400
    
    lblRecordCount.Top = Me.ScaleHeight - 300
    lblRecordCount.Width = Me.ScaleWidth - 240
End Sub

' ============================================================================
' Sub: SetupGrid
' Description: Configures MSFlexGrid columns and headers
' ============================================================================
Private Sub SetupGrid()
    On Error GoTo ErrHandler
    
    With grdInventory
        .Cols = 8
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        
        ' Set column widths
        .ColWidth(0) = 900     ' Material ID
        .ColWidth(1) = 2400    ' Material Name
        .ColWidth(2) = 1800    ' Supplier
        .ColWidth(3) = 1200    ' Qty On Hand
        .ColWidth(4) = 1200    ' Reorder Point
        .ColWidth(5) = 1100    ' Unit Cost
        .ColWidth(6) = 1200    ' Location
        .ColWidth(7) = 1200    ' Status
        
        ' Set column headers
        .Row = 0
        .Col = 0: .Text = "ID"
        .Col = 1: .Text = "Material Name"
        .Col = 2: .Text = "Supplier"
        .Col = 3: .Text = "Qty On Hand"
        .Col = 4: .Text = "Reorder Pt"
        .Col = 5: .Text = "Unit Cost"
        .Col = 6: .Text = "Location"
        .Col = 7: .Text = "Status"
        
        ' Set column alignment
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignRightCenter
    End With
    
    Exit Sub
    
ErrHandler:
    ShowError "SetupGrid", Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: LoadSupplierList
' Description: Populates the supplier filter combo box
' ============================================================================
Private Sub LoadSupplierList()
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    cboSupplier.Clear
    cboSupplier.AddItem "(All Suppliers)"
    
    sSQL = "SELECT DISTINCT Supplier FROM RawMaterials WHERE Supplier IS NOT NULL ORDER BY Supplier"
    
    Set rs = GetRecordset(sSQL)
    
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            If Len(SafeString(rs!Supplier)) > 0 Then
                cboSupplier.AddItem SafeString(rs!Supplier)
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    
    cboSupplier.ListIndex = 0
    
    Exit Sub
    
ErrHandler:
    ShowError "LoadSupplierList", Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: LoadInventoryData
' Description: Loads inventory data into the MSFlexGrid
' ============================================================================
Private Sub LoadInventoryData()
    On Error GoTo ErrHandler
    
    SetWaitCursor True
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim lRow As Long
    
    sSQL = "SELECT MaterialID, MaterialName, Supplier, QuantityOnHand, " & _
           "ReorderPoint, UnitCost, Location " & _
           "FROM RawMaterials"
    
    ' Apply supplier filter
    If cboSupplier.ListIndex > 0 Then
        sSQL = sSQL & " WHERE Supplier = '" & SQLSafe(cboSupplier.Text) & "'"
    End If
    
    sSQL = sSQL & " ORDER BY MaterialName"
    
    Set rs = GetRecordset(sSQL)
    
    If rs Is Nothing Then
        MsgBox "Failed to load inventory data.", vbExclamation, APP_TITLE
        SetWaitCursor False
        Exit Sub
    End If
    
    ' Clear grid (keep header row)
    grdInventory.Rows = 1
    
    If rs.EOF Then
        lblRecordCount.Caption = "0 Records"
        rs.Close
        Set rs = Nothing
        SetWaitCursor False
        Exit Sub
    End If
    
    ' Populate grid
    lRow = 0
    Do While Not rs.EOF
        lRow = lRow + 1
        grdInventory.Rows = lRow + 1
        
        Dim lQty As Long
        Dim lReorder As Long
        Dim sStatus As String
        
        lQty = SafeLong(rs!QuantityOnHand)
        lReorder = SafeLong(rs!ReorderPoint)
        
        ' Determine stock status
        If lQty <= 0 Then
            sStatus = "CRITICAL"
        ElseIf lQty <= lReorder Then
            sStatus = "REORDER"
        ElseIf lQty <= lReorder * 1.5 Then
            sStatus = "Low"
        Else
            sStatus = "OK"
        End If
        
        grdInventory.Row = lRow
        grdInventory.Col = 0: grdInventory.Text = SafeString(rs!MaterialID)
        grdInventory.Col = 1: grdInventory.Text = SafeString(rs!MaterialName)
        grdInventory.Col = 2: grdInventory.Text = SafeString(rs!Supplier)
        grdInventory.Col = 3: grdInventory.Text = FormatQuantity(lQty)
        grdInventory.Col = 4: grdInventory.Text = FormatQuantity(lReorder)
        grdInventory.Col = 5: grdInventory.Text = FormatCurrencyValue(SafeNumber(rs!UnitCost))
        grdInventory.Col = 6: grdInventory.Text = SafeString(rs!Location)
        grdInventory.Col = 7: grdInventory.Text = sStatus
        
        ' Color code rows based on status (VB6 cell coloring)
        If sStatus = "CRITICAL" Then
            grdInventory.CellBackColor = &HC0C0FF    ' Light red
        ElseIf sStatus = "REORDER" Then
            grdInventory.CellBackColor = &HC0FFFF    ' Light yellow
        End If
        
        rs.MoveNext
        
        ' Allow UI updates during large loads (VB6 anti-pattern)
        If lRow Mod 50 = 0 Then DoEvents
    Loop
    
    lblRecordCount.Caption = lRow & " Records"
    
    rs.Close
    Set rs = Nothing
    
    ' Select first data row
    If grdInventory.Rows > 1 Then
        grdInventory.Row = 1
    End If
    
    SetWaitCursor False
    
    Exit Sub
    
ErrHandler:
    SetWaitCursor False
    ShowError "LoadInventoryData", Err.Description, Err.Number
    
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Sub

' ============================================================================
' Grid Events
' ============================================================================
Private Sub grdInventory_Click()
    On Error Resume Next
    ' Ensure full row is selected
    grdInventory.ColSel = grdInventory.Cols - 1
End Sub

Private Sub grdInventory_DblClick()
    ' Double-click to edit
    cmdEdit_Click
End Sub

Private Sub grdInventory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrHandler
    
    ' Check if clicked on header row for sorting
    If grdInventory.MouseRow = 0 Then
        Dim iClickedCol As Integer
        iClickedCol = grdInventory.MouseCol
        
        If iClickedCol = m_SortColumn Then
            m_SortAscending = Not m_SortAscending
        Else
            m_SortColumn = iClickedCol
            m_SortAscending = True
        End If
        
        SortGrid m_SortColumn, m_SortAscending
    End If
    
    Exit Sub
    
ErrHandler:
    ' Ignore sort errors
End Sub

' ============================================================================
' Sub: SortGrid
' Description: Sorts the grid by the specified column (bubble sort)
' ============================================================================
Private Sub SortGrid(ByVal iCol As Integer, ByVal bAscending As Boolean)
    On Error GoTo ErrHandler
    
    If grdInventory.Rows <= 2 Then Exit Sub
    
    SetWaitCursor True
    
    Dim i As Long, j As Long
    Dim sTemp As String
    Dim bSwap As Boolean
    
    ' Simple bubble sort (VB6 anti-pattern - inefficient)
    For i = 1 To grdInventory.Rows - 2
        For j = 1 To grdInventory.Rows - i - 1
            
            grdInventory.Row = j
            grdInventory.Col = iCol
            Dim sVal1 As String
            sVal1 = grdInventory.Text
            
            grdInventory.Row = j + 1
            grdInventory.Col = iCol
            Dim sVal2 As String
            sVal2 = grdInventory.Text
            
            If bAscending Then
                bSwap = (sVal1 > sVal2)
            Else
                bSwap = (sVal1 < sVal2)
            End If
            
            If bSwap Then
                ' Swap entire rows
                Dim k As Integer
                For k = 0 To grdInventory.Cols - 1
                    grdInventory.Row = j
                    grdInventory.Col = k
                    sTemp = grdInventory.Text
                    
                    grdInventory.Row = j + 1
                    grdInventory.Col = k
                    grdInventory.Row = j
                    grdInventory.Text = grdInventory.Text
                    
                    ' Actually swap - need temp storage per cell
                    Dim sRowJ As String
                    Dim sRowJ1 As String
                    
                    grdInventory.Row = j
                    grdInventory.Col = k
                    sRowJ = grdInventory.Text
                    
                    grdInventory.Row = j + 1
                    grdInventory.Col = k
                    sRowJ1 = grdInventory.Text
                    
                    grdInventory.Row = j
                    grdInventory.Col = k
                    grdInventory.Text = sRowJ1
                    
                    grdInventory.Row = j + 1
                    grdInventory.Col = k
                    grdInventory.Text = sRowJ
                Next k
            End If
            
            DoEvents
        Next j
    Next i
    
    SetWaitCursor False
    
    Exit Sub
    
ErrHandler:
    SetWaitCursor False
End Sub

' ============================================================================
' Button Event Handlers
' ============================================================================
Private Sub cmdAdd_Click()
    On Error GoTo ErrHandler
    
    Dim sName As String
    Dim sSupplier As String
    Dim sQty As String
    Dim sReorder As String
    Dim sCost As String
    Dim sLocation As String
    
    ' Simple input via InputBox chain (VB6 anti-pattern - should use a dialog)
    sName = InputBox("Enter Material Name:", APP_TITLE)
    If Len(sName) = 0 Then Exit Sub
    
    sSupplier = InputBox("Enter Supplier Name:", APP_TITLE)
    If Len(sSupplier) = 0 Then Exit Sub
    
    sQty = InputBox("Enter Quantity On Hand:", APP_TITLE, "0")
    If Len(sQty) = 0 Then Exit Sub
    If Not IsNumeric(sQty) Then
        MsgBox "Invalid quantity.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    sReorder = InputBox("Enter Reorder Point:", APP_TITLE, "100")
    If Len(sReorder) = 0 Then Exit Sub
    
    sCost = InputBox("Enter Unit Cost:", APP_TITLE, "0.00")
    If Len(sCost) = 0 Then Exit Sub
    
    sLocation = InputBox("Enter Storage Location:", APP_TITLE, "Warehouse A")
    
    ' Save using clsInventory
    Dim objInv As New clsInventory
    objInv.MaterialName = sName
    objInv.Supplier = sSupplier
    objInv.QuantityOnHand = CLng(sQty)
    objInv.ReorderPoint = CLng(sReorder)
    objInv.UnitCost = CDbl(sCost)
    objInv.Location = sLocation
    
    If objInv.Save() Then
        MsgBox "Material added successfully.", vbInformation, APP_TITLE
        LoadInventoryData
        LoadSupplierList
    Else
        MsgBox "Failed to add material: " & objInv.LastError, vbExclamation, APP_TITLE
    End If
    
    Set objInv = Nothing
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdAdd_Click", Err.Description, Err.Number
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo ErrHandler
    
    If grdInventory.Rows <= 1 Then
        MsgBox "No records to edit.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    ' Get selected material ID
    Dim lMaterialID As Long
    grdInventory.Col = 0
    lMaterialID = CLng(grdInventory.Text)
    
    ' Load material
    Dim objInv As New clsInventory
    If Not objInv.Load(lMaterialID) Then
        MsgBox "Could not load material: " & objInv.LastError, vbExclamation, APP_TITLE
        Set objInv = Nothing
        Exit Sub
    End If
    
    ' Edit via InputBox chain (VB6 anti-pattern)
    Dim sValue As String
    
    sValue = InputBox("Material Name:", APP_TITLE, objInv.MaterialName)
    If Len(sValue) = 0 Then
        Set objInv = Nothing
        Exit Sub
    End If
    objInv.MaterialName = sValue
    
    sValue = InputBox("Quantity On Hand:", APP_TITLE, CStr(objInv.QuantityOnHand))
    If Len(sValue) > 0 And IsNumeric(sValue) Then
        objInv.QuantityOnHand = CLng(sValue)
    End If
    
    sValue = InputBox("Unit Cost:", APP_TITLE, CStr(objInv.UnitCost))
    If Len(sValue) > 0 And IsNumeric(sValue) Then
        objInv.UnitCost = CDbl(sValue)
    End If
    
    sValue = InputBox("Location:", APP_TITLE, objInv.Location)
    If Len(sValue) > 0 Then
        objInv.Location = sValue
    End If
    
    If objInv.Save() Then
        MsgBox "Material updated successfully.", vbInformation, APP_TITLE
        LoadInventoryData
    Else
        MsgBox "Failed to update material: " & objInv.LastError, vbExclamation, APP_TITLE
    End If
    
    Set objInv = Nothing
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdEdit_Click", Err.Description, Err.Number
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ErrHandler
    
    If grdInventory.Rows <= 1 Then
        MsgBox "No records to delete.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    ' Get selected material
    grdInventory.Col = 0
    Dim lMaterialID As Long
    lMaterialID = CLng(grdInventory.Text)
    
    grdInventory.Col = 1
    Dim sMaterialName As String
    sMaterialName = grdInventory.Text
    
    If Not ConfirmAction("Are you sure you want to delete '" & sMaterialName & "'?") Then
        Exit Sub
    End If
    
    Dim lResult As Long
    lResult = ExecuteSQL("DELETE FROM RawMaterials WHERE MaterialID = " & lMaterialID)
    
    If lResult > 0 Then
        MsgBox "Material deleted successfully.", vbInformation, APP_TITLE
        LogMessage "Material #" & lMaterialID & " (" & sMaterialName & ") deleted"
        LoadInventoryData
    Else
        MsgBox "Failed to delete material.", vbExclamation, APP_TITLE
    End If
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdDelete_Click", Err.Description, Err.Number
End Sub

Private Sub cmdRefresh_Click()
    LoadInventoryData
End Sub

Private Sub cmdExport_Click()
    On Error GoTo ErrHandler
    
    Dim sFile As String
    sFile = BrowseForFile("CSV Files (*.csv)|*.csv", "Export Inventory", "Inventory_" & Format$(Date, "yyyymmdd") & ".csv")
    
    If Len(sFile) = 0 Then Exit Sub
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT MaterialID, MaterialName, Supplier, QuantityOnHand, " & _
           "ReorderPoint, UnitCost, Location FROM RawMaterials ORDER BY MaterialName"
    
    Set rs = GetRecordset(sSQL)
    
    If Not rs Is Nothing Then
        ExportToCSV rs, sFile
        rs.Close
        Set rs = Nothing
    End If
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdExport_Click", Err.Description, Err.Number
End Sub

Private Sub cboSupplier_Click()
    ' Reload data when supplier filter changes
    LoadInventoryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' Nothing to clean up
End Sub
