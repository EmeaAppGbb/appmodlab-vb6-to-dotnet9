VERSION 5.00
Begin VB.Form frmWorkOrder 
   Caption         =   "Work Order Entry"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   8400
   Begin VB.Frame fraDetails 
      Caption         =   "Work Order Details"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.TextBox txtWorkOrderID 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtWorkOrderNumber 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtPartNumber 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdPartLookup 
         Caption         =   "..."
         Height          =   315
         Left            =   3600
         TabIndex        =   7
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtPartDescription 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtQuantity 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtCustomerPO 
         Height          =   315
         Left            =   5280
         TabIndex        =   13
         Top             =   1320
         Width           =   2295
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ComboBox cboPriority 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtDueDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtStartDate 
         Height          =   315
         Left            =   5280
         TabIndex        =   21
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtCustomerName 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   2760
         Width           =   5895
      End
      Begin VB.TextBox txtUnitCost 
         Height          =   315
         Left            =   1680
         TabIndex        =   25
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtTotalCost 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3240
         Width           =   2295
      End
      Begin VB.TextBox txtNotes 
         Height          =   1335
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   3720
         Width           =   5895
      End
      Begin VB.Label lblWorkOrderID 
         Caption         =   "Work Order ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   400
         Width           =   1335
      End
      Begin VB.Label lblWorkOrderNumber 
         Caption         =   "WO Number:"
         Height          =   255
         Left            =   4080
         TabIndex        =   3
         Top             =   400
         Width           =   1095
      End
      Begin VB.Label lblPartNumber 
         Caption         =   "Part Number:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   880
         Width           =   1335
      End
      Begin VB.Label lblPartDescription 
         Caption         =   "Description:"
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   880
         Width           =   1095
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Quantity:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1360
         Width           =   1335
      End
      Begin VB.Label lblCustomerPO 
         Caption         =   "Customer PO:"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   1360
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1840
         Width           =   1335
      End
      Begin VB.Label lblPriority 
         Caption         =   "Priority:"
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   1840
         Width           =   1095
      End
      Begin VB.Label lblDueDate 
         Caption         =   "Due Date:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2320
         Width           =   1335
      End
      Begin VB.Label lblStartDate 
         Caption         =   "Start Date:"
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   2320
         Width           =   1095
      End
      Begin VB.Label lblCustomerName 
         Caption         =   "Customer:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2800
         Width           =   1335
      End
      Begin VB.Label lblUnitCost 
         Caption         =   "Unit Cost:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   3280
         Width           =   1335
      End
      Begin VB.Label lblTotalCost 
         Caption         =   "Total Cost:"
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   3280
         Width           =   1095
      End
      Begin VB.Label lblNotes 
         Caption         =   "Notes:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   3760
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   435
      Left            =   960
      TabIndex        =   30
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   435
      Left            =   2520
      TabIndex        =   31
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   435
      Left            =   4080
      TabIndex        =   32
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   5640
      TabIndex        =   33
      Top             =   6120
      Width           =   1335
   End
End
Attribute VB_Name = "frmWorkOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' Form: frmWorkOrder
' Description: Work order entry and editing form
' Created: January 2001
' Last Modified: March 2024
' ============================================================================

Option Explicit

Private m_WorkOrder As clsWorkOrder
Private m_IsNewRecord As Boolean

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    Set m_WorkOrder = New clsWorkOrder
    
    ' Populate combo boxes
    cboStatus.Clear
    cboStatus.AddItem WO_STATUS_NEW
    cboStatus.AddItem WO_STATUS_INPROGRESS
    cboStatus.AddItem WO_STATUS_COMPLETED
    cboStatus.AddItem WO_STATUS_CANCELLED
    cboStatus.AddItem WO_STATUS_HOLD
    
    cboPriority.Clear
    cboPriority.AddItem "Low"
    cboPriority.AddItem "Normal"
    cboPriority.AddItem "High"
    cboPriority.AddItem "Urgent"
    
    ' Check if we're editing an existing work order
    If g_CurrentWorkOrderID > 0 Then
        ' Load existing work order
        If m_WorkOrder.Load(g_CurrentWorkOrderID) Then
            PopulateForm
            m_IsNewRecord = False
            Me.Caption = "Work Order - " & m_WorkOrder.WorkOrderNumber
        Else
            MsgBox "Could not load Work Order #" & g_CurrentWorkOrderID & vbCrLf & _
                   m_WorkOrder.LastError, vbExclamation, APP_TITLE
            m_IsNewRecord = True
            InitNewRecord
        End If
        g_CurrentWorkOrderID = 0
    Else
        ' New work order
        m_IsNewRecord = True
        InitNewRecord
    End If
    
    Exit Sub
    
ErrHandler:
    ShowError "frmWorkOrder.Form_Load", Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: InitNewRecord
' Description: Sets up form for a new work order
' ============================================================================
Private Sub InitNewRecord()
    On Error Resume Next
    
    txtWorkOrderID.Text = "(New)"
    txtWorkOrderNumber.Text = GenerateWorkOrderNumber()
    txtPartNumber.Text = ""
    txtPartDescription.Text = ""
    txtQuantity.Text = ""
    txtCustomerPO.Text = ""
    cboStatus.ListIndex = 0  ' New
    cboPriority.ListIndex = 1  ' Normal
    txtDueDate.Text = Format$(DateAdd("d", 14, Date), g_DateFormat)
    txtStartDate.Text = Format$(Date, g_DateFormat)
    txtCustomerName.Text = ""
    txtUnitCost.Text = "0.00"
    txtTotalCost.Text = "$0.00"
    txtNotes.Text = ""
    
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    
    Me.Caption = "Work Order - New"
    
    txtPartNumber.SetFocus
End Sub

' ============================================================================
' Sub: PopulateForm
' Description: Fills form controls from the work order object
' ============================================================================
Private Sub PopulateForm()
    On Error Resume Next
    
    txtWorkOrderID.Text = CStr(m_WorkOrder.WorkOrderID)
    txtWorkOrderNumber.Text = m_WorkOrder.WorkOrderNumber
    txtPartNumber.Text = m_WorkOrder.PartNumber
    txtPartDescription.Text = m_WorkOrder.PartDescription
    txtQuantity.Text = CStr(m_WorkOrder.Quantity)
    txtCustomerPO.Text = m_WorkOrder.CustomerPO
    txtCustomerName.Text = m_WorkOrder.CustomerName
    txtUnitCost.Text = Format$(m_WorkOrder.UnitCost, "0.00")
    txtTotalCost.Text = FormatCurrencyValue(m_WorkOrder.TotalCost)
    txtDueDate.Text = Format$(m_WorkOrder.DueDate, g_DateFormat)
    txtStartDate.Text = Format$(m_WorkOrder.StartDate, g_DateFormat)
    txtNotes.Text = m_WorkOrder.Notes
    
    ' Set combo boxes
    SelectComboItem cboStatus, m_WorkOrder.Status
    SelectComboItem cboPriority, m_WorkOrder.Priority
    
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
    
    g_DataModified = False
End Sub

' ============================================================================
' Sub: PopulateObject
' Description: Fills the work order object from form controls
' ============================================================================
Private Sub PopulateObject()
    On Error Resume Next
    
    If Not m_IsNewRecord Then
        m_WorkOrder.WorkOrderID = CLng(txtWorkOrderID.Text)
    End If
    
    m_WorkOrder.WorkOrderNumber = txtWorkOrderNumber.Text
    m_WorkOrder.PartNumber = Trim$(txtPartNumber.Text)
    m_WorkOrder.Quantity = SafeLong(txtQuantity.Text)
    m_WorkOrder.CustomerPO = Trim$(txtCustomerPO.Text)
    m_WorkOrder.CustomerName = Trim$(txtCustomerName.Text)
    m_WorkOrder.UnitCost = SafeNumber(txtUnitCost.Text)
    m_WorkOrder.Notes = txtNotes.Text
    
    ' Combo boxes
    If cboStatus.ListIndex >= 0 Then
        m_WorkOrder.Status = cboStatus.Text
    End If
    If cboPriority.ListIndex >= 0 Then
        m_WorkOrder.Priority = cboPriority.Text
    End If
    
    ' Dates
    If IsDate(txtDueDate.Text) Then
        m_WorkOrder.DueDate = CDate(txtDueDate.Text)
    End If
    If IsDate(txtStartDate.Text) Then
        m_WorkOrder.StartDate = CDate(txtStartDate.Text)
    End If
End Sub

' ============================================================================
' Helper: SelectComboItem
' ============================================================================
Private Sub SelectComboItem(ByRef cbo As ComboBox, ByVal sValue As String)
    On Error Resume Next
    
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = sValue Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i
    
    cbo.ListIndex = 0
End Sub

' ============================================================================
' Button Events
' ============================================================================
Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    
    ' Validate required fields
    If Len(Trim$(txtPartNumber.Text)) = 0 Then
        MsgBox "Part Number is required.", vbExclamation, APP_TITLE
        txtPartNumber.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtQuantity.Text)) = 0 Or Not IsNumeric(txtQuantity.Text) Then
        MsgBox "Please enter a valid quantity.", vbExclamation, APP_TITLE
        txtQuantity.SetFocus
        Exit Sub
    End If
    
    If CLng(txtQuantity.Text) <= 0 Then
        MsgBox "Quantity must be greater than zero.", vbExclamation, APP_TITLE
        txtQuantity.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txtDueDate.Text) Then
        MsgBox "Please enter a valid Due Date.", vbExclamation, APP_TITLE
        txtDueDate.SetFocus
        Exit Sub
    End If
    
    ' Populate object from form
    PopulateObject
    
    ' Save
    SetWaitCursor True
    
    If m_WorkOrder.Save() Then
        SetWaitCursor False
        MsgBox "Work Order saved successfully." & vbCrLf & _
               "WO#: " & m_WorkOrder.WorkOrderNumber, vbInformation, APP_TITLE
        
        m_IsNewRecord = False
        txtWorkOrderID.Text = CStr(m_WorkOrder.WorkOrderID)
        txtTotalCost.Text = FormatCurrencyValue(m_WorkOrder.TotalCost)
        
        cmdDelete.Enabled = True
        cmdPrint.Enabled = True
        
        Me.Caption = "Work Order - " & m_WorkOrder.WorkOrderNumber
        
        g_DataModified = False
    Else
        SetWaitCursor False
        MsgBox "Failed to save Work Order:" & vbCrLf & m_WorkOrder.LastError, _
               vbExclamation, APP_TITLE
    End If
    
    Exit Sub
    
ErrHandler:
    SetWaitCursor False
    ShowError "cmdSave_Click", Err.Description, Err.Number
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ErrHandler
    
    If m_IsNewRecord Then
        MsgBox "This work order has not been saved yet.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    If Not ConfirmAction("Are you sure you want to delete Work Order " & _
                         m_WorkOrder.WorkOrderNumber & "?") Then
        Exit Sub
    End If
    
    If m_WorkOrder.Delete() Then
        MsgBox "Work Order deleted successfully.", vbInformation, APP_TITLE
        Unload Me
    Else
        MsgBox "Failed to delete Work Order:" & vbCrLf & m_WorkOrder.LastError, _
               vbExclamation, APP_TITLE
    End If
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdDelete_Click", Err.Description, Err.Number
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrHandler
    
    If m_IsNewRecord Then
        MsgBox "Please save the work order before printing.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    PrintWorkOrder m_WorkOrder.WorkOrderID, g_PrintPreviewEnabled
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdPrint_Click", Err.Description, Err.Number
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    
    If g_DataModified Then
        If Not ConfirmAction("Discard unsaved changes?") Then
            Exit Sub
        End If
    End If
    
    Unload Me
End Sub

Private Sub cmdPartLookup_Click()
    On Error GoTo ErrHandler
    
    ' Show part lookup dialog
    g_CurrentPartNumber = ""
    
    Dim frm As New frmPartLookup
    frm.Show vbModal
    
    If Len(g_CurrentPartNumber) > 0 Then
        txtPartNumber.Text = g_CurrentPartNumber
        
        ' Look up part description
        Dim objPart As New clsPart
        If objPart.Load(g_CurrentPartNumber) Then
            txtPartDescription.Text = objPart.Description
            txtUnitCost.Text = Format$(objPart.UnitCost, "0.00")
        End If
        Set objPart = Nothing
    End If
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdPartLookup_Click", Err.Description, Err.Number
End Sub

' ============================================================================
' Change tracking
' ============================================================================
Private Sub txtPartNumber_Change()
    g_DataModified = True
End Sub

Private Sub txtQuantity_Change()
    g_DataModified = True
    
    ' Recalculate total cost
    On Error Resume Next
    If IsNumeric(txtQuantity.Text) And IsNumeric(txtUnitCost.Text) Then
        Dim dTotal As Double
        dTotal = CDbl(txtQuantity.Text) * CDbl(txtUnitCost.Text) * 1.15  ' 15% overhead
        dTotal = dTotal + (CDbl(txtQuantity.Text) * 0.5 * 45)  ' Labor
        txtTotalCost.Text = FormatCurrencyValue(dTotal)
    End If
End Sub

Private Sub txtCustomerPO_Change()
    g_DataModified = True
End Sub

Private Sub txtNotes_Change()
    g_DataModified = True
End Sub

Private Sub cboStatus_Click()
    g_DataModified = True
End Sub

Private Sub cboPriority_Click()
    g_DataModified = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set m_WorkOrder = Nothing
End Sub
