VERSION 5.00
Begin VB.Form frmShipping 
   Caption         =   "Shipping Manifest"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   7800
   Begin VB.Frame fraShipping 
      Caption         =   "Shipping Details"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtManifestID 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtWorkOrderID 
         Height          =   315
         Left            =   5280
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtTrackingNumber 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   5415
      End
      Begin VB.ComboBox cboCarrier 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtShipDate 
         Height          =   315
         Left            =   5280
         TabIndex        =   10
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtWeight 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtBoxes 
         Height          =   315
         Left            =   5280
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtShipToName 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   2280
         Width           =   5415
      End
      Begin VB.TextBox txtShipToAddress 
         Height          =   1335
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   2760
         Width           =   5415
      End
      Begin VB.TextBox txtNotes 
         Height          =   615
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   4320
         Width           =   5415
      End
      Begin VB.Label lblManifestID 
         Caption         =   "Manifest ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   400
         Width           =   1335
      End
      Begin VB.Label lblWorkOrderID 
         Caption         =   "Work Order ID:"
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   400
         Width           =   1335
      End
      Begin VB.Label lblTrackingNumber 
         Caption         =   "Tracking #:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   880
         Width           =   1335
      End
      Begin VB.Label lblCarrier 
         Caption         =   "Carrier:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1360
         Width           =   1335
      End
      Begin VB.Label lblShipDate 
         Caption         =   "Ship Date:"
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   1360
         Width           =   1335
      End
      Begin VB.Label lblWeight 
         Caption         =   "Weight (lbs):"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1840
         Width           =   1335
      End
      Begin VB.Label lblBoxes 
         Caption         =   "# of Boxes:"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   1840
         Width           =   1335
      End
      Begin VB.Label lblShipToName 
         Caption         =   "Ship To Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2320
         Width           =   1335
      End
      Begin VB.Label lblShipToAddress 
         Caption         =   "Ship To Addr:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2800
         Width           =   1335
      End
      Begin VB.Label lblNotes 
         Caption         =   "Notes:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   4360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   435
      Left            =   1080
      TabIndex        =   21
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrintLabel 
      Caption         =   "Print &Label"
      Height          =   435
      Left            =   2640
      TabIndex        =   22
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   4200
      TabIndex        =   23
      Top             =   5640
      Width           =   1335
   End
End
Attribute VB_Name = "frmShipping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' Form: frmShipping
' Description: Shipping manifest entry form
' Created: February 2001
' Last Modified: March 2024
' ============================================================================

Option Explicit

Private m_ManifestID As Long
Private m_IsNewRecord As Boolean

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    m_IsNewRecord = True
    m_ManifestID = 0
    
    ' Populate carrier combo
    cboCarrier.Clear
    cboCarrier.AddItem "UPS"
    cboCarrier.AddItem "FedEx"
    cboCarrier.AddItem "USPS"
    cboCarrier.AddItem "DHL"
    cboCarrier.AddItem "Freight"
    cboCarrier.AddItem "Customer Pickup"
    cboCarrier.ListIndex = 0
    
    ' Default values
    txtManifestID.Text = "(New)"
    txtWorkOrderID.Text = ""
    txtTrackingNumber.Text = ""
    txtShipDate.Text = Format$(Date, g_DateFormat)
    txtWeight.Text = "0.0"
    txtBoxes.Text = "1"
    txtShipToName.Text = ""
    txtShipToAddress.Text = ""
    txtNotes.Text = ""
    
    cmdPrintLabel.Enabled = False
    
    ' Pre-fill work order if one is selected
    If g_CurrentWorkOrderID > 0 Then
        txtWorkOrderID.Text = CStr(g_CurrentWorkOrderID)
        LoadWorkOrderInfo CLng(g_CurrentWorkOrderID)
        g_CurrentWorkOrderID = 0
    End If
    
    txtWorkOrderID.SetFocus
    
    Exit Sub
    
ErrHandler:
    ShowError "frmShipping.Form_Load", Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: LoadWorkOrderInfo
' Description: Pre-fills shipping info from work order data
' ============================================================================
Private Sub LoadWorkOrderInfo(ByVal lWorkOrderID As Long)
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT CustomerName, CustomerPO FROM WorkOrders WHERE WorkOrderID = " & lWorkOrderID
    
    Set rs = GetRecordset(sSQL)
    
    If Not rs Is Nothing And Not rs.EOF Then
        txtShipToName.Text = SafeString(rs!CustomerName)
    End If
    
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    Exit Sub
    
ErrHandler:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub txtWorkOrderID_LostFocus()
    On Error Resume Next
    
    ' Auto-load work order info when user tabs out
    If Len(Trim$(txtWorkOrderID.Text)) > 0 And IsNumeric(txtWorkOrderID.Text) Then
        LoadWorkOrderInfo CLng(txtWorkOrderID.Text)
    End If
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    
    ' Validate required fields
    If Len(Trim$(txtWorkOrderID.Text)) = 0 Then
        MsgBox "Work Order ID is required.", vbExclamation, APP_TITLE
        txtWorkOrderID.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtWorkOrderID.Text) Then
        MsgBox "Work Order ID must be a number.", vbExclamation, APP_TITLE
        txtWorkOrderID.SetFocus
        Exit Sub
    End If
    
    Dim lWorkOrderID As Long
    lWorkOrderID = CLng(txtWorkOrderID.Text)
    
    ' Verify work order exists
    If Not RecordExists("WorkOrders", "WorkOrderID = " & lWorkOrderID) Then
        MsgBox "Work Order #" & lWorkOrderID & " was not found.", vbExclamation, APP_TITLE
        txtWorkOrderID.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtTrackingNumber.Text)) = 0 Then
        MsgBox "Tracking Number is required.", vbExclamation, APP_TITLE
        txtTrackingNumber.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txtShipDate.Text) Then
        MsgBox "Please enter a valid Ship Date.", vbExclamation, APP_TITLE
        txtShipDate.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtWeight.Text) Or CDbl(txtWeight.Text) <= 0 Then
        MsgBox "Please enter a valid weight.", vbExclamation, APP_TITLE
        txtWeight.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtBoxes.Text) Or CLng(txtBoxes.Text) <= 0 Then
        MsgBox "Please enter a valid number of boxes.", vbExclamation, APP_TITLE
        txtBoxes.SetFocus
        Exit Sub
    End If
    
    ' Generate new Manifest ID
    m_ManifestID = GetNextID("ShippingManifests", "ManifestID")
    
    ' Build INSERT SQL
    Dim sSQL As String
    sSQL = "INSERT INTO ShippingManifests (ManifestID, WorkOrderID, TrackingNumber, " & _
           "Carrier, ShipDate, Weight, NumberOfBoxes, ShipToName, ShipToAddress, " & _
           "Notes, CustomerPO, CreatedBy, CreatedDate) VALUES (" & _
           m_ManifestID & ", " & _
           lWorkOrderID & ", " & _
           "'" & SQLSafe(Trim$(txtTrackingNumber.Text)) & "', " & _
           "'" & SQLSafe(cboCarrier.Text) & "', " & _
           SQLDate(CDate(txtShipDate.Text)) & ", " & _
           CDbl(txtWeight.Text) & ", " & _
           CLng(txtBoxes.Text) & ", " & _
           "'" & SQLSafe(Trim$(txtShipToName.Text)) & "', " & _
           "'" & SQLSafe(txtShipToAddress.Text) & "', " & _
           "'" & SQLSafe(txtNotes.Text) & "', " & _
           "'" & SQLSafe(SafeString(GetScalarValue( _
               "SELECT CustomerPO FROM WorkOrders WHERE WorkOrderID = " & lWorkOrderID, ""))) & "', " & _
           g_CurrentUserID & ", " & _
           SQLDateTime(Now) & ")"
    
    SetWaitCursor True
    
    Dim lResult As Long
    lResult = ExecuteSQL(sSQL)
    
    If lResult > 0 Then
        ' Update work order status to Completed
        ExecuteSQL "UPDATE WorkOrders SET Status = '" & WO_STATUS_COMPLETED & "', " & _
                  "CompletionDate = " & SQLDateTime(Now) & " " & _
                  "WHERE WorkOrderID = " & lWorkOrderID
        
        LogMessage "Shipping Manifest #" & m_ManifestID & " created for WO#" & lWorkOrderID
        
        SetWaitCursor False
        
        MsgBox "Shipping Manifest saved successfully." & vbCrLf & _
               "Manifest ID: " & m_ManifestID & vbCrLf & _
               "Carrier: " & cboCarrier.Text & vbCrLf & _
               "Tracking #: " & txtTrackingNumber.Text & vbCrLf & vbCrLf & _
               "Work Order has been marked as Completed.", vbInformation, APP_TITLE
        
        m_IsNewRecord = False
        txtManifestID.Text = CStr(m_ManifestID)
        cmdPrintLabel.Enabled = True
        
        ' Ask if user wants to print label
        If ConfirmAction("Do you want to print a shipping label now?") Then
            cmdPrintLabel_Click
        End If
    Else
        SetWaitCursor False
        MsgBox "Failed to save Shipping Manifest.", vbCritical, APP_TITLE
    End If
    
    Exit Sub
    
ErrHandler:
    SetWaitCursor False
    ShowError "cmdSave_Click", Err.Description, Err.Number
End Sub

Private Sub cmdPrintLabel_Click()
    On Error GoTo ErrHandler
    
    If m_ManifestID = 0 Then
        MsgBox "Please save the manifest before printing.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    PrintShippingLabel m_ManifestID, g_PrintPreviewEnabled
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdPrintLabel_Click", Err.Description, Err.Number
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' Nothing to clean up
End Sub
