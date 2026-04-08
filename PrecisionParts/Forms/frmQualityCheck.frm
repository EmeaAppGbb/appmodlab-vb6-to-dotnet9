VERSION 5.00
Begin VB.Form frmQualityCheck 
   Caption         =   "Quality Control Inspection"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   7800
   Begin VB.Frame fraInspection 
      Caption         =   "Inspection Details"
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtCheckID 
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
      Begin VB.TextBox txtInspector 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtCheckDate 
         Height          =   315
         Left            =   5280
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.Frame fraResult 
         Caption         =   "Inspection Result"
         Height          =   735
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   7095
         Begin VB.OptionButton optPass 
            Caption         =   "&Pass"
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Top             =   300
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optFail 
            Caption         =   "&Fail"
            Height          =   315
            Left            =   1560
            TabIndex        =   11
            Top             =   300
            Width           =   1095
         End
         Begin VB.OptionButton optConditional 
            Caption         =   "&Conditional"
            Height          =   315
            Left            =   2880
            TabIndex        =   12
            Top             =   300
            Width           =   1455
         End
         Begin VB.OptionButton optRecheck 
            Caption         =   "&Recheck"
            Height          =   315
            Left            =   4560
            TabIndex        =   13
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.TextBox txtMeasurements 
         Height          =   1335
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2280
         Width           =   5415
      End
      Begin VB.TextBox txtNotes 
         Height          =   1335
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   3840
         Width           =   5415
      End
      Begin VB.Label lblCheckID 
         Caption         =   "Check ID:"
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
      Begin VB.Label lblInspector 
         Caption         =   "Inspector:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   880
         Width           =   1335
      End
      Begin VB.Label lblCheckDate 
         Caption         =   "Check Date:"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   880
         Width           =   1335
      End
      Begin VB.Label lblMeasurements 
         Caption         =   "Measurements:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2320
         Width           =   1335
      End
      Begin VB.Label lblNotes 
         Caption         =   "Notes:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3880
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   435
      Left            =   1080
      TabIndex        =   18
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdViewWorkOrder 
      Caption         =   "&View Work Order"
      Height          =   435
      Left            =   2640
      TabIndex        =   19
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   4560
      TabIndex        =   20
      Top             =   5880
      Width           =   1335
   End
End
Attribute VB_Name = "frmQualityCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' Form: frmQualityCheck
' Description: Quality control inspection entry form
' Created: February 2001
' Last Modified: March 2024
' ============================================================================

Option Explicit

Private m_CheckID As Long
Private m_IsNewRecord As Boolean

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    m_IsNewRecord = True
    m_CheckID = 0
    
    ' Default values
    txtCheckID.Text = "(New)"
    txtWorkOrderID.Text = ""
    txtInspector.Text = g_CurrentUsername
    txtCheckDate.Text = Format$(Now, g_DateFormat)
    txtMeasurements.Text = ""
    txtNotes.Text = ""
    optPass.Value = True
    
    ' If a work order is currently selected, pre-fill it
    If g_CurrentWorkOrderID > 0 Then
        txtWorkOrderID.Text = CStr(g_CurrentWorkOrderID)
        g_CurrentWorkOrderID = 0
    End If
    
    txtWorkOrderID.SetFocus
    
    Exit Sub
    
ErrHandler:
    ShowError "frmQualityCheck.Form_Load", Err.Description, Err.Number
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
    
    ' Verify work order exists
    Dim lWorkOrderID As Long
    lWorkOrderID = CLng(txtWorkOrderID.Text)
    
    If Not RecordExists("WorkOrders", "WorkOrderID = " & lWorkOrderID) Then
        MsgBox "Work Order #" & lWorkOrderID & " was not found.", vbExclamation, APP_TITLE
        txtWorkOrderID.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtInspector.Text)) = 0 Then
        MsgBox "Inspector name is required.", vbExclamation, APP_TITLE
        txtInspector.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txtCheckDate.Text) Then
        MsgBox "Please enter a valid Check Date.", vbExclamation, APP_TITLE
        txtCheckDate.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtMeasurements.Text)) = 0 Then
        MsgBox "Measurements are required for quality inspection.", vbExclamation, APP_TITLE
        txtMeasurements.SetFocus
        Exit Sub
    End If
    
    ' Determine result
    Dim sResult As String
    If optPass.Value Then
        sResult = QC_PASS
    ElseIf optFail.Value Then
        sResult = QC_FAIL
    ElseIf optConditional.Value Then
        sResult = QC_CONDITIONAL
    ElseIf optRecheck.Value Then
        sResult = QC_RECHECK
    End If
    
    ' Generate new Check ID
    m_CheckID = GetNextID("QualityChecks", "CheckID")
    
    ' Build INSERT SQL
    Dim sSQL As String
    sSQL = "INSERT INTO QualityChecks (CheckID, WorkOrderID, Inspector, " & _
           "CheckDate, Result, Measurements, Notes, CreatedBy, CreatedDate) " & _
           "VALUES (" & m_CheckID & ", " & _
           lWorkOrderID & ", " & _
           "'" & SQLSafe(Trim$(txtInspector.Text)) & "', " & _
           SQLDate(CDate(txtCheckDate.Text)) & ", " & _
           "'" & SQLSafe(sResult) & "', " & _
           "'" & SQLSafe(txtMeasurements.Text) & "', " & _
           "'" & SQLSafe(txtNotes.Text) & "', " & _
           g_CurrentUserID & ", " & _
           SQLDateTime(Now) & ")"
    
    SetWaitCursor True
    
    Dim lResult As Long
    lResult = ExecuteSQL(sSQL)
    
    SetWaitCursor False
    
    If lResult > 0 Then
        ' If failed, update work order status
        If sResult = QC_FAIL Then
            ExecuteSQL "UPDATE WorkOrders SET Status = '" & WO_STATUS_HOLD & _
                      "' WHERE WorkOrderID = " & lWorkOrderID
            
            MsgBox "Quality Check FAILED - Work Order has been placed on hold.", _
                   vbExclamation, APP_TITLE
        End If
        
        LogMessage "Quality Check #" & m_CheckID & " saved for WO#" & lWorkOrderID & _
                   " - Result: " & sResult
        
        MsgBox "Quality Check saved successfully." & vbCrLf & _
               "Check ID: " & m_CheckID & vbCrLf & _
               "Result: " & sResult, vbInformation, APP_TITLE
        
        ' Reset form for next entry
        m_IsNewRecord = True
        txtCheckID.Text = "(New)"
        txtWorkOrderID.Text = ""
        txtMeasurements.Text = ""
        txtNotes.Text = ""
        optPass.Value = True
        txtCheckDate.Text = Format$(Now, g_DateFormat)
        
        txtWorkOrderID.SetFocus
    Else
        MsgBox "Failed to save Quality Check record.", vbCritical, APP_TITLE
    End If
    
    Exit Sub
    
ErrHandler:
    SetWaitCursor False
    ShowError "cmdSave_Click", Err.Description, Err.Number
End Sub

Private Sub cmdViewWorkOrder_Click()
    On Error GoTo ErrHandler
    
    If Len(Trim$(txtWorkOrderID.Text)) = 0 Or Not IsNumeric(txtWorkOrderID.Text) Then
        MsgBox "Please enter a valid Work Order ID.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    g_CurrentWorkOrderID = CLng(txtWorkOrderID.Text)
    
    Dim frm As New frmWorkOrder
    frm.Show
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdViewWorkOrder_Click", Err.Description, Err.Number
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' Nothing to clean up
End Sub
