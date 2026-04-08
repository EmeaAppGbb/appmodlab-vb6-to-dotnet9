VERSION 5.00
Begin VB.Form frmReports 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Reports"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstReports 
      Height          =   2400
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.CheckBox chkPreview 
      Caption         =   "Print Pre&view"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3060
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Pre&view"
      Height          =   435
      Left            =   480
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   435
      Left            =   1920
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   435
      Left            =   3360
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblReports 
      Caption         =   "Select a report to print:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   3015
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' Form: frmReports
' Description: Report selection and printing form
' Created: February 2001
' Last Modified: March 2024
' ============================================================================

Option Explicit

' Report indices
Private Const RPT_IDX_WORKORDERS = 0
Private Const RPT_IDX_SHIPPING = 1
Private Const RPT_IDX_INVENTORY = 2
Private Const RPT_IDX_QUALITY = 3
Private Const RPT_IDX_PRODUCTION = 4

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    ' Populate report list
    lstReports.Clear
    lstReports.AddItem "Work Orders Report"               ' 0
    lstReports.AddItem "Shipping Labels"                   ' 1
    lstReports.AddItem "Inventory Status Report"           ' 2
    lstReports.AddItem "Quality Control Summary"           ' 3
    lstReports.AddItem "Production Summary Report"         ' 4
    
    lstReports.ListIndex = 0
    
    ' Set preview checkbox from global
    chkPreview.Value = IIf(g_PrintPreviewEnabled, 1, 0)
    
    CenterForm Me
    
    Exit Sub
    
ErrHandler:
    ShowError "frmReports.Form_Load", Err.Description, Err.Number
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo ErrHandler
    
    ' Force preview mode
    g_PrintPreviewEnabled = True
    chkPreview.Value = 1
    
    PrintSelectedReport True
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdPreview_Click", Err.Description, Err.Number
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrHandler
    
    ' Use checkbox setting
    g_PrintPreviewEnabled = (chkPreview.Value = 1)
    
    PrintSelectedReport (chkPreview.Value = 1)
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdPrint_Click", Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: PrintSelectedReport
' Description: Executes the selected report
' Parameters: bPreview - Whether to show print preview
' ============================================================================
Private Sub PrintSelectedReport(ByVal bPreview As Boolean)
    On Error GoTo ErrHandler
    
    If lstReports.ListIndex < 0 Then
        MsgBox "Please select a report.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    SetWaitCursor True
    
    Select Case lstReports.ListIndex
        Case RPT_IDX_WORKORDERS
            ' Print work orders - ask for specific WO or all
            Dim sWorkOrderID As String
            sWorkOrderID = InputBox("Enter Work Order ID to print (or leave blank for all):", _
                                     APP_TITLE, "")
            
            If Len(sWorkOrderID) > 0 And IsNumeric(sWorkOrderID) Then
                PrintWorkOrder CLng(sWorkOrderID), bPreview
            Else
                ' Print all open work orders
                Dim rs As ADODB.Recordset
                Dim sSQL As String
                sSQL = "SELECT WorkOrderID FROM WorkOrders " & _
                       "WHERE Status IN ('" & WO_STATUS_NEW & "', '" & WO_STATUS_INPROGRESS & "') " & _
                       "ORDER BY DueDate"
                
                Set rs = GetRecordset(sSQL)
                
                If Not rs Is Nothing Then
                    If rs.EOF Then
                        MsgBox "No open work orders to print.", vbInformation, APP_TITLE
                    Else
                        Dim lCount As Long
                        Do While Not rs.EOF
                            PrintWorkOrder SafeLong(rs!WorkOrderID), bPreview
                            lCount = lCount + 1
                            rs.MoveNext
                            DoEvents
                        Loop
                        
                        If Not bPreview Then
                            MsgBox lCount & " work order(s) sent to printer.", vbInformation, APP_TITLE
                        End If
                    End If
                    rs.Close
                    Set rs = Nothing
                End If
            End If
            
        Case RPT_IDX_SHIPPING
            ' Print shipping labels
            Dim sManifestID As String
            sManifestID = InputBox("Enter Manifest ID to print label:", APP_TITLE, "")
            
            If Len(sManifestID) > 0 And IsNumeric(sManifestID) Then
                PrintShippingLabel CLng(sManifestID), bPreview
            Else
                MsgBox "Please enter a valid Manifest ID.", vbExclamation, APP_TITLE
            End If
            
        Case RPT_IDX_INVENTORY
            ' Print inventory status
            PrintInventoryReport bPreview
            
        Case RPT_IDX_QUALITY
            ' Print quality check report
            Dim sCheckID As String
            sCheckID = InputBox("Enter Quality Check ID to print (or leave blank for summary):", _
                                APP_TITLE, "")
            
            If Len(sCheckID) > 0 And IsNumeric(sCheckID) Then
                PrintQualityCheckReport CLng(sCheckID), bPreview
            Else
                MsgBox "Please enter a valid Quality Check ID for detailed report.", _
                       vbExclamation, APP_TITLE
            End If
            
        Case RPT_IDX_PRODUCTION
            ' Production summary - not yet implemented
            MsgBox "Production Summary Report is not yet available." & vbCrLf & _
                   "This feature will be available in a future update.", _
                   vbInformation, APP_TITLE
            
        Case Else
            MsgBox "Unknown report selection.", vbExclamation, APP_TITLE
    End Select
    
    SetWaitCursor False
    
    Exit Sub
    
ErrHandler:
    SetWaitCursor False
    ShowError "PrintSelectedReport", Err.Description, Err.Number
    
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub chkPreview_Click()
    g_PrintPreviewEnabled = (chkPreview.Value = 1)
End Sub

Private Sub lstReports_DblClick()
    ' Double-click to preview
    cmdPreview_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' Nothing to clean up
End Sub
