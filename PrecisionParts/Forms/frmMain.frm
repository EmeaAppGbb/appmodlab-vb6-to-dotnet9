VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Precision Parts Manufacturing System"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11520
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrRefresh 
      Enabled         =   -1  'True
      Interval        =   30000
      Left            =   120
      Top             =   120
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      ClientHeight    =   375
      ClientLeft      =   0
      ClientTop       =   6825
      ClientWidth     =   11520
      Left            =   0
      Top             =   6825
      Width           =   11520
      Height          =   375
      _ExtentX        =   20320
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Key             =   "pnlUser"
            Text            =   "User: "
            MinWidth        =   2000
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Key             =   "pnlRole"
            Text            =   "Role: "
            MinWidth        =   1500
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   1
            Key             =   "pnlTime"
            MinWidth        =   2000
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Key             =   "pnlDatabase"
            Text            =   "Database: Connected"
            MinWidth        =   2000
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileBackup 
         Caption         =   "&Backup Database"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrinterSetup 
         Caption         =   "Printer &Setup..."
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTransactions 
      Caption         =   "&Transactions"
      Begin VB.Menu mnuTransWorkOrders 
         Caption         =   "&Work Orders"
      End
      Begin VB.Menu mnuTransInventory 
         Caption         =   "&Inventory"
      End
      Begin VB.Menu mnuTransSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransQualityCheck 
         Caption         =   "&Quality Control"
      End
      Begin VB.Menu mnuTransShipping 
         Caption         =   "&Shipping"
      End
      Begin VB.Menu mnuTransSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransPartLookup 
         Caption         =   "&Part Lookup"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportsPrint 
         Caption         =   "&Print Reports..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' Form: frmMain
' Description: MDI Parent form - main application window
' Created: January 2001
' Last Modified: March 2024
' ============================================================================

Option Explicit

Private Sub MDIForm_Load()
    On Error GoTo ErrHandler
    
    g_MDIFormLoaded = True
    
    ' Set window title
    Me.Caption = APP_TITLE & " - [" & g_CurrentUsername & "]"
    
    ' Update status bar
    UpdateStatusBar
    
    ' Set timer interval from config
    tmrRefresh.Interval = g_AutoRefreshInterval * 1000
    tmrRefresh.Enabled = g_TimerEnabled
    
    ' Log application start
    LogMessage "Main form loaded - Application started"
    
    ' Disable menu items based on user role
    ApplyRolePermissions
    
    Exit Sub
    
ErrHandler:
    ShowError "frmMain.MDIForm_Load", Err.Description, Err.Number
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    
    ' Confirm exit
    If g_DataModified Then
        Dim iResult As Integer
        iResult = MsgBox("There are unsaved changes. Are you sure you want to exit?", _
                         vbQuestion + vbYesNo, APP_TITLE)
        If iResult = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    ' Cleanup
    tmrRefresh.Enabled = False
    
    CleanupCrystalReports
    CloseDatabaseConnection
    SaveConfigToRegistry
    CleanupGlobals
    
    LogMessage "Application closed by user: " & g_CurrentUsername
    
    g_MDIFormLoaded = False
    
    End
    
    Exit Sub
    
ErrHandler:
    LogMessage "MDIForm_Unload Error: " & Err.Description
    End
End Sub

' ============================================================================
' Menu Event Handlers
' ============================================================================
Private Sub mnuFileBackup_Click()
    On Error GoTo ErrHandler
    
    If ConfirmAction("This will create a backup of the database." & vbCrLf & _
                     "Do you want to proceed?") Then
        BackupDatabase
    End If
    
    Exit Sub
    
ErrHandler:
    ShowError "mnuFileBackup_Click", Err.Description, Err.Number
End Sub

Private Sub mnuFilePrinterSetup_Click()
    On Error GoTo ErrHandler
    
    ShowPrinterDialog
    
    Exit Sub
    
ErrHandler:
    ShowError "mnuFilePrinterSetup_Click", Err.Description, Err.Number
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuTransWorkOrders_Click()
    On Error GoTo ErrHandler
    
    Dim frm As New frmWorkOrder
    frm.Show
    
    Exit Sub
    
ErrHandler:
    ShowError "mnuTransWorkOrders_Click", Err.Description, Err.Number
End Sub

Private Sub mnuTransInventory_Click()
    On Error GoTo ErrHandler
    
    Dim frm As New frmInventory
    frm.Show
    
    Exit Sub
    
ErrHandler:
    ShowError "mnuTransInventory_Click", Err.Description, Err.Number
End Sub

Private Sub mnuTransQualityCheck_Click()
    On Error GoTo ErrHandler
    
    Dim frm As New frmQualityCheck
    frm.Show
    
    Exit Sub
    
ErrHandler:
    ShowError "mnuTransQualityCheck_Click", Err.Description, Err.Number
End Sub

Private Sub mnuTransShipping_Click()
    On Error GoTo ErrHandler
    
    Dim frm As New frmShipping
    frm.Show
    
    Exit Sub
    
ErrHandler:
    ShowError "mnuTransShipping_Click", Err.Description, Err.Number
End Sub

Private Sub mnuTransPartLookup_Click()
    On Error GoTo ErrHandler
    
    Dim frm As New frmPartLookup
    frm.Show vbModal
    
    If Len(g_CurrentPartNumber) > 0 Then
        ' Part was selected - could open work order form
        LogMessage "Part selected from lookup: " & g_CurrentPartNumber
    End If
    
    Exit Sub
    
ErrHandler:
    ShowError "mnuTransPartLookup_Click", Err.Description, Err.Number
End Sub

Private Sub mnuReportsPrint_Click()
    On Error GoTo ErrHandler
    
    Dim frm As New frmReports
    frm.Show vbModal
    
    Exit Sub
    
ErrHandler:
    ShowError "mnuReportsPrint_Click", Err.Description, Err.Number
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox APP_TITLE & vbCrLf & vbCrLf & _
           "Version: " & APP_VERSION & vbCrLf & _
           COPYRIGHT & vbCrLf & _
           COMPANY_NAME & vbCrLf & vbCrLf & _
           "Current User: " & g_CurrentUsername & vbCrLf & _
           "Role: " & g_CurrentUserRole & vbCrLf & _
           "Login Time: " & Format$(g_LoginTime, "mm/dd/yyyy hh:nn:ss AM/PM") & vbCrLf & _
           "Database: " & g_DatabasePath, _
           vbInformation, "About " & APP_TITLE
End Sub

' ============================================================================
' Timer Event - Auto refresh
' ============================================================================
Private Sub tmrRefresh_Timer()
    On Error Resume Next
    
    ' Update status bar
    UpdateStatusBar
    
    ' Check database connection
    If Not IsDatabaseConnected() Then
        sbMain.Panels("pnlDatabase").Text = "Database: DISCONNECTED"
        sbMain.Panels("pnlDatabase").ForeColor = vbRed
        
        ' Try to reconnect
        If EnsureConnection() Then
            sbMain.Panels("pnlDatabase").Text = "Database: Connected"
            sbMain.Panels("pnlDatabase").ForeColor = vbBlack
        End If
    End If
    
    DoEvents
End Sub

' ============================================================================
' Sub: UpdateStatusBar
' Description: Updates status bar panels with current info
' ============================================================================
Private Sub UpdateStatusBar()
    On Error Resume Next
    
    sbMain.Panels("pnlUser").Text = "User: " & g_CurrentUsername
    sbMain.Panels("pnlRole").Text = "Role: " & g_CurrentUserRole
    
    If IsDatabaseConnected() Then
        sbMain.Panels("pnlDatabase").Text = "Database: Connected"
    Else
        sbMain.Panels("pnlDatabase").Text = "Database: DISCONNECTED"
    End If
End Sub

' ============================================================================
' Sub: ApplyRolePermissions
' Description: Enable/disable menu items based on user role
' ============================================================================
Private Sub ApplyRolePermissions()
    On Error Resume Next
    
    Select Case g_CurrentUserRole
        Case ROLE_ADMIN
            ' Admin can access everything
            mnuFileBackup.Enabled = True
            mnuTransWorkOrders.Enabled = True
            mnuTransInventory.Enabled = True
            mnuTransQualityCheck.Enabled = True
            mnuTransShipping.Enabled = True
            mnuReportsPrint.Enabled = True
            
        Case ROLE_MANAGER
            mnuFileBackup.Enabled = False
            mnuTransWorkOrders.Enabled = True
            mnuTransInventory.Enabled = True
            mnuTransQualityCheck.Enabled = True
            mnuTransShipping.Enabled = True
            mnuReportsPrint.Enabled = True
            
        Case ROLE_OPERATOR
            mnuFileBackup.Enabled = False
            mnuTransWorkOrders.Enabled = True
            mnuTransInventory.Enabled = False
            mnuTransQualityCheck.Enabled = False
            mnuTransShipping.Enabled = False
            mnuReportsPrint.Enabled = False
            
        Case ROLE_QC
            mnuFileBackup.Enabled = False
            mnuTransWorkOrders.Enabled = True
            mnuTransInventory.Enabled = False
            mnuTransQualityCheck.Enabled = True
            mnuTransShipping.Enabled = False
            mnuReportsPrint.Enabled = True
            
        Case ROLE_SHIPPING
            mnuFileBackup.Enabled = False
            mnuTransWorkOrders.Enabled = True
            mnuTransInventory.Enabled = False
            mnuTransQualityCheck.Enabled = False
            mnuTransShipping.Enabled = True
            mnuReportsPrint.Enabled = True
            
        Case Else
            ' Default: minimal access
            mnuFileBackup.Enabled = False
            mnuTransWorkOrders.Enabled = True
            mnuTransInventory.Enabled = False
            mnuTransQualityCheck.Enabled = False
            mnuTransShipping.Enabled = False
            mnuReportsPrint.Enabled = False
    End Select
End Sub
