VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPartLookup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Number Lookup"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   100
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grdParts 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0   'None
      ScrollBars      =   2
      SelectionMode   =   1   'By Row
      AllowUserResizing=   1   'Columns
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   435
      Left            =   2160
      TabIndex        =   4
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   3840
      TabIndex        =   5
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblSearch 
      Caption         =   "Search:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   160
      Width           =   735
   End
   Begin VB.Label lblRecordCount 
      Alignment       =   2  'Center
      Caption         =   "Enter a part number or description to search"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   7575
   End
End
Attribute VB_Name = "frmPartLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' Form: frmPartLookup
' Description: Part number search dialog
' Created: February 2001
' Last Modified: March 2024
' ============================================================================

Option Explicit

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    g_CurrentPartNumber = ""
    
    ' Setup grid
    SetupGrid
    
    ' Load all active parts initially
    SearchParts ""
    
    txtSearch.SetFocus
    
    Exit Sub
    
ErrHandler:
    ShowError "frmPartLookup.Form_Load", Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: SetupGrid
' Description: Configures the parts grid columns
' ============================================================================
Private Sub SetupGrid()
    On Error Resume Next
    
    With grdParts
        .Cols = 6
        .Rows = 1
        .FixedRows = 1
        .FixedCols = 0
        .SelectionMode = flexSelectionByRow
        
        ' Column widths
        .ColWidth(0) = 600     ' ID
        .ColWidth(1) = 1500    ' Part Number
        .ColWidth(2) = 2400    ' Description
        .ColWidth(3) = 1200    ' Material
        .ColWidth(4) = 900     ' Unit Cost
        .ColWidth(5) = 800     ' Active
        
        ' Column headers
        .Row = 0
        .Col = 0: .Text = "ID"
        .Col = 1: .Text = "Part Number"
        .Col = 2: .Text = "Description"
        .Col = 3: .Text = "Material"
        .Col = 4: .Text = "Unit Cost"
        .Col = 5: .Text = "Active"
        
        ' Alignment
        .ColAlignment(4) = flexAlignRightCenter
    End With
End Sub

' ============================================================================
' Sub: SearchParts
' Description: Searches Parts table and populates grid
' Parameters: sSearchText - Text to search for
' ============================================================================
Private Sub SearchParts(ByVal sSearchText As String)
    On Error GoTo ErrHandler
    
    SetWaitCursor True
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim lRow As Long
    
    sSQL = "SELECT PartID, PartNumber, Description, Material, UnitCost, IsActive " & _
           "FROM Parts"
    
    If Len(Trim$(sSearchText)) > 0 Then
        sSQL = sSQL & " WHERE (PartNumber LIKE '%" & SQLSafe(sSearchText) & "%' " & _
               "OR Description LIKE '%" & SQLSafe(sSearchText) & "%' " & _
               "OR Material LIKE '%" & SQLSafe(sSearchText) & "%')"
    End If
    
    sSQL = sSQL & " ORDER BY PartNumber"
    
    Set rs = GetRecordset(sSQL)
    
    ' Clear grid
    grdParts.Rows = 1
    
    If rs Is Nothing Then
        lblRecordCount.Caption = "Error loading parts data"
        SetWaitCursor False
        Exit Sub
    End If
    
    If rs.EOF Then
        lblRecordCount.Caption = "No parts found"
        rs.Close
        Set rs = Nothing
        SetWaitCursor False
        Exit Sub
    End If
    
    ' Populate grid
    lRow = 0
    Do While Not rs.EOF
        lRow = lRow + 1
        grdParts.Rows = lRow + 1
        
        grdParts.Row = lRow
        grdParts.Col = 0: grdParts.Text = SafeString(rs!PartID)
        grdParts.Col = 1: grdParts.Text = SafeString(rs!PartNumber)
        grdParts.Col = 2: grdParts.Text = SafeString(rs!Description)
        grdParts.Col = 3: grdParts.Text = SafeString(rs!Material)
        grdParts.Col = 4: grdParts.Text = FormatCurrencyValue(SafeNumber(rs!UnitCost))
        grdParts.Col = 5: grdParts.Text = IIf(SafeBool(rs!IsActive), "Yes", "No")
        
        rs.MoveNext
        
        If lRow Mod 100 = 0 Then DoEvents
    Loop
    
    lblRecordCount.Caption = lRow & " part(s) found"
    
    rs.Close
    Set rs = Nothing
    
    ' Select first row
    If grdParts.Rows > 1 Then
        grdParts.Row = 1
    End If
    
    SetWaitCursor False
    
    Exit Sub
    
ErrHandler:
    SetWaitCursor False
    ShowError "SearchParts", Err.Description, Err.Number
    
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Sub

' ============================================================================
' Event Handlers
' ============================================================================
Private Sub cmdSearch_Click()
    SearchParts Trim$(txtSearch.Text)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    ' Search on Enter key
    If KeyAscii = 13 Then
        KeyAscii = 0
        SearchParts Trim$(txtSearch.Text)
    End If
End Sub

Private Sub txtSearch_Change()
    ' Auto-search after typing (with minimum 2 characters)
    If Len(txtSearch.Text) >= 2 Or Len(txtSearch.Text) = 0 Then
        SearchParts Trim$(txtSearch.Text)
    End If
End Sub

Private Sub grdParts_DblClick()
    ' Double-click to select
    cmdSelect_Click
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo ErrHandler
    
    If grdParts.Rows <= 1 Then
        MsgBox "No parts to select.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    ' Get selected part number
    grdParts.Col = 1
    g_CurrentPartNumber = grdParts.Text
    
    LogMessage "Part selected from lookup: " & g_CurrentPartNumber
    
    Unload Me
    
    Exit Sub
    
ErrHandler:
    ShowError "cmdSelect_Click", Err.Description, Err.Number
End Sub

Private Sub cmdCancel_Click()
    g_CurrentPartNumber = ""
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' Nothing to clean up
End Sub
