VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Precision Parts Manufacturing - Login"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "(None)"
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtUsername 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Precision Parts Manufacturing System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1720
      Width           =   1095
   End
   Begin VB.Label lblUsername 
      Caption         =   "&Username:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1120
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================================
' Form: frmLogin
' Description: User login form - authenticates against Users table
' Created: January 2001
' Last Modified: March 2024
' ============================================================================

Option Explicit

Private m_LoginAttempts As Integer
Private Const MAX_LOGIN_ATTEMPTS = 3

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    m_LoginAttempts = 0
    
    ' Initialize globals
    InitializeGlobals
    
    ' Open database connection
    If Not OpenDatabaseConnection() Then
        MsgBox "Cannot connect to database. The application will now close.", _
               vbCritical, APP_TITLE
        Unload Me
        Exit Sub
    End If
    
    ' Set form caption with version
    Me.Caption = APP_TITLE & " - Login (v" & APP_VERSION & ")"
    
    ' Focus on username field
    txtUsername.SetFocus
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error initializing login form: " & Err.Description, vbCritical, APP_TITLE
End Sub

Private Sub cmdLogin_Click()
    On Error GoTo ErrHandler
    
    Dim sUsername As String
    Dim sPassword As String
    
    sUsername = Trim$(txtUsername.Text)
    sPassword = txtPassword.Text
    
    ' Basic validation
    If Len(sUsername) = 0 Then
        MsgBox "Please enter your username.", vbExclamation, APP_TITLE
        txtUsername.SetFocus
        Exit Sub
    End If
    
    If Len(sPassword) = 0 Then
        MsgBox "Please enter your password.", vbExclamation, APP_TITLE
        txtPassword.SetFocus
        Exit Sub
    End If
    
    ' Authenticate against Users table
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT UserID, Username, UserRole, IsActive " & _
           "FROM Users " & _
           "WHERE Username = '" & SQLSafe(sUsername) & "' " & _
           "AND Password = '" & SQLSafe(sPassword) & "'"
    
    Set rs = GetRecordset(sSQL)
    
    If rs Is Nothing Then
        MsgBox "Database error during authentication.", vbCritical, APP_TITLE
        Exit Sub
    End If
    
    If rs.EOF Then
        ' Invalid credentials
        m_LoginAttempts = m_LoginAttempts + 1
        
        LogMessage "Failed login attempt for user: " & sUsername & " (Attempt " & m_LoginAttempts & ")"
        
        If m_LoginAttempts >= MAX_LOGIN_ATTEMPTS Then
            MsgBox "Maximum login attempts exceeded. The application will now close.", _
                   vbCritical, APP_TITLE
            rs.Close
            Set rs = Nothing
            CloseDatabaseConnection
            Unload Me
            End
        End If
        
        MsgBox "Invalid username or password." & vbCrLf & _
               "Attempts remaining: " & (MAX_LOGIN_ATTEMPTS - m_LoginAttempts), _
               vbExclamation, APP_TITLE
        
        txtPassword.Text = ""
        txtPassword.SetFocus
        
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    ' Check if account is active
    If Not SafeBool(rs!IsActive) Then
        MsgBox "Your account has been deactivated. Please contact your administrator.", _
               vbExclamation, APP_TITLE
        rs.Close
        Set rs = Nothing
        txtUsername.Text = ""
        txtPassword.Text = ""
        txtUsername.SetFocus
        Exit Sub
    End If
    
    ' Login successful - set globals
    g_CurrentUserID = SafeLong(rs!UserID)
    g_CurrentUsername = SafeString(rs!Username)
    g_CurrentUserRole = SafeString(rs!UserRole)
    g_IsAdministrator = (g_CurrentUserRole = ROLE_ADMIN)
    g_LoginTime = Now
    
    rs.Close
    Set rs = Nothing
    
    ' Update last login timestamp
    ExecuteSQL "UPDATE Users SET LastLogin = " & SQLDateTime(Now) & _
              " WHERE UserID = " & g_CurrentUserID
    
    LogMessage "User '" & g_CurrentUsername & "' logged in successfully (Role: " & g_CurrentUserRole & ")"
    
    ' Show main form
    Me.Hide
    frmMain.Show
    Unload Me
    
    Exit Sub
    
ErrHandler:
    ShowError "frmLogin.cmdLogin_Click", Err.Description, Err.Number
    
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    
    CloseDatabaseConnection
    CleanupGlobals
    
    Unload Me
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' Nothing to clean up
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    ' Move to password field on Enter
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPassword.SetFocus
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    ' Trigger login on Enter
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdLogin_Click
    End If
End Sub
