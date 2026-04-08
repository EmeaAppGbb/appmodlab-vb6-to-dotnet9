Attribute VB_Name = "modGlobals"
' ============================================================================
' Module: modGlobals
' Description: Global variables and constants for Precision Parts System
' Created: January 2001
' Last Modified: March 2024
' ============================================================================

Option Explicit

' Application-wide global variables (VB6 anti-pattern)
Public g_CurrentUserID As Long
Public g_CurrentUsername As String
Public g_CurrentUserRole As String
Public g_IsAdministrator As Boolean
Public g_LoginTime As Date
Public g_DatabasePath As String
Public g_ReportsPath As String
Public g_CompanyName As String
Public g_AppVersion As String
Public g_DefaultPrinter As String

' UI state globals
Public g_MDIFormLoaded As Boolean
Public g_CurrentWorkOrderID As Long
Public g_CurrentPartNumber As String
Public g_GridSortColumn As Integer
Public g_GridSortAscending As Boolean
Public g_FilterActive As Boolean
Public g_FilterCriteria As String

' Configuration globals (loaded from registry)
Public g_DatabaseServer As String
Public g_BackupPath As String
Public g_AutoRefreshInterval As Integer
Public g_ShowSplashScreen As Boolean
Public g_EnableLogging As Boolean
Public g_LogFilePath As String
Public g_MaxRecordsToDisplay As Long
Public g_DateFormat As String
Public g_TimeFormat As String

' Printing globals
Public g_PrinterOrientation As Integer
Public g_PrinterCopies As Integer
Public g_PrinterCollate As Boolean
Public g_PrintPreviewEnabled As Boolean

' Status flags
Public g_DataModified As Boolean
Public g_RecordLocked As Boolean
Public g_SystemBusy As Boolean
Public g_TimerEnabled As Boolean

' Cache for frequently used data
Public g_PartsList As Collection
Public g_CustomersList As Collection
Public g_UsersList As Collection
Public g_MaterialsList As Collection

' Constants
Public Const APP_TITLE = "Precision Parts Manufacturing System"
Public Const APP_VERSION = "2.1.5"
Public Const COMPANY_NAME = "Precision Parts Manufacturing Inc."
Public Const COPYRIGHT = "Copyright © 2001-2024"

' Database constants
Public Const DB_TIMEOUT = 30
Public Const MAX_RETRY_COUNT = 3
Public Const RECORD_LOCK_TIMEOUT = 5000

' Work Order Status constants
Public Const WO_STATUS_NEW = "New"
Public Const WO_STATUS_INPROGRESS = "In Progress"
Public Const WO_STATUS_COMPLETED = "Completed"
Public Const WO_STATUS_CANCELLED = "Cancelled"
Public Const WO_STATUS_HOLD = "On Hold"

' Quality Check Results
Public Const QC_PASS = "Pass"
Public Const QC_FAIL = "Fail"
Public Const QC_CONDITIONAL = "Conditional"
Public Const QC_RECHECK = "Recheck Required"

' User Roles
Public Const ROLE_ADMIN = "Administrator"
Public Const ROLE_MANAGER = "Manager"
Public Const ROLE_OPERATOR = "Operator"
Public Const ROLE_QC = "Quality Control"
Public Const ROLE_SHIPPING = "Shipping"

' Grid column indices (MSFlexGrid)
Public Const COL_SELECT = 0
Public Const COL_ID = 1
Public Const COL_PART_NUMBER = 2
Public Const COL_DESCRIPTION = 3
Public Const COL_QUANTITY = 4
Public Const COL_STATUS = 5

' Error codes
Public Const ERR_DATABASE_CONNECTION = vbObjectError + 1001
Public Const ERR_RECORD_LOCKED = vbObjectError + 1002
Public Const ERR_INVALID_DATA = vbObjectError + 1003
Public Const ERR_PERMISSION_DENIED = vbObjectError + 1004
Public Const ERR_FILE_NOT_FOUND = vbObjectError + 1005

' Registry keys
Public Const REG_SECTION = "PrecisionParts"
Public Const REG_DATABASE = "DatabasePath"
Public Const REG_BACKUP = "BackupPath"
Public Const REG_REFRESH = "RefreshInterval"
Public Const REG_SHOW_SPLASH = "ShowSplash"
Public Const REG_ENABLE_LOG = "EnableLogging"
Public Const REG_LOG_PATH = "LogPath"

' ============================================================================
' Sub: InitializeGlobals
' Description: Initialize all global variables at startup
' ============================================================================
Public Sub InitializeGlobals()
    g_AppVersion = APP_VERSION
    g_CompanyName = COMPANY_NAME
    g_MDIFormLoaded = False
    g_DataModified = False
    g_RecordLocked = False
    g_SystemBusy = False
    g_TimerEnabled = True
    g_MaxRecordsToDisplay = 1000
    g_DateFormat = "mm/dd/yyyy"
    g_TimeFormat = "hh:nn:ss AM/PM"
    g_AutoRefreshInterval = 30
    g_PrinterCopies = 1
    g_PrinterCollate = False
    g_GridSortAscending = True
    
    Set g_PartsList = New Collection
    Set g_CustomersList = New Collection
    Set g_UsersList = New Collection
    Set g_MaterialsList = New Collection
    
    LoadConfigFromRegistry
End Sub

' ============================================================================
' Sub: LoadConfigFromRegistry
' Description: Load configuration settings from Windows Registry
' ============================================================================
Public Sub LoadConfigFromRegistry()
    On Error Resume Next
    
    g_DatabasePath = GetSetting(REG_SECTION, "Database", REG_DATABASE, "C:\PrecisionParts\Database\PrecisionParts.mdb")
    g_BackupPath = GetSetting(REG_SECTION, "Database", REG_BACKUP, "C:\PrecisionParts\Backup\")
    g_AutoRefreshInterval = Val(GetSetting(REG_SECTION, "UI", REG_REFRESH, "30"))
    g_ShowSplashScreen = CBool(GetSetting(REG_SECTION, "UI", REG_SHOW_SPLASH, "1"))
    g_EnableLogging = CBool(GetSetting(REG_SECTION, "System", REG_ENABLE_LOG, "1"))
    g_LogFilePath = GetSetting(REG_SECTION, "System", REG_LOG_PATH, "C:\PrecisionParts\Logs\")
    g_ReportsPath = GetSetting(REG_SECTION, "Reports", "Path", "C:\PrecisionParts\Reports\")
    
    On Error GoTo 0
End Sub

' ============================================================================
' Sub: SaveConfigToRegistry
' Description: Save configuration settings to Windows Registry
' ============================================================================
Public Sub SaveConfigToRegistry()
    On Error Resume Next
    
    SaveSetting REG_SECTION, "Database", REG_DATABASE, g_DatabasePath
    SaveSetting REG_SECTION, "Database", REG_BACKUP, g_BackupPath
    SaveSetting REG_SECTION, "UI", REG_REFRESH, CStr(g_AutoRefreshInterval)
    SaveSetting REG_SECTION, "UI", REG_SHOW_SPLASH, CStr(Abs(g_ShowSplashScreen))
    SaveSetting REG_SECTION, "System", REG_ENABLE_LOG, CStr(Abs(g_EnableLogging))
    SaveSetting REG_SECTION, "System", REG_LOG_PATH, g_LogFilePath
    SaveSetting REG_SECTION, "Reports", "Path", g_ReportsPath
    
    On Error GoTo 0
End Sub

' ============================================================================
' Sub: CleanupGlobals
' Description: Clean up global objects before application exit
' ============================================================================
Public Sub CleanupGlobals()
    On Error Resume Next
    
    Set g_PartsList = Nothing
    Set g_CustomersList = Nothing
    Set g_UsersList = Nothing
    Set g_MaterialsList = Nothing
    
    On Error GoTo 0
End Sub
