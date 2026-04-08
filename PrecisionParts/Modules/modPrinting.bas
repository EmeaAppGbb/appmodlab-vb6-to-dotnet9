Attribute VB_Name = "modPrinting"
' ============================================================================
' Module: modPrinting
' Description: Windows API printer functions and Crystal Reports integration
' Created: February 2001
' Last Modified: March 2024
' Notes: Depends on Crystal Reports 8.5 runtime
' ============================================================================

Option Explicit

' Windows API declarations for printer management
Private Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterA" _
    (ByVal pszBuffer As String, pcchBuffer As Long) As Long
Private Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" _
    (ByVal pPrinterName As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" _
    (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Printer orientation constants
Private Const DMORIENT_PORTRAIT = 1
Private Const DMORIENT_LANDSCAPE = 2

' Report file paths (hardcoded - VB6 anti-pattern)
Private Const RPT_WORK_ORDER = "C:\PrecisionParts\Reports\WorkOrder.rpt"
Private Const RPT_SHIPPING_LABEL = "C:\PrecisionParts\Reports\ShippingLabel.rpt"
Private Const RPT_INVENTORY = "C:\PrecisionParts\Reports\InventoryStatus.rpt"
Private Const RPT_QUALITY_CHECK = "C:\PrecisionParts\Reports\QualityCheck.rpt"
Private Const RPT_PRODUCTION_SUMMARY = "C:\PrecisionParts\Reports\ProductionSummary.rpt"

' Module-level Crystal Reports objects
Private m_CrystalApp As Object     ' CRAXDRT.Application
Private m_CrystalReport As Object  ' CRAXDRT.Report
Private m_ReportLoaded As Boolean

' ============================================================================
' Sub: InitializeCrystalReports
' Description: Creates Crystal Reports application object
' ============================================================================
Private Sub InitializeCrystalReports()
    On Error GoTo ErrHandler
    
    If m_CrystalApp Is Nothing Then
        Set m_CrystalApp = CreateObject("CrystalRuntime.Application")
        LogMessage "Crystal Reports runtime initialized"
    End If
    
    Exit Sub
    
ErrHandler:
    LogMessage "InitializeCrystalReports Error: " & Err.Description
    MsgBox "Crystal Reports runtime is not installed or could not be initialized." & vbCrLf & _
           "Please install Crystal Reports 8.5 runtime to enable printing." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, APP_TITLE
End Sub

' ============================================================================
' Sub: CleanupCrystalReports
' Description: Releases Crystal Reports objects
' ============================================================================
Public Sub CleanupCrystalReports()
    On Error Resume Next
    
    Set m_CrystalReport = Nothing
    Set m_CrystalApp = Nothing
    m_ReportLoaded = False
    
    LogMessage "Crystal Reports objects released"
End Sub

' ============================================================================
' Function: GetDefaultPrinterName
' Description: Returns the name of the default Windows printer
' ============================================================================
Public Function GetDefaultPrinterName() As String
    On Error GoTo ErrHandler
    
    Dim sBuffer As String
    Dim lSize As Long
    
    sBuffer = String$(256, Chr$(0))
    lSize = 256
    
    If GetDefaultPrinter(sBuffer, lSize) <> 0 Then
        GetDefaultPrinterName = Left$(sBuffer, lSize - 1)
    Else
        ' Fallback to win.ini method
        Dim sResult As String
        sResult = String$(256, Chr$(0))
        GetProfileString "windows", "device", "", sResult, 256
        
        If Len(sResult) > 0 Then
            GetDefaultPrinterName = Left$(sResult, InStr(sResult, ",") - 1)
        Else
            GetDefaultPrinterName = ""
        End If
    End If
    
    Exit Function
    
ErrHandler:
    GetDefaultPrinterName = ""
    LogMessage "GetDefaultPrinterName Error: " & Err.Description
End Function

' ============================================================================
' Sub: PrintWorkOrder
' Description: Prints a work order report using Crystal Reports
' Parameters: lWorkOrderID - ID of the work order to print
'             bPreview - Show print preview instead of printing directly
' ============================================================================
Public Sub PrintWorkOrder(ByVal lWorkOrderID As Long, Optional ByVal bPreview As Boolean = False)
    On Error GoTo ErrHandler
    
    SetWaitCursor True
    
    ' Validate work order exists
    If Not RecordExists("WorkOrders", "WorkOrderID = " & lWorkOrderID) Then
        MsgBox "Work Order #" & lWorkOrderID & " was not found in the database.", _
               vbExclamation, APP_TITLE
        SetWaitCursor False
        Exit Sub
    End If
    
    ' Check if report file exists
    If Dir(RPT_WORK_ORDER) = "" Then
        ' Fallback: use g_ReportsPath
        Dim sReportPath As String
        sReportPath = g_ReportsPath & "WorkOrder.rpt"
        
        If Dir(sReportPath) = "" Then
            MsgBox "Work Order report template not found." & vbCrLf & _
                   "Expected: " & RPT_WORK_ORDER & vbCrLf & _
                   "Or: " & sReportPath, vbCritical, APP_TITLE
            SetWaitCursor False
            Exit Sub
        End If
    Else
        sReportPath = RPT_WORK_ORDER
    End If
    
    ' Initialize Crystal Reports
    InitializeCrystalReports
    
    If m_CrystalApp Is Nothing Then
        SetWaitCursor False
        Exit Sub
    End If
    
    ' Open the report
    Set m_CrystalReport = m_CrystalApp.OpenReport(sReportPath)
    m_ReportLoaded = True
    
    ' Set database connection info
    m_CrystalReport.Database.Tables(1).SetLogOnInfo "", g_DatabasePath, "", ""
    
    ' Set selection formula for specific work order
    m_CrystalReport.RecordSelectionFormula = "{WorkOrders.WorkOrderID} = " & lWorkOrderID
    
    ' Set printer options
    If g_PrinterCopies > 1 Then
        m_CrystalReport.PrinterSetup g_PrinterCopies
    End If
    
    If bPreview Or g_PrintPreviewEnabled Then
        ' Show preview window
        ' Note: In real VB6, we'd use CRViewer control
        m_CrystalReport.Preview
        LogMessage "Work Order #" & lWorkOrderID & " previewed"
    Else
        ' Print directly
        m_CrystalReport.PrintOut False, g_PrinterCopies
        LogMessage "Work Order #" & lWorkOrderID & " printed (" & g_PrinterCopies & " copies)"
    End If
    
    DoEvents
    
    ' Cleanup
    Set m_CrystalReport = Nothing
    m_ReportLoaded = False
    
    SetWaitCursor False
    
    Exit Sub
    
ErrHandler:
    SetWaitCursor False
    
    On Error Resume Next
    Set m_CrystalReport = Nothing
    m_ReportLoaded = False
    
    ShowError "PrintWorkOrder", "Failed to print Work Order #" & lWorkOrderID & vbCrLf & Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: PrintShippingLabel
' Description: Prints a shipping label for a manifest
' Parameters: lManifestID - ID of the shipping manifest
'             bPreview - Show print preview
' ============================================================================
Public Sub PrintShippingLabel(ByVal lManifestID As Long, Optional ByVal bPreview As Boolean = False)
    On Error GoTo ErrHandler
    
    SetWaitCursor True
    
    ' Validate manifest exists
    If Not RecordExists("ShippingManifests", "ManifestID = " & lManifestID) Then
        MsgBox "Shipping Manifest #" & lManifestID & " was not found.", _
               vbExclamation, APP_TITLE
        SetWaitCursor False
        Exit Sub
    End If
    
    ' Get manifest details for label
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    sSQL = "SELECT sm.*, wo.PartNumber, wo.CustomerPO " & _
           "FROM ShippingManifests sm " & _
           "INNER JOIN WorkOrders wo ON sm.WorkOrderID = wo.WorkOrderID " & _
           "WHERE sm.ManifestID = " & lManifestID
    
    Set rs = GetRecordset(sSQL)
    
    If rs Is Nothing Or rs.EOF Then
        MsgBox "Could not retrieve shipping data for label.", vbExclamation, APP_TITLE
        SetWaitCursor False
        Exit Sub
    End If
    
    ' Check for report file
    Dim sReportPath As String
    sReportPath = g_ReportsPath & "ShippingLabel.rpt"
    
    If Dir(sReportPath) = "" Then
        sReportPath = RPT_SHIPPING_LABEL
    End If
    
    If Dir(sReportPath) = "" Then
        ' Fallback: Print label directly to printer using Printer object
        PrintShippingLabelDirect rs
        rs.Close
        Set rs = Nothing
        SetWaitCursor False
        Exit Sub
    End If
    
    ' Use Crystal Reports
    InitializeCrystalReports
    
    If m_CrystalApp Is Nothing Then
        ' Fallback to direct print
        PrintShippingLabelDirect rs
        rs.Close
        Set rs = Nothing
        SetWaitCursor False
        Exit Sub
    End If
    
    Set m_CrystalReport = m_CrystalApp.OpenReport(sReportPath)
    m_CrystalReport.Database.Tables(1).SetLogOnInfo "", g_DatabasePath, "", ""
    m_CrystalReport.RecordSelectionFormula = "{ShippingManifests.ManifestID} = " & lManifestID
    
    If bPreview Or g_PrintPreviewEnabled Then
        m_CrystalReport.Preview
    Else
        m_CrystalReport.PrintOut False, 1
    End If
    
    LogMessage "Shipping label printed for Manifest #" & lManifestID
    
    rs.Close
    Set rs = Nothing
    Set m_CrystalReport = Nothing
    
    SetWaitCursor False
    
    Exit Sub
    
ErrHandler:
    SetWaitCursor False
    
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    Set m_CrystalReport = Nothing
    
    ShowError "PrintShippingLabel", "Failed to print shipping label for Manifest #" & lManifestID & vbCrLf & Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: PrintShippingLabelDirect
' Description: Prints shipping label directly to printer without Crystal Reports
' Parameters: rs - Recordset with shipping data
' ============================================================================
Private Sub PrintShippingLabelDirect(ByVal rs As ADODB.Recordset)
    On Error GoTo ErrHandler
    
    ' Use VB6 Printer object for direct printing (fallback when Crystal not available)
    Printer.FontName = "Arial"
    Printer.FontSize = 14
    Printer.FontBold = True
    
    ' Company header
    Printer.Print COMPANY_NAME
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print String$(50, "-")
    Printer.Print ""
    
    ' Shipping details
    Printer.FontSize = 12
    Printer.Print "Ship To:"
    Printer.FontSize = 10
    Printer.Print "  Tracking #: " & SafeString(rs!TrackingNumber)
    Printer.Print "  Carrier: " & SafeString(rs!Carrier)
    Printer.Print "  Ship Date: " & Format$(SafeDate(rs!ShipDate), "mm/dd/yyyy")
    Printer.Print ""
    Printer.Print "  Work Order: " & SafeString(rs!WorkOrderID)
    Printer.Print "  Customer PO: " & SafeString(rs!CustomerPO)
    Printer.Print "  Weight: " & SafeString(rs!Weight) & " lbs"
    Printer.Print "  Boxes: " & SafeString(rs!NumberOfBoxes)
    
    Printer.Print ""
    Printer.Print String$(50, "-")
    Printer.FontSize = 8
    Printer.Print "Printed: " & Format$(Now, "mm/dd/yyyy hh:nn:ss AM/PM")
    
    Printer.EndDoc
    
    LogMessage "Shipping label printed (direct) for WO #" & SafeString(rs!WorkOrderID)
    
    Exit Sub
    
ErrHandler:
    ShowError "PrintShippingLabelDirect", Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: PrintInventoryReport
' Description: Prints the inventory status report
' Parameters: bPreview - Show print preview
'             sSupplierFilter - Optional supplier filter
' ============================================================================
Public Sub PrintInventoryReport(Optional ByVal bPreview As Boolean = False, _
                                 Optional ByVal sSupplierFilter As String = "")
    On Error GoTo ErrHandler
    
    SetWaitCursor True
    
    Dim sReportPath As String
    sReportPath = g_ReportsPath & "InventoryStatus.rpt"
    
    If Dir(sReportPath) = "" Then
        sReportPath = RPT_INVENTORY
    End If
    
    If Dir(sReportPath) = "" Then
        ' Fallback: Print inventory directly
        PrintInventoryReportDirect sSupplierFilter
        SetWaitCursor False
        Exit Sub
    End If
    
    InitializeCrystalReports
    
    If m_CrystalApp Is Nothing Then
        PrintInventoryReportDirect sSupplierFilter
        SetWaitCursor False
        Exit Sub
    End If
    
    Set m_CrystalReport = m_CrystalApp.OpenReport(sReportPath)
    m_CrystalReport.Database.Tables(1).SetLogOnInfo "", g_DatabasePath, "", ""
    
    ' Apply supplier filter if specified
    If Len(sSupplierFilter) > 0 Then
        m_CrystalReport.RecordSelectionFormula = "{RawMaterials.Supplier} = '" & SQLSafe(sSupplierFilter) & "'"
    End If
    
    ' Set landscape orientation for inventory report
    m_CrystalReport.PaperOrientation = DMORIENT_LANDSCAPE
    
    If bPreview Or g_PrintPreviewEnabled Then
        m_CrystalReport.Preview
    Else
        m_CrystalReport.PrintOut False, g_PrinterCopies
    End If
    
    LogMessage "Inventory report printed"
    
    Set m_CrystalReport = Nothing
    
    SetWaitCursor False
    
    Exit Sub
    
ErrHandler:
    SetWaitCursor False
    
    On Error Resume Next
    Set m_CrystalReport = Nothing
    
    ShowError "PrintInventoryReport", Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: PrintInventoryReportDirect
' Description: Prints inventory report directly to printer
' ============================================================================
Private Sub PrintInventoryReportDirect(Optional ByVal sSupplierFilter As String = "")
    On Error GoTo ErrHandler
    
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT MaterialID, MaterialName, Supplier, QuantityOnHand, " & _
           "ReorderPoint, UnitCost, LastOrderDate " & _
           "FROM RawMaterials"
    
    If Len(sSupplierFilter) > 0 Then
        sSQL = sSQL & " WHERE Supplier = '" & SQLSafe(sSupplierFilter) & "'"
    End If
    
    sSQL = sSQL & " ORDER BY MaterialName"
    
    Set rs = GetRecordset(sSQL)
    
    If rs Is Nothing Then
        MsgBox "No inventory data available.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    
    ' Print using Printer object
    Printer.Orientation = vbPRORLandscape
    Printer.FontName = "Courier New"
    
    ' Header
    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.Print COMPANY_NAME & " - Inventory Status Report"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print "Printed: " & Format$(Now, "mm/dd/yyyy hh:nn:ss AM/PM")
    If Len(sSupplierFilter) > 0 Then
        Printer.Print "Filtered by Supplier: " & sSupplierFilter
    End If
    Printer.Print String$(100, "-")
    
    ' Column headers
    Printer.FontBold = True
    Printer.Print PadRight("Material ID", 12) & _
                  PadRight("Material Name", 30) & _
                  PadRight("Supplier", 20) & _
                  PadLeft("Qty On Hand", 12) & _
                  PadLeft("Reorder Pt", 12) & _
                  PadLeft("Unit Cost", 12)
    Printer.FontBold = False
    Printer.Print String$(100, "-")
    
    ' Data rows
    Dim lCount As Long
    Do While Not rs.EOF
        Printer.Print PadRight(SafeString(rs!MaterialID), 12) & _
                      PadRight(SafeString(rs!MaterialName), 30) & _
                      PadRight(SafeString(rs!Supplier), 20) & _
                      PadLeft(FormatQuantity(SafeLong(rs!QuantityOnHand)), 12) & _
                      PadLeft(FormatQuantity(SafeLong(rs!ReorderPoint)), 12) & _
                      PadLeft(FormatCurrencyValue(SafeNumber(rs!UnitCost)), 12)
        
        lCount = lCount + 1
        rs.MoveNext
        DoEvents
    Loop
    
    Printer.Print String$(100, "-")
    Printer.Print "Total Records: " & lCount
    
    Printer.EndDoc
    
    rs.Close
    Set rs = Nothing
    
    LogMessage "Inventory report printed (direct), " & lCount & " records"
    
    Exit Sub
    
ErrHandler:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    ShowError "PrintInventoryReportDirect", Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: PrintQualityCheckReport
' Description: Prints a quality check report
' ============================================================================
Public Sub PrintQualityCheckReport(ByVal lCheckID As Long, Optional ByVal bPreview As Boolean = False)
    On Error GoTo ErrHandler
    
    SetWaitCursor True
    
    Dim sReportPath As String
    sReportPath = g_ReportsPath & "QualityCheck.rpt"
    
    If Dir(sReportPath) = "" Then
        sReportPath = RPT_QUALITY_CHECK
    End If
    
    If Dir(sReportPath) = "" Then
        MsgBox "Quality Check report template not found." & vbCrLf & _
               "Expected: " & sReportPath, vbExclamation, APP_TITLE
        SetWaitCursor False
        Exit Sub
    End If
    
    InitializeCrystalReports
    
    If m_CrystalApp Is Nothing Then
        SetWaitCursor False
        Exit Sub
    End If
    
    Set m_CrystalReport = m_CrystalApp.OpenReport(sReportPath)
    m_CrystalReport.Database.Tables(1).SetLogOnInfo "", g_DatabasePath, "", ""
    m_CrystalReport.RecordSelectionFormula = "{QualityChecks.CheckID} = " & lCheckID
    
    If bPreview Or g_PrintPreviewEnabled Then
        m_CrystalReport.Preview
    Else
        m_CrystalReport.PrintOut False, 1
    End If
    
    LogMessage "Quality Check report printed for Check #" & lCheckID
    
    Set m_CrystalReport = Nothing
    SetWaitCursor False
    
    Exit Sub
    
ErrHandler:
    SetWaitCursor False
    On Error Resume Next
    Set m_CrystalReport = Nothing
    ShowError "PrintQualityCheckReport", Err.Description, Err.Number
End Sub

' ============================================================================
' Sub: SetupPrinter
' Description: Configures printer settings
' ============================================================================
Public Sub SetupPrinter(Optional ByVal iOrientation As Integer = DMORIENT_PORTRAIT, _
                         Optional ByVal iCopies As Integer = 1)
    On Error GoTo ErrHandler
    
    g_PrinterOrientation = iOrientation
    g_PrinterCopies = iCopies
    
    ' Store default printer name
    g_DefaultPrinter = GetDefaultPrinterName()
    
    LogMessage "Printer configured: " & g_DefaultPrinter & _
               " Orientation=" & iOrientation & " Copies=" & iCopies
    
    Exit Sub
    
ErrHandler:
    LogMessage "SetupPrinter Error: " & Err.Description
End Sub

' ============================================================================
' Sub: ShowPrinterDialog
' Description: Shows the Windows printer selection dialog
' ============================================================================
Public Sub ShowPrinterDialog()
    On Error GoTo ErrHandler
    
    Dim dlg As Object
    Set dlg = CreateObject("MSComDlg.CommonDialog")
    
    dlg.CancelError = True
    dlg.PrinterDefault = True
    dlg.ShowPrinter
    
    g_DefaultPrinter = Printer.DeviceName
    g_PrinterCopies = Printer.Copies
    
    LogMessage "Printer changed to: " & g_DefaultPrinter
    
    Set dlg = Nothing
    
    Exit Sub
    
ErrHandler:
    ' User cancelled or error
    Set dlg = Nothing
    If Err.Number <> 32755 Then ' Not a cancel
        LogMessage "ShowPrinterDialog Error: " & Err.Description
    End If
End Sub
