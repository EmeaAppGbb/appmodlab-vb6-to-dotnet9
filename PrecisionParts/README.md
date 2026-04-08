# Precision Parts Manufacturing System

## VB6 Legacy Application - Version 2.1.5

**Company:** Precision Parts Manufacturing Inc.  
**Created:** January 2001  
**Last Updated:** March 2024

---

## Overview

This is a legacy Visual Basic 6.0 desktop application that has been the backbone of shop floor operations for Precision Parts Manufacturing since 2001. The system manages:

- Raw material inventory tracking
- Production work order management
- Quality control inspections
- Shipping manifest generation
- Crystal Reports for printing labels and documents

## System Requirements

### Development/Runtime Environment
- **Operating System:** Windows XP SP3 or later (Windows 10/11 compatible)
- **Runtime:** Visual Basic 6.0 Runtime (SP6)
- **Database:** Microsoft Access 2003 or later (Jet 4.0 engine)
- **ActiveX Controls:**
  - Microsoft Common Controls 6.0 (MSCOMCTL.OCX)
  - Microsoft FlexGrid Control 6.0 (MSFLXGRD.OCX)
  - Microsoft Common Dialog Control 6.0 (COMDLG32.OCX)
- **Data Access:** Microsoft ActiveX Data Objects 2.5 (ADO)
- **Reporting:** Crystal Reports 8.5 Runtime

### Hardware Requirements
- **Processor:** Pentium III 500MHz or better
- **RAM:** 256 MB minimum, 512 MB recommended
- **Disk Space:** 50 MB for application, additional space for database
- **Display:** 1024x768 or higher resolution
- **Network:** Network access to shared database file (multi-user support)

## Installation

1. **Install VB6 Runtime Components**
   - Install Visual Basic 6.0 SP6 Runtime files
   - Register required ActiveX controls (MSCOMCTL.OCX, MSFLXGRD.OCX)

2. **Install Microsoft Access Database Engine**
   - Install Jet 4.0 OLEDB provider (included with Windows or Office)

3. **Install Crystal Reports Runtime**
   - Install Crystal Reports 8.5 runtime components

4. **Configure Application**
   - Copy PrecisionParts.exe to desired location (e.g., C:\PrecisionParts\)
   - Place PrecisionParts.mdb in accessible location
   - Update registry settings (see Configuration section)

5. **First Run**
   - Launch PrecisionParts.exe
   - Login: Username: `admin`, Password: `admin123`
   - Configure database path in File > Settings

## Database Structure

### Tables
- **Users** - User authentication (plain text passwords - legacy)
- **Parts** - Part master data (part numbers, descriptions, specifications)
- **RawMaterials** - Inventory tracking with reorder points
- **Customers** - Customer information
- **WorkOrders** - Production work orders
- **QualityChecks** - QC inspection records
- **ShippingManifests** - Shipping documents
- **ProductionLog** - Real-time production tracking
- **MachineStatus** - Manufacturing equipment status
- **ShiftSchedule** - Shift scheduling

### Database Location
Default: `C:\PrecisionParts\Database\PrecisionParts.mdb`

Stored in registry: `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\PrecisionParts\Database\DatabasePath`

## Application Structure

```
PrecisionParts/
├── PrecisionParts.vbp          # Visual Basic 6 project file
├── PrecisionParts.vbw          # Workspace file
├── Forms/                      # UI Forms
│   ├── frmLogin.frm           # Login authentication
│   ├── frmMain.frm            # MDI parent with menu
│   ├── frmInventory.frm       # Inventory management grid
│   ├── frmWorkOrder.frm       # Work order entry
│   ├── frmQualityCheck.frm    # QC inspection form
│   ├── frmShipping.frm        # Shipping manifest
│   ├── frmPartLookup.frm      # Part search dialog
│   └── frmReports.frm         # Report selection
├── Modules/                    # Business logic modules
│   ├── modDatabase.bas        # ADO connection/data access
│   ├── modGlobals.bas         # Global variables (50+)
│   ├── modPrinting.bas        # Crystal Reports integration
│   └── modUtilities.bas       # Helper functions
├── Classes/                    # Business object classes
│   ├── clsWorkOrder.cls       # Work order business logic
│   ├── clsInventory.cls       # Inventory calculations
│   └── clsPart.cls            # Part validation
├── Reports/                    # Crystal Reports templates
│   ├── rptWorkOrder.rpt       # Work order traveler
│   ├── rptShippingLabel.rpt   # Shipping labels
│   └── rptInventoryStatus.rpt # Inventory report
├── Database/
│   ├── PrecisionParts.mdb     # Access database
│   ├── schema.sql             # Database schema
│   └── seed_data.sql          # Sample data
└── Resources/
    └── icons/                  # UI icons (BMP format)
```

## Features

### Login & Security
- User authentication against Users table
- Role-based access: Administrator, Manager, Operator, QC, Shipping
- **Security Warning:** Passwords stored in plain text (legacy anti-pattern)

### Inventory Management
- Real-time inventory levels for raw materials
- Automatic reorder point alerts
- Supplier filtering and reporting
- Export to CSV functionality

### Work Order Management
- Create, edit, delete work orders
- Link to parts and customers
- Track status: New → In Progress → Completed
- Priority levels: Low, Normal, High
- Due date tracking with business day calculations

### Quality Control
- QC inspection recording
- Pass/Fail/Conditional/Recheck results
- Measurement data entry
- Link to work orders

### Shipping
- Shipping manifest creation
- Carrier selection (UPS, FedEx, USPS, DHL)
- Tracking number entry
- Print shipping labels via Crystal Reports
- Auto-update work order status to Completed

### Reporting
- Work Order Traveler sheets
- Shipping labels with barcodes
- Inventory status reports
- Print preview and direct printing

### Real-Time Monitoring
- Auto-refresh timer (configurable interval)
- Machine status dashboard
- Production log tracking

## Configuration

### Registry Settings
Application settings are stored in Windows Registry:

**Base Key:** `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\PrecisionParts`

Settings:
- `Database\DatabasePath` - Path to .mdb file
- `Database\BackupPath` - Backup location
- `UI\RefreshInterval` - Auto-refresh seconds
- `UI\ShowSplash` - Show splash screen (0/1)
- `System\EnableLogging` - Enable logging (0/1)
- `System\LogPath` - Log file directory
- `Reports\Path` - Crystal Reports location

### Modifying Settings
Use File > Settings menu or manually edit registry

## User Roles & Permissions

| Role | Work Orders | Inventory | QC | Shipping | Reports | Admin |
|------|-------------|-----------|-----|----------|---------|-------|
| Administrator | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Manager | ✓ | ✓ | ✓ | ✓ | ✓ | - |
| Operator | Edit | View | - | - | View | - |
| Quality Control | View | View | ✓ | - | ✓ | - |
| Shipping | View | View | - | ✓ | ✓ | - |

## Default Users

| Username | Password | Role |
|----------|----------|------|
| admin | admin123 | Administrator |
| manager1 | manager | Manager |
| operator1 | operator | Operator |
| qc1 | quality | Quality Control |
| shipping1 | ship123 | Shipping |

⚠️ **Change passwords immediately in production environments**

## Legacy Anti-Patterns Present

This application intentionally demonstrates common VB6 anti-patterns for educational purposes:

1. **Global Variables** - 50+ global variables in modGlobals.bas
2. **GoTo Error Handling** - `On Error GoTo ErrHandler` throughout
3. **ADO Recordsets as Business Objects** - Direct recordset manipulation
4. **File-Based Database** - Access MDB on network share
5. **Plain Text Passwords** - No encryption or hashing
6. **Registry Configuration** - Settings in Windows Registry
7. **ActiveX Dependencies** - MSFlexGrid, CommonDialog controls
8. **Crystal Reports 8.5** - Outdated reporting engine
9. **Single-Threaded UI** - DoEvents for responsiveness
10. **Hardcoded Paths** - File paths embedded in code
11. **No Unit Testing** - No automated test capability
12. **COM/ActiveX Interop** - External COM dependencies

## Known Issues

1. **Multi-user concurrency** - File locking can cause conflicts with multiple users
2. **Database corruption** - Network interruptions can corrupt Access database
3. **32-bit only** - Will not run on 64-bit Office installations
4. **ActiveX registration** - Controls must be registered on each machine
5. **Crystal Reports compatibility** - Reports may not work on newer systems
6. **No data validation** - Limited input validation allows bad data
7. **Memory leaks** - Long-running sessions can consume memory
8. **No audit trail** - Changes not logged or tracked

## Troubleshooting

### Application won't start
- Verify VB6 runtime is installed
- Register ActiveX controls: `regsvr32 MSCOMCTL.OCX`
- Check database path in registry

### Database connection errors
- Verify database file exists at configured path
- Check network connectivity (if database on share)
- Ensure Jet 4.0 OLEDB provider is installed

### Crystal Reports errors
- Install Crystal Reports 8.5 runtime
- Verify .rpt files exist in Reports folder
- Check report database connection settings

### Permission errors
- Login with administrator account
- Verify user role in Users table
- Check file system permissions on database

## Backup & Maintenance

### Database Backup
- Use File > Backup Database menu
- Creates timestamped copy in backup folder
- Recommended: Daily backups before business operations

### Log Files
Located in: `C:\PrecisionParts\Logs\PrecisionParts_YYYYMMDD.log`
- Rotation: New file daily
- Contains: Database operations, errors, user actions
- Review regularly for errors

### Performance
- Compact & Repair database monthly (Access Tools)
- Archive old work orders annually
- Reindex database tables quarterly

## Migration Notes

This application is a candidate for modernization to .NET 9. Key migration considerations:

1. **Database:** Migrate Access MDB to SQL Server/Azure SQL
2. **UI:** Convert to WPF (desktop) or Blazor (web)
3. **Data Access:** Replace ADO with Entity Framework Core
4. **Reporting:** Replace Crystal Reports with QuestPDF or similar
5. **Configuration:** Move registry settings to appsettings.json
6. **Security:** Implement ASP.NET Core Identity with hashed passwords
7. **Global State:** Replace global variables with dependency injection
8. **Error Handling:** Use try/catch instead of On Error GoTo

## Support

**Original Development:** Internal IT Department, 2001  
**Current Maintenance:** Legacy support only  
**Status:** Scheduled for modernization to .NET 9

For modernization lab instructions, see: `APPMODLAB.md`

---

**Copyright © 2001-2024 Precision Parts Manufacturing Inc.**  
**Legacy Application - For Educational Purposes**
