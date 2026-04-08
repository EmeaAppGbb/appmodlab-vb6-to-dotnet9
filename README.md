# 🎮 VB6 → .NET 9 🚀

```
██╗   ██╗██████╗  ██████╗     ████████╗ ██████╗     ███╗   ██╗███████╗████████╗     █████╗ 
██║   ██║██╔══██╗██╔════╝     ╚══██╔══╝██╔═══██╗    ████╗  ██║██╔════╝╚══██╔══╝    ██╔══██╗
██║   ██║██████╔╝███████╗        ██║   ██║   ██║    ██╔██╗ ██║█████╗     ██║       ╚██████║
╚██╗ ██╔╝██╔══██╗██╔═══██╗       ██║   ██║   ██║    ██║╚██╗██║██╔══╝     ██║        ╚═══██║
 ╚████╔╝ ██████╔╝╚██████╔╝       ██║   ╚██████╔╝    ██║ ╚████║███████╗   ██║        █████╔╝
  ╚═══╝  ╚═════╝  ╚═════╝        ╚═╝    ╚═════╝     ╚═╝  ╚═══╝╚══════╝   ╚═╝        ╚════╝ 
```

[![.NET 9](https://img.shields.io/badge/.NET-9.0-512BD4?style=for-the-badge&logo=dotnet)](https://dotnet.microsoft.com/)
[![Blazor](https://img.shields.io/badge/Blazor-Server-8A2BE2?style=for-the-badge&logo=blazor)](https://dotnet.microsoft.com/apps/aspnet/web-apps/blazor)
[![WPF](https://img.shields.io/badge/WPF-XAML-00BFFF?style=for-the-badge)](https://learn.microsoft.com/dotnet/desktop/wpf/)
[![SQL Server](https://img.shields.io/badge/SQL%20Server-Azure-CC2927?style=for-the-badge&logo=microsoft-sql-server)](https://azure.microsoft.com/products/azure-sql/database/)

---

## 💾 OVERVIEW 💾

**Welcome to the time machine!** 🕰️✨ 

Remember when Visual Basic 6 was the *king* of RAD (Rapid Application Development)? When you could drag-drop an ActiveX control and ship enterprise software faster than dial-up could connect? When `On Error Resume Next` was... well, let's not talk about that. 😅

This lab takes you on a **radical journey** from the golden age of Windows 98 to the modern cloud era! You'll modernize a **real-world VB6 manufacturing application** that's been running shop floor operations since 2001. We're talking MSFlexGrid, ADO Recordsets, Crystal Reports, and enough global variables to make your head spin! 🌪️

**Business Domain:** Precision Parts Manufacturing 🏭  
A small manufacturing company's inventory and work order tracking system. It manages raw materials, production orders, quality inspections, and shipping — all from a trusty VB6 desktop app connected to an Access database on a network share. Classic! 💪

---

## 🎯 WHAT YOU'LL LEARN 🎯

🔥 **Assess** VB6 applications for modernization scope and complexity  
🔥 **Migrate** Microsoft Access databases to SQL Server / Azure SQL  
🔥 **Translate** VB6 forms and business logic to modern C# patterns  
🔥 **Eliminate** global state with dependency injection (goodbye `modGlobals.bas`!)  
🔥 **Choose** between desktop (WPF) and web (Blazor) modernization targets  
🔥 **Replace** ActiveX controls with modern equivalents  
🔥 **Transform** Crystal Reports into QuestPDF masterpieces  
🔥 **Deploy** to Azure App Service (Blazor) or MSIX packages (WPF)  

---

## 🕹️ PREREQUISITES 🕹️

### SETUP WIZARD 🧙‍♂️

**Required Software:**
- ✅ .NET 9 SDK ([Download](https://dotnet.microsoft.com/download/dotnet/9.0))
- ✅ Visual Studio 2022 (or VS Code with C# Dev Kit)
- ✅ SQL Server LocalDB or Azure SQL Database
- ✅ Azure subscription (for web deployment option)
- ✅ Git & GitHub CLI

**Knowledge Requirements:**
- 📘 Basic VB6 reading knowledge (no VB6 IDE needed!)
- 📗 C# and .NET experience
- 📙 Entity Framework Core fundamentals
- 📕 Blazor or WPF experience (depending on your chosen path)

**Optional Retro Gear:**
- 🎧 90s synthwave playlist (for authentic vibes)
- 🖥️ Windows 98 screensaver (just kidding... or are we?)

---

## ⚡ QUICK START ⚡

```powershell
# Clone the repo
git clone https://github.com/EmeaAppGbb/appmodlab-vb6-to-dotnet9.git
cd appmodlab-vb6-to-dotnet9

# Checkout the legacy branch to explore VB6 goodness
git checkout legacy

# Check out the step-by-step branches
git branch -r
```

**🎮 MSGBOX: "Choose Your Adventure!" 🎮**

```
┌─────────────────────────────────────────┐
│  Press [W] for WPF Desktop Path        │
│  Press [B] for Blazor Web Path         │
│                                         │
│  [ W ]     [ B ]     [ Cancel ]        │
└─────────────────────────────────────────┘
```

---

## 📁 PROJECT STRUCTURE 📁

```
PrecisionParts/
├── 🗂️ Forms/                          # VB6 Forms (*.frm)
│   ├── frmMain.frm                    # MDI parent with menu bar
│   ├── frmInventory.frm               # Raw material inventory grid
│   ├── frmWorkOrder.frm               # Work order entry/editing
│   ├── frmQualityCheck.frm            # QC inspection form
│   ├── frmShipping.frm                # Shipping manifest
│   ├── frmPartLookup.frm              # Part number search dialog
│   ├── frmReports.frm                 # Report selection and preview
│   └── frmLogin.frm                   # Simple login form
│
├── 📦 Modules/                         # Global modules (*.bas)
│   ├── modDatabase.bas                # ADO connection & helpers
│   ├── modGlobals.bas                 # 50+ global variables 😱
│   ├── modPrinting.bas                # Windows API printer magic
│   └── modUtilities.bas               # String/date/number utils
│
├── 🏗️ Classes/                         # VB6 Classes (*.cls)
│   ├── clsWorkOrder.cls               # Work order business logic
│   ├── clsInventory.cls               # Inventory calculations
│   └── clsPart.cls                    # Part number validation
│
├── 📊 Reports/                         # Crystal Reports (*.rpt)
│   ├── rptWorkOrder.rpt               # Work order printout
│   ├── rptShippingLabel.rpt           # Shipping label template
│   └── rptInventoryStatus.rpt         # Inventory status report
│
├── 💾 Database/
│   └── PrecisionParts.mdb             # Access 2003 database
│
└── 🎨 Resources/
    └── icons/                          # Toolbar icons (BMP format!)
```

---

## 🏛️ LEGACY STACK 🏛️

### 💿 THE CLASSICS 💿

**Language & Runtime:**
- 🎹 Visual Basic 6 SP6
- 🪟 Windows Common Controls 6.0
- 🔌 ActiveX Components (OCX files galore!)

**Data Access Layer:**
- 📡 **ADO 2.8** (ADODB.Connection, ADODB.Recordset)
- 💾 **Microsoft Access 2003** (MDB file on network share)
- 🔐 File-locking for concurrent access (oh boy!)

**UI Controls:**
- 📊 **MSFlexGrid** — The grid to rule them all
- 🗂️ **CommonDialog** — File picker supreme
- 🖱️ **Windows Common Controls** — Toolbar, StatusBar, TreeView

**Reporting:**
- 📄 **Crystal Reports 8.5** (unsupported vintage edition)

**System Integration:**
- 🧰 **Windows API Calls** (user32.dll, kernel32.dll)
- 📝 **Registry Settings** (SaveSetting/GetSetting)

**Database Schema:**
```
Tables: Parts, RawMaterials, WorkOrders, QualityChecks, 
        ShippingManifests, Customers, Users
```

### ⚠️ BLUE SCREEN 💙 — LEGACY ANTI-PATTERNS DETECTED!

🚨 **Global Variable Overload** — `modGlobals.bas` with 50+ public variables  
🚨 **GoTo Error Handling** — `On Error GoTo ErrHandler` everywhere  
🚨 **Network Share Database** — Access MDB with file locking  
🚨 **Plain Text Passwords** — Yes, really  
🚨 **Hardcoded Paths** — File paths and connection strings everywhere  
🚨 **DoEvents Hell** — Keeping UI responsive the old-school way  
🚨 **No Unit Tests** — Testing was for the weak! (jk, please test)  

---

## 🚀 TARGET ARCHITECTURE 🚀

### 🌟 THE FUTURE IS NOW! 🌟

**Option A: Desktop Modernization (WPF)**
```
┌──────────────────────────────────────────────┐
│  .NET 9 WPF Application                     │
│  ├─ MVVM Pattern                            │
│  ├─ WPF DataGrid (goodbye MSFlexGrid!)      │
│  ├─ Dependency Injection                    │
│  ├─ Windows Authentication                  │
│  └─ MSIX Deployment                         │
└──────────────────────────────────────────────┘
```

**Option B: Web Modernization (Blazor Server)**
```
┌──────────────────────────────────────────────┐
│  .NET 9 Blazor Server Application           │
│  ├─ Component-based UI                      │
│  ├─ SignalR Real-time Updates               │
│  ├─ ASP.NET Core Identity                   │
│  ├─ Responsive Design                       │
│  └─ Azure App Service Deployment            │
└──────────────────────────────────────────────┘
```

### 🎨 UNIFIED MODERN STACK 🎨

**Framework:**
- ⚡ **.NET 9** — The latest and greatest!
- 🏗️ **C# 13** — Modern, safe, and fast

**Data Access:**
- 🗄️ **Entity Framework Core 9** — No more ADO Recordsets!
- 💎 **Repository Pattern** — Clean separation of concerns
- 🌐 **Azure SQL Database** — Cloud-scale data storage

**Reporting:**
- 📑 **QuestPDF** — Beautiful PDF generation in C#

**Configuration:**
- ⚙️ **appsettings.json** — Goodbye registry!
- 🔧 **Azure App Configuration** (optional) — Centralized config

**Authentication:**
- 🔐 **ASP.NET Core Identity** (Blazor) — Real security!
- 🪟 **Windows Authentication** (WPF) — Domain integration

**Deployment:**
- ☁️ **Azure App Service** (Blazor path)
- 📦 **MSIX Packaging** (WPF path)

---

## 🎓 LAB WALKTHROUGH USING COPILOT CLI 🎓

### RUNTIME ERROR 💥 — Just kidding! Let's modernize! 💪

#### **Step 1: Analyze the Legacy App** 🔍
```bash
# Checkout the legacy branch
git checkout legacy

# Use Copilot CLI to explore the VB6 source
gh copilot explain "What does this VB6 code do?" --file Forms/frmMain.frm
```

**What to Look For:**
- Global state management patterns
- Database access patterns (ADO Recordsets)
- ActiveX control usage
- Business logic location (forms vs. classes)

#### **Step 2: Migrate the Database** 🗄️

**SETUP WIZARD 🧙‍♂️ — Database Migration**

```bash
# Checkout the database migration branch
git checkout step-1-database-migration

# Review migration scripts
ls Database/Migration/

# Run the migration with Copilot CLI assistance
gh copilot suggest "Create an Azure SQL Database and run migration scripts"
```

**Migration Tasks:**
- Convert Access MDB schema to SQL Server
- Migrate seed data (100+ parts, 50+ work orders)
- Add proper indexes and constraints
- Update password storage (bcrypt hashing!)

#### **Step 3: Create Domain Model** 🏗️

```bash
# Checkout the domain model branch
git checkout step-2-domain-model

# Explore EF Core entities
gh copilot explain "Show me the EF Core entity mapping for the Parts table"
```

**What's Happening:**
- ADO Recordsets → EF Core Entities
- Global variables → Dependency Injection services
- VB6 classes → C# domain models with proper encapsulation

#### **Step 4: Build UI Shell** 🖼️

**MSGBOX: "Choose Your Path!" ✅**

```bash
# For Blazor path:
git checkout step-3-ui-migration

# For WPF path:
git checkout solution-wpf
```

**UI Transformation:**
- MDI Forms → Blazor Components / WPF Views
- MSFlexGrid → Blazor Tables / WPF DataGrid
- CommonDialog → Modern file pickers

#### **Step 5: Migrate Forms** 📋

```bash
# Work through each form migration
gh copilot suggest "Convert this VB6 form to a Blazor component"
```

**Form-by-Form Journey:**
1. 🏠 **frmMain** → MainLayout/Shell
2. 📦 **frmInventory** → Inventory Component/View
3. 📝 **frmWorkOrder** → WorkOrder Component/View
4. ✅ **frmQualityCheck** → QualityCheck Component/View
5. 📮 **frmShipping** → Shipping Component/View
6. 🔍 **frmPartLookup** → PartLookup Dialog/Modal
7. 📊 **frmReports** → Reports Component/View
8. 🔐 **frmLogin** → Login Component/View

#### **Step 6: Replace Reports** 📄

```bash
# Checkout the reporting branch
git checkout step-4-reports-and-printing

# Explore QuestPDF templates
gh copilot explain "How does QuestPDF replace Crystal Reports?"
```

**Crystal → QuestPDF:**
- Work order printouts → QuestPDF documents
- Shipping labels → QuestPDF templates
- Inventory reports → QuestPDF with charts

#### **Step 7: Add Authentication** 🔐

**SETUP WIZARD 🧙‍♂️ — Security Upgrade**

```bash
# Configure ASP.NET Core Identity
gh copilot suggest "Set up ASP.NET Core Identity with role-based authorization"
```

**Security Wins:**
- Plain text passwords → Hashed with bcrypt
- File-based auth → Database-backed identity
- No roles → Role-based access control (Admin, Operator, Inspector)

#### **Step 8: Deploy to Azure** ☁️

```bash
# Checkout deployment branch
git checkout step-5-deploy

# Review Bicep templates
ls Infrastructure/

# Deploy with Copilot CLI
gh copilot suggest "Deploy Blazor app to Azure App Service with CI/CD"
```

**Deployment Stack:**
- 🌐 Azure App Service (Blazor)
- 🗄️ Azure SQL Database
- 📦 Azure Blob Storage (file attachments)
- 🔄 GitHub Actions (CI/CD)

---

## ⏱️ DURATION ⏱️

**Estimated Time:** 5–7 hours ⏰

```
┌──────────────────────────────────────────┐
│  Step 1: Analyze Legacy        (1 hr)   │
│  Step 2: Database Migration    (1 hr)   │
│  Step 3: Domain Model         (1 hr)   │
│  Step 4: UI Shell             (1 hr)   │
│  Step 5: Form Migration       (2 hrs)  │
│  Step 6: Reports              (1 hr)   │
│  Step 7: Authentication       (30 min) │
│  Step 8: Deploy               (30 min) │
└──────────────────────────────────────────┘
```

**💡 Pro Tip:** Use GitHub Copilot CLI throughout to speed up code translation and debugging!

---

## 📚 RESOURCES 📚

### 🎓 Learning Materials

**Microsoft Docs:**
- [.NET 9 Documentation](https://learn.microsoft.com/dotnet/core/whats-new/dotnet-9)
- [Blazor Documentation](https://learn.microsoft.com/aspnet/core/blazor/)
- [WPF Documentation](https://learn.microsoft.com/dotnet/desktop/wpf/)
- [Entity Framework Core 9](https://learn.microsoft.com/ef/core/)
- [QuestPDF Documentation](https://www.questpdf.com/)

**Migration Guides:**
- [VB6 to .NET Migration Guide](https://learn.microsoft.com/previous-versions/visualstudio/visual-basic-6/visual-basic-6.0-documentation)
- [Access to SQL Server Migration](https://learn.microsoft.com/sql/ssma/access/sql-server-migration-assistant-for-access-accesstosql)

**Azure Resources:**
- [Azure SQL Database](https://learn.microsoft.com/azure/azure-sql/database/)
- [Azure App Service](https://learn.microsoft.com/azure/app-service/)
- [MSIX Packaging](https://learn.microsoft.com/windows/msix/)

### 🛠️ Tools & Extensions

- **GitHub Copilot CLI** — Your AI pair programmer
- **Visual Studio 2022** — Premium IDE experience
- **Azure Data Studio** — Database management
- **SSMA for Access** — Automated database migration

### 🎯 Branch Structure Reference

```
📍 main                           # Complete lab documentation
📍 legacy                         # VB6 source + compiled EXE
📍 solution                       # Blazor implementation
📍 solution-wpf                   # WPF implementation
📍 step-1-database-migration      # Access → SQL Server
📍 step-2-domain-model            # EF Core entities
📍 step-3-ui-migration            # UI framework setup
📍 step-4-reports-and-printing    # QuestPDF implementation
📍 step-5-deploy                  # Azure deployment
```

---

## 🎊 MSGBOX: SUCCESS ✅

```
┌─────────────────────────────────────────────────────┐
│                                                     │
│   🎉 CONGRATULATIONS! 🎉                          │
│                                                     │
│   You've successfully modernized a VB6 app         │
│   to .NET 9! From ActiveX to Azure,                │
│   from Access to SQL Server,                       │
│   from Windows 98 vibes to Cloud Native glory!     │
│                                                     │
│   🏆 Achievement Unlocked:                         │
│   "Retro Modernizer" 🏆                           │
│                                                     │
│              [ OK ]                                 │
│                                                     │
└─────────────────────────────────────────────────────┘
```

---

## 🤝 CONTRIBUTING 🤝

Found a bug? Want to add more retro flair? PRs welcome! 💜

1. Fork the repo
2. Create your feature branch (`git checkout -b feature/more-synthwave`)
3. Commit your changes (`git commit -m '✨ Add more neon'`)
4. Push to the branch (`git push origin feature/more-synthwave`)
5. Open a Pull Request

---

## 📜 LICENSE 📜

This project is licensed under the MIT License.

---

## 🙏 ACKNOWLEDGMENTS 🙏

- 💙 **VB6 Community** — For keeping the legacy alive
- 🚀 **Microsoft** — For .NET 9 and the amazing evolution
- 🎮 **Synthwave Producers** — For the perfect coding soundtrack
- 🏭 **Manufacturing Teams** — Still running VB6 in production!

---

<div align="center">

### 🌟 REMEMBER: EVERY GREAT APP STARTS WITH A MSGBOX 🌟

Made with 💜 by the EMEA App GBB Team

**Now go forth and modernize!** 🚀✨

```
[ Start ]  [ Programs ]  [ Documents ]  [ Settings ]  [ Shutdown ]
```

</div>
