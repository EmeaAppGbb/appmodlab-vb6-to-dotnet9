---
title: "VB6 to .NET 9 Modernization"
description: "Transform VB6 desktop apps to .NET 9 WPF/Blazor with cloud architecture"
authors: ["marconsilva"]
category: "Code Modernization"
industry: "Manufacturing"
services: ["Azure App Service", "Azure SQL Database", "Azure Blob Storage"]
languages: [".NET"]
frameworks: ["Blazor", "WPF", "Entity Framework Core"]
modernizationTools: []
agenticTools: []
tags: ["vb6-migration", "desktop-to-web", "access-to-sql", "legacy-modernization"]
extensions: ["github.copilot"]
thumbnail: "assets/screenshots/01-project-structure.png"
video: ""
version: "1.0.0"
screenshots:
  - path: "assets/screenshots/01-project-structure.png"
    alt: "VB6 Project Structure"
    description: "PrecisionParts VB6 project file tree showing 8 forms, 4 modules, 3 classes, and 3 Crystal Reports"
  - path: "assets/screenshots/02-frm-login.png"
    alt: "frmLogin.frm source code"
    description: "VB6 Login form with plain-text password authentication against Access database"
  - path: "assets/screenshots/03-frm-main.png"
    alt: "frmMain.frm source code"
    description: "MDI parent form with menu bar, status bar, and role-based permission enforcement"
  - path: "assets/screenshots/04-frm-inventory.png"
    alt: "frmInventory.frm source code"
    description: "Inventory management form using MSFlexGrid, bubble sort, and DoEvents anti-pattern"
  - path: "assets/screenshots/05-mod-globals.png"
    alt: "modGlobals.bas source code"
    description: "Global variables module with 50+ Public declarations, constants, and registry-based config"
  - path: "assets/screenshots/06-mod-database.png"
    alt: "modDatabase.bas source code"
    description: "Database module with ADO/Jet OLEDB connection management and SQL helpers"
  - path: "assets/screenshots/07-cls-workorder.png"
    alt: "clsWorkOrder.cls source code"
    description: "Work order class mixing business logic with data access, hardcoded labor rates"
  - path: "assets/screenshots/08-anti-patterns.png"
    alt: "Legacy anti-patterns analysis"
    description: "Anti-patterns found in VB6 codebase vs .NET 9 modernization targets"
---

## Legacy Application Overview

The PrecisionParts VB6 application is a desktop manufacturing management system built on Visual Basic 6 with a Microsoft Access backend. Since VB6 requires the legacy IDE to run, the screenshots below document the project structure, key forms, modules, and anti-patterns identified during analysis.

### Project Structure

The VB6 project contains 8 forms, 4 modules, 3 classes, and 3 Crystal Reports — a typical late-1990s desktop application architecture.

![VB6 Project Structure](assets/screenshots/01-project-structure.png)

### Login Form (frmLogin.frm)

The login form authenticates users with plain-text password comparison against the Access database — no hashing or salting is used.

![frmLogin.frm source code](assets/screenshots/02-frm-login.png)

### Main MDI Form (frmMain.frm)

The MDI parent form provides the application shell with a menu bar, status bar, and role-based permission enforcement controlling access to child forms.

![frmMain.frm source code](assets/screenshots/03-frm-main.png)

### Inventory Management (frmInventory.frm)

The inventory form uses MSFlexGrid for data display, implements a manual bubble sort algorithm, and contains the `DoEvents` anti-pattern for UI responsiveness.

![frmInventory.frm source code](assets/screenshots/04-frm-inventory.png)

### Global Variables Module (modGlobals.bas)

Over 50 `Public` declarations, application-wide constants, and registry-based configuration — a common VB6 pattern that creates tight coupling across the entire application.

![modGlobals.bas source code](assets/screenshots/05-mod-globals.png)

### Database Module (modDatabase.bas)

The database module manages ADO connections through Jet OLEDB providers with inline SQL helpers — all data access is string-concatenated without parameterization.

![modDatabase.bas source code](assets/screenshots/06-mod-database.png)

### Work Order Class (clsWorkOrder.cls)

The work order class mixes business logic with direct data access and contains hardcoded labor rates — violating separation of concerns.

![clsWorkOrder.cls source code](assets/screenshots/07-cls-workorder.png)

### Anti-Patterns Analysis

A comprehensive mapping of legacy anti-patterns found in the VB6 codebase against their .NET 9 modernization targets.

![Legacy anti-patterns analysis](assets/screenshots/08-anti-patterns.png)

---

## Modernization Solution

The `solution-final` branch contains the complete .NET 9 modernization of the PrecisionParts VB6 application, built step-by-step using GitHub Copilot CLI.

### Solution Architecture
- **PrecisionParts.Web** — Blazor Server app (Dashboard, Work Orders, Inventory, QC, Shipping, Parts)
- **PrecisionParts.Core** — Domain models (11 entities, 5 enums) and service interfaces
- **PrecisionParts.Infrastructure** — EF Core 9 DbContext, service implementations, database seeder
- **PrecisionParts.Tests** — xUnit test project

### Step-by-Step Tags
| Tag | Step | Description |
|-----|------|-------------|
| `step-01-explore-vb6` | Explore Legacy VB6 | Analyzed 8 forms, 4 modules, 3 classes, 10 tables |
| `step-02-analyze-logic` | Analyze Business Logic | Extracted rules: cost calc, reorder logic, QC workflow |
| `step-03-design-architecture` | Design Architecture | Planned Blazor Server + EF Core clean architecture |
| `step-04-create-project` | Create .NET Project | Scaffolded 4-project solution targeting net9.0 |
| `step-05-migrate-models` | Migrate Data Models | 11 C# entities, 5 enums, EF Core Fluent API config |
| `step-06-migrate-logic` | Migrate Business Logic | 5 service interfaces + implementations with DI |
| `step-07-migrate-ui` | Migrate UI | 6 Blazor pages replacing VB6 forms |
| `step-08-build-validate` | Build & Validate | Solution builds with 0 errors, 0 warnings |

### Key Modernization Outcomes
| VB6 Legacy | .NET 9 Modern |
|------------|---------------|
| Plain-text passwords | BCrypt-ready hashed storage |
| 50+ global variables | Scoped DI services |
| String-concatenated SQL | EF Core parameterized queries |
| Hardcoded $45/hr labor rate | Configurable via appsettings.json |
| Crystal Reports 8.5 | HTML tables (extensible to QuestPDF) |
| Windows Registry config | appsettings.json + IOptions pattern |
| Access .mdb database | EF Core 9 (SQL Server / InMemory) |
| On Error GoTo | Structured try/catch + ILogger |
