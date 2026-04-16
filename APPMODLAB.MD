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
