# <img src="https://raw.githubusercontent.com/PowerShell/PowerShell/master/assets/ps_black_64.svg" width="28" align="center"/> PowerShell Module Manager `v7.0`

> **A dark-themed WPF GUI for managing PowerShell modules across Windows PowerShell 5.1 and PowerShell 7 — built entirely in PowerShell.**

---

## 📋 Table of Contents

- [Overview](#overview)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [Features](#features)
- [Tabs](#tabs)
- [Architecture Notes](#architecture-notes)
- [Files](#files)

---

## Overview

`PSModuleManager.ps1` is a single-file PowerShell script that launches a full WPF GUI for managing, scanning, installing, and maintaining PowerShell modules. It runs on a dedicated STA thread via a PowerShell runspace and communicates between the GUI and background workers using thread-safe concurrent queues.

```
Author  : Nikolaos Karanikolas
Version : 7.0
Engines : Windows PowerShell 5.1  ·  PowerShell 7.x
Theme   : Dark (switchable to Light)
```

---

## Requirements

| Requirement | Detail |
|---|---|
| **OS** | Windows 10 / 11 |
| **PowerShell** | Windows PowerShell 5.1 (built-in) |
| **PS7 (optional)** | PowerShell 7.x — `pwsh.exe` in `$PATH` or default install path |
| **WPF** | .NET Framework 4.x (included with Windows) |
| **Admin** | Optional — required for `AllUsers` scope installs |
| **Internet** | Required for Gallery checks and module installs |

---

## Quick Start

```powershell
# Run directly (no install needed)
pwsh.exe -File PSModuleManager.ps1

# Or from Windows PowerShell
powershell.exe -File PSModuleManager.ps1
```

> On first run the app auto-detects both PS engines and loads the module catalog.

---

## Features

### 🔍 Module Catalog — Scan & Status
- Scans **65 predefined modules** across both PS engines in a single batch call per engine
- Color-coded status columns: `PS5 Ver`, `PS5 Scope`, `PS7 Ver`, `PS7 Scope`, `Gallery`, `Status`
- Status values: `Up to Date` · `Update Available` · `Not Installed` · `Pending`
- Scope color coding: `CurrentUser` (green) · `AllUsers` (blue) · `System` (grey) · `WinPS-System` · `PS7-System`
- Category chip tabs: **All · ActiveDir · Azure · Database · Graph · Security · System · Terminal · Utilities · VMware**
- Filter by text, installed-only, or updates-only
- Select All / Clear / Select Updatable / Select Missing

### 📦 Install / Update
- **Install / Update Selected** — installs or updates checked modules via selected engine
- **Update ALL Installed** — batch-updates every installed catalog module
- **Remove Selected** — uninstalls selected modules
- Scope selection: `AllUsers` (requires Admin) or `CurrentUser`
- Real-time progress in **Terminal Output** panel

### 🏭 Repository Management (`Repositories` tab)
| Feature | Detail |
|---|---|
| Toggle switches | Register / unregister repos with an ON/OFF pill switch |
| Known repos | PSGallery · NuGet · Chocolatey — with OFFICIAL badge |
| Trust badge | TRUSTED / UNTRUSTED badge per registered repo |
| Register custom repo | Dialog for Name + Source URL + Policy (NuGet v2/v3 feeds, Nexus, Artifactory, ProGet) |
| Unregister | Click card to select, then Unregister Selected |
| Auto-refresh | Tab refreshes automatically when selected |
| Browse Repo Modules | Lists up to 500 modules from all registered repos in the right panel with checkboxes |
| Search Gallery | Async keyword search — results shown as checkboxes in right panel |
| Install from results | Tick modules → Install Selected → choose AllUsers / CurrentUser |
| Auto-add to catalog | Newly installed repo modules are automatically added to Module Catalog |

### 🛠️ Utilities
| Button | Function |
|---|---|
| `Set PSModulePath` | Add/remove paths from `$env:PSModulePath` permanently |
| `Move Module` | Relocate a module between CurrentUser and AllUsers |
| `Open Module Folder` | Opens the module's install folder in Explorer |
| `Disable Module` | Renames `.psd1` → `.psd1.disabled` to deactivate without uninstalling |
| `Enable Module` | Restores `.psd1.disabled` → `.psd1` |
| `Export Module List` | Exports full status to `.txt` or `.html` report |
| `Clean ALL Modules` | Removes all PSGet-installed modules (with confirmation) |
| `Clean Old Versions` | Detects and removes duplicate old versions, keeping the latest |

### ℹ️ Engine Info tab
- Detects PS5 and PS7 engine paths, versions, PSModulePath entries
- Installs PS7 via winget if not found

### 📂 Module Paths tab
- Lists all PSModulePath entries per engine
- Highlights missing or duplicate paths

### 📄 Log Viewer tab
- Live log display with reload, clear, open-in-Notepad
- **Set Log Path** — choose custom folder for log file
- Timestamped entries: `[INFO]` `[SUCCESS]` `[WARN]` `[ERROR]`

---

## Tabs

```
┌─────────────────┬────────────┬─────────────┬──────────────┬──────────────┐
│  Module Catalog │ Log Viewer │ Engine Info │ Module Paths │ Repositories │
└─────────────────┴────────────┴─────────────┴──────────────┴──────────────┘
```

| Tab | Purpose |
|---|---|
| **Module Catalog** | Main scan view — status, install, update, remove |
| **Log Viewer** | Timestamped log with Set Log Path |
| **Engine Info** | PS5/PS7 detection, PS7 install via winget |
| **Module Paths** | PSModulePath editor |
| **Repositories** | Repo registration, Browse/Search, Install from repo |

---

## Architecture Notes

| Component | Implementation |
|---|---|
| **GUI Thread** | Dedicated STA runspace (`PowerShell.Create()` + `RunspaceFactory`) |
| **Scan Worker** | MTA runspace — single batch `Get-InstalledModule` + `Get-Module -ListAvailable` call per engine |
| **Repo Workers** | MTA runspace per operation — PS7 engine preferred for `Get-PSRepository` / `Register-PSRepository` |
| **Timer Polling** | `DispatcherTimer` polls `ConcurrentQueue<string>` between worker and UI thread |
| **Closures** | All timer tick handlers use `.GetNewClosure()` to capture outer scope variables |
| **Persistence** | Custom modules saved to `PSModuleManager_custom.json` next to the script |
| **Theming** | `BrushConverter` + `ApplyTheme` function — Dark / Light |

### Threading Model

```
Main Thread (PS5/PS7)
    │
    └── GUI Runspace (STA Thread)
            │
            ├── ConcurrentQueue<string>  ← Scan results
            ├── ConcurrentQueue<string>  ← Terminal output
            ├── ConcurrentQueue<object>  ← Batch results
            │
            ├── Scan Runspace (MTA)      → Get-InstalledModule / Get-Module
            ├── Repo Runspace (MTA)      → Get/Register/Unregister-PSRepository
            ├── Install Runspace (MTA)   → Install-Module
            └── Browse Runspace (MTA)    → Find-Module
```

---

## Files

```
PSModuleManager.ps1           ← Main script (single file, ~4600 lines)
PSModuleManager_custom.json   ← Auto-created — persists custom catalog modules
PSModMgr_YYYYMMDD_HHMMSS.log  ← Auto-created in %TEMP% — runtime log
```

---

## Keyboard & UI Tips

| Action | How |
|---|---|
| Filter modules | Type in the **Filter** box (instant) |
| Select all visible | **Select All** button |
| Multi-select rows | Click checkbox column |
| Resize panels | Drag the **GridSplitter** dividers |
| Change engine | **Engine** dropdown — affects Install/Remove scope |
| Change theme | **Theme** dropdown (top-right) |
| View full log | **Log Viewer** tab → Open in Notepad |

---

*Developed by [karanik](https://karanik.gr) — PowerShell Module Manager v7.0*
