A professional WPF GUI for managing PowerShell modules across Windows PowerShell 5.1 and PowerShell 7.
Overview
Single-file .ps1 that launches a full WPF desktop application for IT professionals. Provides a unified interface to scan, install, update, move, and manage PowerShell modules across both PS5.1 and PS7 simultaneously.
Features
Module Catalog & Scanning

Catalog of 65+ curated modules (expandable) across 9 categories: Azure, Graph, Security, System, Database, Terminal, Utilities, VMware, Custom
Batch scan: only 3 process spawns for all modules (PS5 + PS7 + Gallery)
Per-engine version tracking: PS5 Ver / Scope / PS7 Ver / Scope / Gallery / Status
Color-coded scope: AllUsers CurrentUser System WinPS-System PS7-System Unknown
Category filter chips, text search, Installed-only / Updates-only filters

Install & Update

Smart Dual-Engine Batch routing: PS7 modules → pwsh.exe, PS5-only → powershell.exe
Auto-removes old versions after update; auto-refresh confirms result
Clean Old Versions: scans both engines for duplicate versions, shows list, removes old with Uninstall-Module -RequiredVersion + deletes leftover folders

Module Lifecycle

Move Module: CurrentUser ↔ AllUsers, cleans old scope completely
Disable / Enable: renames folder to .disabled without uninstalling
Remove Selected: Uninstall-Module + deletes files from all known paths
Clean ALL Modules: full wipe with double confirmation
Open Module Folder: opens in Explorer

Path Management

Set PSModulePath: folder picker, session + persistent user env
Module Paths tab: card view of all $env:PSModulePath entries, module counts, Open in Explorer, Add/Remove paths

Repositories tab (new in v7.0)

Shows PSGallery, NuGet, Chocolatey always, with toggle switch (CheckBox) to register/unregister each
Each repo shows: registration status, TRUSTED/UNTRUSTED badge, URL, description
Register custom repos (Nexus, Artifactory, ProGet, Azure Artifacts) via dialog
Unregister custom repos; Trust/Untrust toggle per repo
Add Module to Catalog: search PSGallery, verify, pick category → appears in scan tabs immediately
Remove from Catalog: removes custom modules (built-in protected)
Custom modules persist to PSModuleManager_custom.json next to script

Engine Info

GitHub API check for latest PS7 and winget
One-click install/update PS7 via winget

Logging

Three-channel: Operation Log, Terminal Output (raw), log file
Tags: [SCAN] [RESULT] [BATCH] [ENGINE] [OK] [UP] [ERR] [OLDVER]

UI

Dark / Light / Auto theme, resizable panels, Info button with full feature guide
Export to CSV / HTML / TXT

Usage
powershellpwsh.exe -File PSModuleManager.ps1        # Recommended (PS7)
powershell.exe -File PSModuleManager.ps1  # Also supported (PS5.1)
Requires Administrator for AllUsers scope operations.
