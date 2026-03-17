# KeePass Network Checker

A [KeePass 2.x](https://keepass.info/) plugin that performs real-time network diagnostics — ping, TCP port check, and HTTP status — directly from your password entries.

![KeePass Network Checker](https://img.shields.io/badge/KeePass-Plugin-blue) ![Version](https://img.shields.io/badge/version-1.2.0-green) ![.NET](https://img.shields.io/badge/.NET%20Framework-4.8-purple) ![License](https://img.shields.io/badge/license-GPL--3.0-orange) ![Built with Claude AI](https://img.shields.io/badge/built%20with-Claude%20AI-blueviolet?logo=anthropic)

---

## Features

- **Ping check** — ICMP round-trip time in milliseconds
- **Port check** — TCP connect to the entry's URL port (auto-detected: 443 for HTTPS, 22 for SSH, 80 for HTTP, or custom)
- **HTTP status** — GET request with SSL certificate validation bypass (useful for self-signed certs on routers, NAS devices, etc.)
- **Net Status column** — inline UP/DOWN indicator directly in the KeePass entry list
- **Entry check** — right-click any entry → Network Check
- **Group check** — right-click any group → Network Check All Group Entries (checks all entries in that group only)
- **Popup window** — detailed results table with Device, URL, Ping, Port, HTTP and Status columns
- **Configurable** — option to show or hide the popup window via Tools → Network Checker Options

---

## Screenshots

### Network Checker popup
![Network Checker popup](https://raw.githubusercontent.com/karanikn/KeePassNetworkChecker/main/Screenshots/network_checker1.png)

The popup shows detailed results per entry with color-coded status indicators.

| Column | Description |
|--------|-------------|
| Device | Entry title |
| URL    | Entry URL field |
| Ping   | ICMP round-trip time (ms) or TIMEOUT |
| Port   | TCP port connect result and latency |
| HTTP   | HTTP response code (200, 403, etc.) or ERR |
| Status | UP / DOWN composite result |

### Results with UP/DOWN status
![Network Checker results](https://raw.githubusercontent.com/karanikn/KeePassNetworkChecker/main/Screenshots/network_checker2.png)

### Group context menu
![Group context menu](https://raw.githubusercontent.com/karanikn/KeePassNetworkChecker/main/Screenshots/network_checker_group-menu.png)

Right-click any group → **Network Check All Group Entries** to check all entries in that group at once.

### Options dialog
![Options dialog](https://raw.githubusercontent.com/karanikn/KeePassNetworkChecker/main/Screenshots/network_options.png)

### Net Status column
After running a check, the `Net Status` column in the KeePass entry list is updated with UP or DOWN for each checked entry. To enable it: **View → Configure Columns → enable "Net Status"**.

---

## Requirements

| Requirement | Version |
|-------------|---------|
| KeePass | 2.x (tested on 2.61) |
| .NET Framework | 4.8 |
| Windows | 10 / 11 |

---

## Installation

### Option A — Install pre-built DLL

1. Download `KeePassNetworkChecker.dll` from the [Releases](../../releases) page
2. Copy it to your KeePass `Plugins\` folder  
   *(e.g. `C:\Program Files\KeePass Password Safe 2\Plugins\`)*
3. Restart KeePass
4. Approve the plugin in the KeePass security dialog

### Option B — Build from source

**Prerequisites:**
- Visual Studio 2019/2022 **or** [Build Tools for Visual Studio](https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022) (select `.NET desktop build tools` during install)
- .NET Framework 4.8 (included with Windows 10/11)

```powershell
# Clone the repository
git clone https://github.com/your-username/KeePassNetworkChecker.git
cd KeePassNetworkChecker

# Set the path to your KeePass.exe (required — used as build reference and install target)
$env:KEEPASS_PATH = "C:\Path\To\KeePass\KeePass.exe"

# Build and install
.\build.ps1
```

> **Tip:** Add `$env:KEEPASS_PATH` to your PowerShell profile (`$PROFILE`) so you don't have to set it every time.

---

## build.ps1 — Build & Install Script

The `build.ps1` script automates the entire build and deployment process. It requires no arguments if `$env:KEEPASS_PATH` is already set.

### What it does — step by step

| Step | Action |
|------|--------|
| 1 | Locates `KeePass.exe` from `$env:KEEPASS_PATH` or common install paths |
| 2 | Finds `MSBuild.exe` automatically via `vswhere.exe` (Visual Studio locator) |
| 3 | Cleans `bin\` and `obj\` folders from any previous build |
| 4 | Compiles the project with `MSBuild` in `Release` configuration |
| 5 | Removes any leftover `.plgx` file from the Plugins folder (prevents compile errors) |
| 6 | Copies the compiled `KeePassNetworkChecker.dll` to the KeePass `Plugins\` folder |
| 7 | Clears the KeePass plugin cache (`%LOCALAPPDATA%\KeePass\PluginCache\*`) |

### Usage

```powershell
# Standard build — uses $env:KEEPASS_PATH
.\build.ps1

# Override KeePass path inline
.\build.ps1 -KeePassPath "C:\Tools\KeePass\KeePass.exe"
```

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-KeePassPath` | string | `$env:KEEPASS_PATH` | Full path to `KeePass.exe` |
| `-Configuration` | string | `Release` | MSBuild configuration (`Release` or `Debug`) |

### Example output

```
KeePass  : C:\Tools\KeePass\KeePass.exe
MSBuild  : C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe
Building...
  KeePassNetworkChecker → bin\Release\KeePassNetworkChecker.dll
DLL      -> C:\...\bin\Release\KeePassNetworkChecker.dll
Installed -> C:\Tools\KeePass\Plugins\
Clearing plugin cache...

Done. Restart KeePass.
```

### After running the script

1. **Restart KeePass**
2. On first load, KeePass will show a **security approval dialog** — click **Yes** to allow the plugin
3. The plugin will appear in **Tools → Plugins** as *KeePass Network Checker*

## Usage

### Check a single entry
1. Select one or more entries in KeePass
2. Right-click → **Network Check (Ping / Port / HTTP)**
3. The popup opens and runs all checks automatically

### Check all entries in a group
1. Right-click on any group in the left panel
2. Select **Network Check All Group Entries**
3. The popup opens with results for all entries in that group

### Net Status column
After any check completes, the `Net Status` column is automatically updated for the checked entries. Enable it via **View → Configure Columns → Net Status**.

### Options
Go to **Tools → Network Checker Options** to toggle the popup window on/off.

---

## How it works

For each entry, the plugin performs three independent checks:

```
Entry URL: https://192.168.1.1
         │
         ├── Ping (ICMP)       → 12 ms ✓
         ├── Port check (443)  → 443 OK (18ms) ✓
         └── HTTP GET          → 200 ✓
                                  └─ Status: UP
```

- **Ping** uses ICMP with a 3-second timeout
- **Port** performs a TCP connect with a 3-second timeout
- **HTTP** performs a GET request with TLS 1.1/1.2 support and self-signed cert acceptance
- **Status** is UP if at least one of the three checks succeeds

Checks run in a `BackgroundWorker` so the KeePass UI remains fully responsive.

---

## Project structure

```
KeePassNetworkChecker/
├── KeePassNetworkChecker.cs        # Plugin entry point, menu registration
├── NetworkCheckerForm.cs           # Popup results window
├── NetworkStatusColumnProvider.cs  # Net Status column for KeePass entry list
├── SettingsForm.cs                 # Options dialog
├── Properties/
│   └── AssemblyInfo.cs             # Version and product metadata
├── KeePassNetworkChecker.csproj    # MSBuild project file
├── build.ps1                       # Build & install script
└── README.md                       # This file
```

---

## License

This project is licensed under the **GNU General Public License v3.0** — see the [LICENSE](LICENSE) file for details.

---

## Author

**Nikolaos Karanikolas**

---

## Acknowledgements

This plugin was developed with the assistance of **[Claude AI](https://www.anthropic.com/claude)** by [Anthropic](https://www.anthropic.com). The iterative development process — from initial concept through build system troubleshooting, KeePass plugin API research, UI refinement, and bug fixing — was carried out in collaboration with Claude, which helped navigate the KeePass plugin framework, resolve .NET Framework compatibility issues, and refine the feature set based on real-world feedback.
