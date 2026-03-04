#Requires -Version 5.1
<#
.SYNOPSIS  PowerShell Module Manager v7.0
.DESCRIPTION
    The ONLY reliable way to host WPF in a PS script:
    1. Create a dedicated STA thread
    2. Inside that thread: new Application(), then ShowDialog()
    3. All UI runs on that thread - zero dispatcher issues
    This is the same pattern used by SAPIEN PowerShell Studio.

.NOTES
    Run from any PS version:
        powershell.exe -File PSModuleManager.ps1
        pwsh.exe       -File PSModuleManager.ps1
    No -STA flag needed - the script creates its own STA thread.
#>

Set-StrictMode -Off
$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'

# ============================================================
#  LOGGING  (thread-safe via ConcurrentQueue)
# ============================================================
$Global:LogFile  = Join-Path $env:TEMP ('PSModMgr_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.log')
$Global:LogQ     = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

function Write-Log {
    param([string]$Msg, [string]$Lvl = 'INFO')
    $e = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')][$Lvl] $Msg"
    $Global:LogQ.Enqueue($e)
    try { Add-Content -Path $Global:LogFile -Value $e -Encoding UTF8 } catch {}
    $c = switch ($Lvl) {
        'ERROR'   { 'Red'    }
        'WARN'    { 'Yellow' }
        'SUCCESS' { 'Green'  }
        'DEBUG'   { 'Gray'   }
        default   { 'Cyan'   }
    }
    Write-Host $e -ForegroundColor $c
}

# ============================================================
#  MODULE CATALOG
# ============================================================
$Global:Catalog = @(
    # System
    [PSCustomObject]@{Name='PSWindowsUpdate';              Cat='System';    Desc='Manage Windows Update programmatically'}
    [PSCustomObject]@{Name='PackageManagement';            Cat='System';    Desc='Core package provider framework'}
    [PSCustomObject]@{Name='PowerShellGet';                Cat='System';    Desc='Required for PSGallery modules'}
    [PSCustomObject]@{Name='WindowsCompatibility';         Cat='System';    Desc='Import Windows modules into PS Core'}
    [PSCustomObject]@{Name='CimCmdlets';                   Cat='System';    Desc='CIM/WMI operations'}
    [PSCustomObject]@{Name='ScheduledTasks';               Cat='System';    Desc='Manage scheduled tasks'}
    [PSCustomObject]@{Name='DnsClient';                    Cat='System';    Desc='DNS client cmdlets'}
    [PSCustomObject]@{Name='NetTCPIP';                     Cat='System';    Desc='TCP/IP networking cmdlets'}
    [PSCustomObject]@{Name='TUN.CredentialManager';        Cat='System';    Desc='Windows Credential Manager wrapper'}
    [PSCustomObject]@{Name='ps2exe';                       Cat='System';    Desc='Convert PS1 scripts to EXE'}
    [PSCustomObject]@{Name='Microsoft.WinGet.Client';      Cat='System';    Desc='Programmatic Windows Package Manager'}
    [PSCustomObject]@{Name='PnP.PowerShell';               Cat='System';    Desc='SharePoint Online administration'}
    [PSCustomObject]@{Name='ExchangeOnlineManagement';     Cat='System';    Desc='Manage Exchange Online'}
    # Graph
    [PSCustomObject]@{Name='Microsoft.Graph';                              Cat='Graph'; Desc='Full Graph SDK meta-module'}
    [PSCustomObject]@{Name='Microsoft.Graph.Authentication';               Cat='Graph'; Desc='Graph authentication module'}
    [PSCustomObject]@{Name='Microsoft.Graph.Users';                        Cat='Graph'; Desc='Microsoft Graph Users'}
    [PSCustomObject]@{Name='Microsoft.Graph.Groups';                       Cat='Graph'; Desc='Microsoft Graph Groups'}
    [PSCustomObject]@{Name='Microsoft.Graph.Mail';                         Cat='Graph'; Desc='Mail/Exchange Graph endpoints'}
    [PSCustomObject]@{Name='Microsoft.Graph.Calendar';                     Cat='Graph'; Desc='Calendar operations via Graph'}
    [PSCustomObject]@{Name='Microsoft.Graph.Reports';                      Cat='Graph'; Desc='Graph Usage reports'}
    [PSCustomObject]@{Name='Microsoft.Graph.Identity.DirectoryManagement'; Cat='Graph'; Desc='Azure AD directory management'}
    [PSCustomObject]@{Name='MicrosoftTeams';                               Cat='Graph'; Desc='Full Microsoft Teams module'}
    # Azure
    [PSCustomObject]@{Name='Az';                    Cat='Azure'; Desc='Azure resource management meta-module'}
    [PSCustomObject]@{Name='Az.Accounts';           Cat='Azure'; Desc='Azure authentication and accounts'}
    [PSCustomObject]@{Name='Az.Resources';          Cat='Azure'; Desc='Azure Resource Manager operations'}
    [PSCustomObject]@{Name='Az.Network';            Cat='Azure'; Desc='Azure Networking VNet NSG LB'}
    [PSCustomObject]@{Name='Az.Compute';            Cat='Azure'; Desc='Azure Virtual Machines and VMSS'}
    [PSCustomObject]@{Name='Az.Storage';            Cat='Azure'; Desc='Azure Storage Accounts and Blobs'}
    [PSCustomObject]@{Name='Az.KeyVault';           Cat='Azure'; Desc='Azure Key Vault secrets management'}
    [PSCustomObject]@{Name='Az.Monitor';            Cat='Azure'; Desc='Azure Monitor metrics and alerts'}
    [PSCustomObject]@{Name='Az.Websites';           Cat='Azure'; Desc='Azure App Service and Web Apps'}
    [PSCustomObject]@{Name='Az.Automation';         Cat='Azure'; Desc='Azure Automation Runbooks'}
    [PSCustomObject]@{Name='Az.ContainerRegistry';  Cat='Azure'; Desc='Azure Container Registry'}
    [PSCustomObject]@{Name='Az.Dns';                Cat='Azure'; Desc='Azure DNS zones and records'}
    [PSCustomObject]@{Name='Az.Sql';                Cat='Azure'; Desc='Azure SQL Database management'}
    [PSCustomObject]@{Name='Az.OperationalInsights';Cat='Azure'; Desc='Azure Log Analytics and Sentinel'}
    [PSCustomObject]@{Name='Az.PolicyInsights';     Cat='Azure'; Desc='Azure Policy compliance'}
    [PSCustomObject]@{Name='Az.RecoveryServices';   Cat='Azure'; Desc='Azure Backup and Site Recovery'}
    [PSCustomObject]@{Name='Az.SecurityInsights';   Cat='Azure'; Desc='Microsoft Sentinel integration'}
    # Active Directory
    [PSCustomObject]@{Name='ActiveDirectory';   Cat='ActiveDir'; Desc='Active Directory management cmdlets'}
    [PSCustomObject]@{Name='ADDSDeployment';    Cat='ActiveDir'; Desc='AD DS deployment and promotion'}
    [PSCustomObject]@{Name='GroupPolicy';       Cat='ActiveDir'; Desc='Group Policy management'}
    # VMware
    [PSCustomObject]@{Name='VMware.PowerCLI';   Cat='VMware';   Desc='vSphere / vCenter / ESXi management'}
    # Security
    [PSCustomObject]@{Name='PowerForensics';    Cat='Security';  Desc='Live forensics toolkit'}
    [PSCustomObject]@{Name='PESecurity';        Cat='Security';  Desc='Check EXE/DLL for ASLR and DEP'}
    [PSCustomObject]@{Name='Defender';          Cat='Security';  Desc='Windows Defender management'}
    [PSCustomObject]@{Name='PowerUpSQL';        Cat='Security';  Desc='SQL Server audit toolkit'}
    # Database
    [PSCustomObject]@{Name='dbatools';          Cat='Database';  Desc='SQL Server automation and DBA tasks'}
    # Utilities
    [PSCustomObject]@{Name='BurntToast';        Cat='Utilities'; Desc='Windows Toast notifications'}
    [PSCustomObject]@{Name='PSReadLine';        Cat='Utilities'; Desc='Enhanced command-line editing'}
    [PSCustomObject]@{Name='ImportExcel';       Cat='Utilities'; Desc='Excel import/export without Office'}
    [PSCustomObject]@{Name='PlatyPS';           Cat='Utilities'; Desc='Generate PS module documentation'}
    [PSCustomObject]@{Name='powershell-yaml';   Cat='Utilities'; Desc='YAML format read/write manipulation'}
    [PSCustomObject]@{Name='Pester';            Cat='Utilities'; Desc='PowerShell testing framework'}
    [PSCustomObject]@{Name='PSTeams';           Cat='Utilities'; Desc='Send formatted messages to Teams'}
    [PSCustomObject]@{Name='PSTelegramAPI';     Cat='Utilities'; Desc='Telegram bot API integration'}
    [PSCustomObject]@{Name='PSWriteHTML';       Cat='Utilities'; Desc='HTML report generation'}
    [PSCustomObject]@{Name='PSWriteColor';      Cat='Utilities'; Desc='Advanced colored console output'}
    [PSCustomObject]@{Name='PoShLog';           Cat='Utilities'; Desc='Serilog wrapper for structured logging'}
    [PSCustomObject]@{Name='Pode';              Cat='Utilities'; Desc='PowerShell web framework'}
    [PSCustomObject]@{Name='ThreadJob';         Cat='Utilities'; Desc='Background thread-based jobs'}
    # Terminal
    [PSCustomObject]@{Name='oh-my-posh';        Cat='Terminal';  Desc='Best prompt and theme engine'}
    [PSCustomObject]@{Name='posh-git';          Cat='Terminal';  Desc='Git status prompt integration'}
    [PSCustomObject]@{Name='Terminal-Icons';    Cat='Terminal';  Desc='File and folder icons in terminal'}
    [PSCustomObject]@{Name='psInlineProgress';  Cat='Terminal';  Desc='Inline progress bars'}
)

# ============================================================
#  ENGINE DETECTION
# ============================================================
function Get-Engines {
    Write-Log 'Detecting PS engines...'
    $list = [System.Collections.Generic.List[object]]::new()

    $c5 = Get-Command 'powershell.exe' -ErrorAction SilentlyContinue
    if ($c5) {
        try {
            $v = & powershell.exe -NoProfile -NonInteractive -Command '$PSVersionTable.PSVersion.ToString()' 2>$null
            if ($v) {
                $list.Add([PSCustomObject]@{Tag='PS5';Label='Windows PowerShell 5.1';Exe='powershell.exe';Ver=$v.Trim();OK=$true;Path=$c5.Source})
                Write-Log "PS5 found: $($c5.Source)  v$($v.Trim())" 'SUCCESS'
            }
        } catch { Write-Log "PS5 error: $_" 'WARN' }
    }

    $ps7cands = [System.Collections.Generic.List[string]]::new()
    $c7 = Get-Command 'pwsh.exe' -ErrorAction SilentlyContinue
    if ($c7 -ne $null -and $c7.Source) { $ps7cands.Add($c7.Source) }
    $ps7cands.Add("$env:ProgramFiles\PowerShell\7\pwsh.exe")
    $ps7cands.Add("$env:ProgramFiles\PowerShell\7-preview\pwsh.exe")

    $got7 = $false
    foreach ($p in $ps7cands) {
        if ([string]::IsNullOrEmpty($p) -or -not (Test-Path $p -PathType Leaf)) { continue }
        try {
            $v = & $p -NoProfile -NonInteractive -Command '$PSVersionTable.PSVersion.ToString()' 2>$null
            if ($v) {
                $list.Add([PSCustomObject]@{Tag='PS7';Label="PowerShell $($v.Trim())";Exe=$p;Ver=$v.Trim();OK=$true;Path=$p})
                Write-Log "PS7 found: $p  v$($v.Trim())" 'SUCCESS'
                $got7 = $true; break
            }
        } catch {}
    }
    if (-not $got7) {
        Write-Log 'PS7 not found.' 'WARN'
        $list.Add([PSCustomObject]@{Tag='PS7';Label='PowerShell 7 (Not Installed)';Exe='';Ver='N/A';OK=$false;Path=''})
    }
    Write-Log "Engine detection done. Found $($list.Count) entries." 'INFO'
    return $list
}

# ============================================================
#  MODULE STATUS CHECK  (runs in separate process)
# ============================================================
function Invoke-ModCheck {
    param([string]$Name, [string]$Exe)
    $sc = (
        '$ep="SilentlyContinue";$n="' + $Name + '";' +
        '$ErrorActionPreference=$ep;$ProgressPreference=$ep;' +
        '$all=@(Get-InstalledModule -Name $n -AllVersions -EA $ep 2>$null);' +
        'if(-not $all){$all=@(Get-Module -ListAvailable -Name $n -EA $ep 2>$null)};' +
        'if($all){' +
        '  $m=$all|Sort-Object Version -Descending|Select-Object -First 1;' +
        '  $sc="Unknown";' +
        '  if($m.ModuleBase -like "*\AllUsers\*"){$sc="AllUsers"}' +
        '  elseif($m.ModuleBase -like "*\CurrentUser\*"){$sc="CurrentUser"}' +
        '  elseif($m.ModuleBase -like "*\WindowsPowerShell\*"){$sc="WinPS-System"}' +
        '  elseif($m.ModuleBase -like "*\PowerShell\*"){$sc="PS7-System"};' +
        '  Write-Output ("INST|"+$m.Version.ToString()+"|"+$sc)' +
        '}else{Write-Output "NONE"};' +
        '$g=Find-Module -Name $n -EA $ep 2>$null|Select-Object -First 1;' +
        'if($g){Write-Output ("GAL|"+$g.Version.ToString())}else{Write-Output "NOGAL"}'
    )
    try { return (& $Exe -NoProfile -NonInteractive -Command $sc 2>$null) }
    catch { return @("ERR|$_") }
}

# ============================================================
#  INSTALL / UPDATE  (runs in separate process)
# ============================================================
function Invoke-ModInstall {
    param([string]$Name, [string]$Exe, [string]$Scope)
    Write-Log "Install [$Name] via $Exe scope=$Scope" 'INFO'
    $sc = (
        '$ep="Stop";$n="' + $Name + '";$s="' + $Scope + '";' +
        '$ErrorActionPreference=$ep;$ProgressPreference="SilentlyContinue";' +
        'try{' +
        '  $i=Get-InstalledModule -Name $n -EA SilentlyContinue;' +
        '  if($i){' +
        '    $g=Find-Module -Name $n -EA SilentlyContinue;' +
        '    if($g -and ([Version]$g.Version -gt [Version]$i.Version)){' +
        '      Update-Module -Name $n -Scope $s -Force;' +
        '      Write-Output ("UPDATED|"+$g.Version)' +
        '    }else{Write-Output ("UPTODATE|"+$i.Version)}' +
        '  }else{' +
        '    Install-Module -Name $n -Scope $s -Force -AllowClobber -SkipPublisherCheck;' +
        '    $v=(Get-InstalledModule -Name $n).Version;' +
        '    Write-Output ("INSTALLED|"+$v)' +
        '  }' +
        '}catch{Write-Output ("ERROR|"+$_.Exception.Message)}'
    )
    try {
        $out = & $Exe -NoProfile -NonInteractive -Command $sc 2>&1
        foreach ($ln in $out) {
            $s = $ln.ToString().Trim()
            if ($s -match '^\w+\|') {
                Write-Log "  [$Name] result: $s" 'INFO'
                return $s
            }
            if ($s) { Write-Log "  [$Name] sub: $s" 'DEBUG' }
        }
    } catch {
        Write-Log "Fatal [$Name]: $_" 'ERROR'
        return "ERROR|$_"
    }
    return 'ERROR|No output'
}

# ============================================================
#  THE C# TYPES  (compiled once, outside GUI thread)
# ============================================================
Add-Type -AssemblyName PresentationFramework -EA Stop
Add-Type -AssemblyName PresentationCore      -EA Stop
Add-Type -AssemblyName WindowsBase           -EA Stop
Add-Type -AssemblyName System.Windows.Forms  -EA SilentlyContinue

if (-not ([System.Management.Automation.PSTypeName]'PSMod.Row').Type) {
    Add-Type -Language CSharp -TypeDefinition @'
using System.ComponentModel;
namespace PSMod {
    public class Row : INotifyPropertyChanged {
        public event PropertyChangedEventHandler PropertyChanged;
        void N(string p){var h=PropertyChanged;if(h!=null)h(this,new PropertyChangedEventArgs(p));}
        bool _sel; public bool Sel{get{return _sel;}set{if(_sel!=value){_sel=value;N("Sel");}}}
        public string Name{get;set;}
        public string Cat {get;set;}
        public string Desc{get;set;}
        string _v5;  public string V5 {get{return _v5;} set{if(_v5!=value){_v5=value;N("V5");}}}
        string _sc5; public string S5 {get{return _sc5;}set{if(_sc5!=value){_sc5=value;N("S5");}}}
        string _v7;  public string V7 {get{return _v7;} set{if(_v7!=value){_v7=value;N("V7");}}}
        string _sc7; public string S7 {get{return _sc7;}set{if(_sc7!=value){_sc7=value;N("S7");}}}
        string _gv;  public string GV {get{return _gv;} set{if(_gv!=value){_gv=value;N("GV");}}}
        string _st;  public string St {get{return _st;} set{if(_st!=value){_st=value;N("St");}}}
        bool   _up;  public bool   Up {get{return _up;} set{if(_up!=value){_up=value;N("Up");}}}
    }
}
'@
}

Write-Log 'Assemblies and types compiled OK.' 'SUCCESS'

# ============================================================
#  GUI  - runs on its own STA thread
# ============================================================
$Global:Engines   = Get-Engines
$Global:AllRows   = [System.Collections.Generic.List[PSMod.Row]]::new()
foreach ($m in $Global:Catalog) {
    $r = [PSMod.Row]::new()
    $r.Name = $m.Name; $r.Cat = $m.Cat; $r.Desc = $m.Desc; $r.St = 'Pending'
    $Global:AllRows.Add($r)
}

# Shared state between GUI thread and scan runspace
$Global:ScanResults  = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
$Global:TermQueue    = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$Global:BatchResults = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
$Global:CancelScan   = $false
$Global:CancelBatch  = $false

# ============================================================
#  CUSTOM MODULES - persisted to JSON next to script
# ============================================================
$Global:CustomModulesFile = Join-Path (Split-Path $MyInvocation.MyCommand.Path) 'PSModuleManager_custom.json'
function Load-CustomModules {
    if (Test-Path $Global:CustomModulesFile) {
        try {
            $raw = Get-Content $Global:CustomModulesFile -Raw -Encoding UTF8
            $arr = $raw | ConvertFrom-Json
            foreach ($m in $arr) {
                if ($m.Name -and -not ($Global:Catalog | Where-Object { $_.Name -eq $m.Name })) {
                    $entry = [PSCustomObject]@{Name=$m.Name; Cat=$m.Cat; Desc=$m.Desc; Custom=$true}
                    $Global:Catalog += $entry
                    $r = [PSMod.Row]::new()
                    $r.Name=$m.Name; $r.Cat=$m.Cat; $r.Desc=$m.Desc; $r.St='Pending'
                    $Global:AllRows.Add($r)
                }
            }
        } catch {}
    }
}
function Save-CustomModules {
    $custom = $Global:Catalog | Where-Object { $_.Custom -eq $true }
    $custom | Select-Object Name,Cat,Desc | ConvertTo-Json -Depth 2 |
        Set-Content $customModsFile -Encoding UTF8
}
Load-CustomModules

# The GUI scriptblock - everything inside runs on the STA thread
$guiScript = {
    param($engines, $allRows, $catalog, $logFile, $logQ, $scanQ, $batchQ, $termQ, $cancelRef, $cancelBRef, $customModsFile)
    $script:LogFilePath = $logFile  # mutable reference for Set Log Path

    Add-Type -AssemblyName PresentationFramework    -EA SilentlyContinue
    Add-Type -AssemblyName PresentationCore          -EA SilentlyContinue
    Add-Type -AssemblyName WindowsBase               -EA SilentlyContinue
    Add-Type -AssemblyName Microsoft.VisualBasic     -EA SilentlyContinue

    # - UI log helper ----------------
    $logLines = [System.Collections.Generic.List[string]]::new()
    function UiLog {
        param([string]$Msg, [string]$Lvl = 'INFO')
        $e = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')][$Lvl] $Msg"
        $logLines.Add($e)
        $logQ.Enqueue($e)
        try { Add-Content -Path $logFile -Value $e -Encoding UTF8 } catch {}
    }

    # - XAML -------------------
    [xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="PowerShell Module Manager v7.0"
    Width="1300" Height="820"
    MinWidth="900" MinHeight="600"
    WindowStartupLocation="CenterScreen"
    Background="#0F0F1A"
    FontFamily="Segoe UI" FontSize="13">
  <Window.Resources>
    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="#0A0A18"/>
      <Setter Property="Foreground" Value="#C0C0E0"/>
      <Setter Property="BorderBrush" Value="#20204A"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="RowBackground" Value="#0E0E22"/>
      <Setter Property="AlternatingRowBackground" Value="#131328"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#18183A"/>
      <Setter Property="VerticalGridLinesBrush" Value="#18183A"/>
      <Setter Property="SelectionMode" Value="Extended"/>
      <Setter Property="SelectionUnit" Value="FullRow"/>
      <Setter Property="CanUserAddRows" Value="False"/>
      <Setter Property="CanUserDeleteRows" Value="False"/>
      <Setter Property="AutoGenerateColumns" Value="False"/>
      <Setter Property="IsReadOnly" Value="False"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#07071A"/>
      <Setter Property="Foreground" Value="#5A88FF"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="8,5"/>
      <Setter Property="BorderBrush" Value="#1A1A40"/>
      <Setter Property="BorderThickness" Value="0,0,1,2"/>
    </Style>
    <Style TargetType="DataGridRow">
      <Setter Property="Height" Value="26"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#162070"/>
        </Trigger>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#14143A"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="Padding" Value="4,2"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Foreground" Value="#B0B0D8"/>
    </Style>
    <Style x:Key="B0" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="12,5"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Background" Value="#1A3CC0"/>
      <Setter Property="FontSize" Value="12"/>
      <Style.Triggers>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Background" Value="#111120"/>
          <Setter Property="Foreground" Value="#2A2A42"/>
          <Setter Property="Cursor" Value="Arrow"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="BG" TargetType="Button" BasedOn="{StaticResource B0}">
      <Setter Property="Background" Value="#145030"/>
    </Style>
    <Style x:Key="BO" TargetType="Button" BasedOn="{StaticResource B0}">
      <Setter Property="Background" Value="#602800"/>
    </Style>
    <Style x:Key="BR" TargetType="Button" BasedOn="{StaticResource B0}">
      <Setter Property="Background" Value="#601010"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Background" Value="#12122A"/>
      <Setter Property="Foreground" Value="#C0C0E0"/>
      <Setter Property="BorderBrush" Value="#28286A"/>
      <Setter Property="Padding" Value="5,3"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#0A0A1E"/>
      <Setter Property="Foreground" Value="#C0C0E0"/>
      <Setter Property="BorderBrush" Value="#28286A"/>
      <Setter Property="Padding" Value="5,3"/>
      <Setter Property="CaretBrush" Value="#5A88FF"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#7070A0"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Background" Value="#0A0A1A"/>
      <Setter Property="Foreground" Value="#404070"/>
      <Setter Property="Padding" Value="12,5"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="12"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#14142E"/>
          <Setter Property="Foreground" Value="#5A88FF"/>
        </Trigger>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Foreground" Value="#8898CC"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="ScrollBar">
      <Setter Property="Background" Value="#10102A"/>
    </Style>
  <!-- Toggle Switch Style -->
  <Style x:Key="ToggleSwitch" TargetType="CheckBox" xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Focusable" Value="False"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="CheckBox">
          <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
            <Border x:Name="Track" Width="46" Height="22" CornerRadius="11"
                    Background="#16162A" BorderBrush="#28284A" BorderThickness="1"
                    VerticalAlignment="Center">
              <Grid>
                <TextBlock x:Name="LblOff" Text="OFF" FontSize="8" FontFamily="Consolas"
                           Foreground="#404060" HorizontalAlignment="Right" Margin="0,0,5,0"
                           VerticalAlignment="Center"/>
                <TextBlock x:Name="LblOn"  Text="ON"  FontSize="8" FontFamily="Consolas"
                           Foreground="#28A048" HorizontalAlignment="Left"  Margin="5,0,0,0"
                           VerticalAlignment="Center" Visibility="Collapsed"/>
                <Ellipse x:Name="Thumb" Width="16" Height="16" Fill="#404060"
                         HorizontalAlignment="Left" VerticalAlignment="Center" Margin="2,0,0,0">
                  <Ellipse.Effect>
                    <DropShadowEffect Color="Black" BlurRadius="3" ShadowDepth="1" Opacity="0.5"/>
                  </Ellipse.Effect>
                </Ellipse>
              </Grid>
            </Border>
            <ContentPresenter x:Name="Lbl" Margin="8,0,0,0" VerticalAlignment="Center"/>
          </StackPanel>
          <ControlTemplate.Triggers>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="Track"  Property="Background"  Value="#08280E"/>
              <Setter TargetName="Track"  Property="BorderBrush" Value="#1A5025"/>
              <Setter TargetName="Thumb"  Property="Fill"        Value="#30B050"/>
              <Setter TargetName="Thumb"  Property="HorizontalAlignment" Value="Right"/>
              <Setter TargetName="Thumb"  Property="Margin"      Value="0,0,2,0"/>
              <Setter TargetName="LblOn"  Property="Visibility"  Value="Visible"/>
              <Setter TargetName="LblOff" Property="Visibility"  Value="Collapsed"/>
            </Trigger>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Track"  Property="BorderBrush" Value="#4A4A7A"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False">
              <Setter TargetName="Track"  Property="Opacity"     Value="0.4"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="46"/>
      <RowDefinition Height="36"/>
      <RowDefinition Height="50"/>
      <RowDefinition Height="*" MinHeight="100"/>
      <RowDefinition Height="5"/>
      <RowDefinition Height="220" MinHeight="80"/>
      <RowDefinition Height="24"/>
    </Grid.RowDefinitions>

    <!-- TITLE -->
    <Border Grid.Row="0" Background="#07071A" BorderBrush="#181838" BorderThickness="0,0,0,1">
      <Grid Margin="14,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="PS" FontSize="20" FontWeight="Black" Foreground="#5A88FF"/>
          <TextBlock Text=" Module Manager" FontSize="17" FontWeight="Bold" Foreground="#D0D0FF"/>
          <TextBlock Text=" v7.0" FontSize="10" Foreground="#282858" VerticalAlignment="Bottom" Margin="3,0,0,2"/>
        </StackPanel>
        <TextBlock Grid.Column="2" Name="lblPs7"
                   Foreground="#F0A030" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </Grid>
    </Border>


    <!-- RIBBON: Move/Open/Disable/Export/Theme/About -->
    <Border Grid.Row="1" Background="#09091F" BorderBrush="#181838" BorderThickness="0,0,0,1" Padding="8,3">
      <DockPanel>
        <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
          <Button Name="btnSetPath"     Content="Set PSModulePath"  Style="{StaticResource B0}" Margin="0,0,4,0" Padding="8,3" FontSize="11" ToolTip="Set a custom directory as default module install location"/>
          <Button Name="btnMoveModule"  Content="Move Module"         Style="{StaticResource B0}" Margin="0,0,4,0" Padding="8,3" FontSize="11"/>
          <Button Name="btnOpenFolder"  Content="Open Module Folder"  Style="{StaticResource B0}" Margin="0,0,4,0" Padding="8,3" FontSize="11"/>
          <Button Name="btnDisable"     Content="Disable Module"      Style="{StaticResource BO}" Margin="0,0,4,0" Padding="8,3" FontSize="11"/>
          <Button Name="btnEnable"      Content="Enable Module"       Style="{StaticResource BG}" Margin="0,0,4,0" Padding="8,3" FontSize="11"/>
          <Button Name="btnExport"      Content="Export Module List"  Style="{StaticResource B0}" Margin="0,0,4,0" Padding="8,3" FontSize="11"/>
          <Button Name="btnCleanAll"   Content="Clean ALL Modules"   Style="{StaticResource BR}" Margin="0,0,4,0"  Padding="8,3" FontSize="11"/>
          <Button Name="btnCleanOld"   Content="Clean Old Versions"  Style="{StaticResource BO}" Margin="0,0,12,0" Padding="8,3" FontSize="11" ToolTip="Find and remove old duplicate module versions keeping only the latest"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" DockPanel.Dock="Right" HorizontalAlignment="Right">
          <TextBlock Text="Theme:" Foreground="#404070" VerticalAlignment="Center" Margin="0,0,5,0" FontSize="11"/>
          <ComboBox Name="cmbTheme" Width="80" FontSize="11" VerticalAlignment="Center" Margin="0,0,10,0">
            <ComboBoxItem Content="Dark" IsSelected="True"/>
            <ComboBoxItem Content="Light"/>
            <ComboBoxItem Content="Auto"/>
          </ComboBox>
          <Button Name="btnInfo"  Content="Info"  Style="{StaticResource B0}" Padding="8,3" FontSize="11" Margin="0,0,6,0" ToolTip="Feature guide and help"/>
          <Button Name="btnAbout" Content="About" Style="{StaticResource B0}" Padding="8,3" FontSize="11"/>
        </StackPanel>
      </DockPanel>
    </Border>
    <!-- TOOLBAR -->
    <Border Grid.Row="2" Background="#0C0C1E" BorderBrush="#181838" BorderThickness="0,0,0,1" Padding="12,6">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="190"/>
          <ColumnDefinition Width="10"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="150"/>
          <ColumnDefinition Width="10"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="170"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="6"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="10"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0"  Text="Engine:" Foreground="#404070" VerticalAlignment="Center" Margin="0,0,6,0"/>
        <ComboBox  Grid.Column="1"  Name="cmbEng" VerticalAlignment="Center"/>
        <TextBlock Grid.Column="3"  Text="Scope:"  Foreground="#404070" VerticalAlignment="Center" Margin="0,0,6,0"/>
        <ComboBox  Grid.Column="4"  Name="cmbScope" VerticalAlignment="Center">
          <ComboBoxItem Content="AllUsers  (requires Admin)" IsSelected="True"/>
          <ComboBoxItem Content="CurrentUser"/>
        </ComboBox>
        <TextBlock Grid.Column="6"  Text="Filter:" Foreground="#404070" VerticalAlignment="Center" Margin="0,0,6,0"/>
        <TextBox   Grid.Column="7"  Name="txtFilter" VerticalAlignment="Center"/>
        <CheckBox  Grid.Column="9"  Name="chkInst" Content="Installed only"/>
        <CheckBox  Grid.Column="11" Name="chkUpd"  Content="Updates only" Foreground="#F0A030"/>
        <Button    Grid.Column="13" Name="btnScan"     Content="Refresh Status"
                   Style="{StaticResource B0}" Width="125" Margin="0,0,4,0"/>
        <Button    Grid.Column="14" Name="btnStopScan"  Content="Stop"
                   Style="{StaticResource BR}" Width="55" IsEnabled="False"/>
      </Grid>
    </Border>

    <!-- MAIN TABS -->
    <TabControl Grid.Row="3" Name="tabC" Background="#0C0C1E" BorderBrush="#181838">

      <!-- ===== MODULE CATALOG ===== -->
      <TabItem Header="Module Catalog">
        <Grid Background="#0C0C1E">
          <Grid.RowDefinitions>
            <RowDefinition Height="34"/>
            <RowDefinition Height="22"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <!-- Category chips row -->
          <ScrollViewer Grid.Row="0" HorizontalScrollBarVisibility="Auto"
                        VerticalScrollBarVisibility="Disabled"
                        Background="#0A0A1C">
            <StackPanel Name="pnlCats" Orientation="Horizontal"
                        Margin="6,4" VerticalAlignment="Center"/>
          </ScrollViewer>

          <!-- Scope Legend -->
          <Border Grid.Row="1" Background="#060616" Padding="10,0"
                  BorderBrush="#181838" BorderThickness="0,0,0,1">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <TextBlock Text="Scope: " Foreground="#252550" FontSize="10" VerticalAlignment="Center"/>
              <TextBlock Text="AllUsers" Foreground="#5080E0" FontSize="10" FontFamily="Consolas" Margin="4,0,8,0"/>
              <TextBlock Text="CurrentUser" Foreground="#50C080" FontSize="10" FontFamily="Consolas" Margin="0,0,8,0"/>
              <TextBlock Text="System/WinPS/PS7" Foreground="#606090" FontSize="10" FontFamily="Consolas" Margin="0,0,8,0"/>
              <TextBlock Text="Unknown" Foreground="#906030" FontSize="10" FontFamily="Consolas" Margin="0,0,8,0"/>
              <TextBlock Text="| empty = not installed in that engine" Foreground="#202045" FontSize="10" Margin="0,0,0,0"/>
            </StackPanel>
          </Border>

          <!-- DataGrid -->
          <DataGrid Grid.Row="2" Name="dg" Margin="6,0,6,6">
            <DataGrid.Columns>
              <DataGridTemplateColumn Width="30" CanUserResize="False" CanUserSort="False" Header=" ">
                <DataGridTemplateColumn.CellTemplate>
                  <DataTemplate>
                    <CheckBox IsChecked="{Binding Sel, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                              HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTextColumn Header="Module Name" Binding="{Binding Name}" Width="195" FontWeight="SemiBold"/>
              <DataGridTextColumn Header="Category"    Binding="{Binding Cat}"  Width="82"/>
              <DataGridTextColumn Header="Description" Binding="{Binding Desc}" Width="*"/>
              <DataGridTextColumn Header="PS5 Ver"   Binding="{Binding V5}"  Width="72">
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock">
                    <Setter Property="Foreground" Value="#5080E0"/>
                    <Setter Property="FontFamily" Value="Consolas"/>
                  </Style>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Header="PS5 Scope" Binding="{Binding S5}"  Width="105">
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock">
                    <Setter Property="Foreground" Value="#505090"/>
                    <Setter Property="FontFamily" Value="Consolas"/>
                    <Setter Property="FontSize"   Value="11"/>
                    <Style.Triggers>
                      <DataTrigger Binding="{Binding S5}" Value="AllUsers">
                        <Setter Property="Foreground" Value="#5080E0"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding S5}" Value="CurrentUser">
                        <Setter Property="Foreground" Value="#50C080"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding S5}" Value="System">
                        <Setter Property="Foreground" Value="#606090"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding S5}" Value="WinPS-System">
                        <Setter Property="Foreground" Value="#505085"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding S5}" Value="Unknown">
                        <Setter Property="Foreground" Value="#906030"/>
                      </DataTrigger>
                    </Style.Triggers>
                  </Style>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Header="PS7 Ver"   Binding="{Binding V7}"  Width="72">
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock">
                    <Setter Property="Foreground" Value="#5080E0"/>
                    <Setter Property="FontFamily" Value="Consolas"/>
                  </Style>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Header="PS7 Scope" Binding="{Binding S7}"  Width="105">
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock">
                    <Setter Property="Foreground" Value="#505090"/>
                    <Setter Property="FontFamily" Value="Consolas"/>
                    <Setter Property="FontSize"   Value="11"/>
                    <Style.Triggers>
                      <DataTrigger Binding="{Binding S7}" Value="AllUsers">
                        <Setter Property="Foreground" Value="#5080E0"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding S7}" Value="CurrentUser">
                        <Setter Property="Foreground" Value="#50C080"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding S7}" Value="System">
                        <Setter Property="Foreground" Value="#606090"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding S7}" Value="WinPS-System">
                        <Setter Property="Foreground" Value="#505085"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding S7}" Value="Unknown">
                        <Setter Property="Foreground" Value="#906030"/>
                      </DataTrigger>
                    </Style.Triggers>
                  </Style>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Header="Gallery"     Binding="{Binding GV}"   Width="85"/>
              <DataGridTextColumn Header="Status"      Binding="{Binding St}"   Width="110"/>
            </DataGrid.Columns>
            <DataGrid.RowStyle>
              <Style TargetType="DataGridRow">
                <Setter Property="Height" Value="26"/>
                <Setter Property="Cursor" Value="Hand"/>
                <Style.Triggers>
                  <DataTrigger Binding="{Binding Up}" Value="True">
                    <Setter Property="Foreground" Value="#F0C040"/>
                  </DataTrigger>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#162070"/>
                  </Trigger>
                  <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#14143A"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </DataGrid.RowStyle>
          </DataGrid>
        </Grid>
      </TabItem>

      <!-- ===== LOG VIEWER ===== -->
      <TabItem Header="Log Viewer">
        <Grid Background="#050512" Margin="6">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,5">
            <Button Name="btnClearLog"  Content="Clear Log"       Style="{StaticResource BR}" Margin="0,0,5,0"/>
            <Button Name="btnOpenLog"   Content="Open in Notepad" Style="{StaticResource B0}" Margin="0,0,5,0"/>
            <Button Name="btnReloadLog" Content="Reload"          Style="{StaticResource B0}" Margin="0,0,5,0"/>
            <Button Name="btnSetLogPath" Content="Set Log Path" Style="{StaticResource B0}" Margin="0,0,10,0" ToolTip="Choose a custom folder for the log file"/>
            <TextBlock Name="lblLogPath" Foreground="#202050" VerticalAlignment="Center" FontSize="11"/>
          </StackPanel>
          <ScrollViewer Grid.Row="1" Name="svLog" VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Auto" Background="#030310">
            <TextBlock Name="txtLog" FontFamily="Consolas" FontSize="11"
                       Foreground="#6060A0" Padding="6" TextWrapping="NoWrap"/>
          </ScrollViewer>
        </Grid>
      </TabItem>

      <!-- ===== ENGINE INFO ===== -->
      <TabItem Header="Engine Info">
        <Grid Background="#050512">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <!-- Top action bar -->
          <Border Grid.Row="0" Background="#07071C" Padding="8,6" BorderBrush="#181838" BorderThickness="0,0,0,1">
            <StackPanel Orientation="Horizontal">
              <Button Name="btnCheckEngVer"   Content="Check for Updates"      Style="{StaticResource B0}" Margin="0,0,6,0" Padding="10,4" FontSize="11"/>
              <Button Name="btnInstPS7"       Content="Install / Update PS7"   Style="{StaticResource BG}" Margin="0,0,6,0" Padding="10,4" FontSize="11"/>
              <Button Name="btnInstPS5"       Content="Install WMF 5.1"        Style="{StaticResource B0}" Margin="0,0,16,0" Padding="10,4" FontSize="11"/>
              <Rectangle Width="1" Fill="#181848" Margin="0,0,16,0"/>
              <Button Name="btnCheckWinget"   Content="Check winget"           Style="{StaticResource B0}" Margin="0,0,6,0" Padding="10,4" FontSize="11"/>
              <Button Name="btnInstWinget"    Content="Install / Update winget" Style="{StaticResource BG}" Margin="0,0,6,0" Padding="10,4" FontSize="11"/>
            </StackPanel>
          </Border>

          <!-- Engine cards -->
          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="6">
            <StackPanel Name="pnlEng" Margin="6"/>
          </ScrollViewer>

          <!-- Engine action log -->
          <Border Grid.Row="2" Background="#030310" Height="100" BorderBrush="#181838" BorderThickness="0,1,0,0">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="18"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" Background="#07071C" Padding="8,1">
                <TextBlock Text="Engine Action Log" Foreground="#202050" FontSize="10" FontWeight="SemiBold"/>
              </Border>
              <ScrollViewer Grid.Row="1" Name="svEngLog" VerticalScrollBarVisibility="Auto">
                <TextBlock Name="txtEngLog" FontFamily="Consolas" FontSize="10"
                           Foreground="#4A8A4A" Padding="6,2" TextWrapping="Wrap"/>
              </ScrollViewer>
            </Grid>
          </Border>
        </Grid>
      </TabItem>

      <!-- ===== MODULE PATHS ===== -->
      <TabItem Header="Module Paths">
        <Grid Background="#050512">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <!-- Top: PSModulePath env var -->
          <Border Grid.Row="0" Background="#07071C" Padding="10,6" BorderBrush="#181838" BorderThickness="0,0,0,1">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="$env:PSModulePath  " Foreground="#303070" FontFamily="Consolas" FontSize="11" VerticalAlignment="Center"/>
              <Button Name="btnRefreshPaths" Content="Refresh" Style="{StaticResource B0}" Padding="8,2" FontSize="10"/>
              <Button Name="btnOpenPathExplorer" Content="Open All in Explorer" Style="{StaticResource B0}" Padding="8,2" FontSize="10" Margin="6,0,0,0"/>
              <Button Name="btnAddToPath"  Content="Add Folder to PSModulePath" Style="{StaticResource BG}" Padding="8,2" FontSize="10" Margin="16,0,0,0"/>
              <Button Name="btnRemFromPath" Content="Remove Selected" Style="{StaticResource BR}" Padding="8,2" FontSize="10" Margin="6,0,0,0"/>
            </StackPanel>
          </Border>

          <!-- Main: DataGrid of paths -->
          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="6">
            <StackPanel Name="pnlPaths" Margin="4"/>
          </ScrollViewer>

          <!-- Bottom: info -->
          <Border Grid.Row="2" Background="#07071C" Padding="10,4" BorderBrush="#181838" BorderThickness="0,1,0,0">
            <TextBlock Name="lblPathInfo" Foreground="#282860" FontSize="11"
                       Text="Click 'Refresh' to load all PS module paths and installed modules per location."/>
          </Border>
        </Grid>
      </TabItem>

      <!-- ===== REPOSITORIES TAB ===== -->
      <TabItem Header="Repositories">
        <Grid Background="#050512">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <!-- Top toolbar -->
          <Border Grid.Row="0" Background="#07071C" Padding="8,6" BorderBrush="#181838" BorderThickness="0,0,0,1">
            <DockPanel>
              <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
                <Button Name="btnRepoRefresh"   Content="Refresh"              Style="{StaticResource B0}" Padding="8,3" FontSize="11" Margin="0,0,4,0"/>
                <Button Name="btnRepoRegister"  Content="Register Repo"        Style="{StaticResource BG}" Padding="8,3" FontSize="11" Margin="0,0,4,0" ToolTip="Register a new PSRepository (PSGallery-compatible NuGet feed)"/>
                <Button Name="btnRepoUnreg"     Content="Unregister Selected"  Style="{StaticResource BR}" Padding="8,3" FontSize="11" Margin="0,0,16,0" ToolTip="Unregister selected repository (PSGallery cannot be removed)"/>
                <Rectangle Width="1" Fill="#181848" Margin="0,0,16,0" VerticalAlignment="Stretch"/>
                <Button Name="btnAddCatalogMod" Content="+ Add Module to Catalog" Style="{StaticResource BG}" Padding="8,3" FontSize="11" Margin="0,0,4,0" ToolTip="Search a module in a repo and add it to the scan catalog"/>
                <Button Name="btnRemCatalogMod" Content="- Remove from Catalog"   Style="{StaticResource BR}" Padding="8,3" FontSize="11" Margin="0,0,0,0" ToolTip="Remove a CUSTOM module from catalog (built-in cannot be removed)"/>
              </StackPanel>
              <StackPanel Orientation="Horizontal" DockPanel.Dock="Right" HorizontalAlignment="Right" Margin="0,0,4,0">
                <TextBlock Text="Find module: " Foreground="#303060" FontSize="11" VerticalAlignment="Center" Margin="0,0,6,0"/>
                <TextBox Name="txtRepoSearch" Width="180" FontSize="11" Padding="4,2" Background="#0A0A1E" Foreground="#C0C0E0" BorderBrush="#28286A" VerticalAlignment="Center"/>
                <Button Name="btnRepoFind"    Content="Search Gallery"     Style="{StaticResource B0}" Padding="8,3" FontSize="11" Margin="6,0,0,0"/>
                <Button Name="btnBrowseRepo" Content="Browse Repo Modules" Style="{StaticResource B0}" Padding="8,3" FontSize="11" Margin="6,0,0,0" ToolTip="List all modules available in registered repositories and add them to the catalog"/>
              </StackPanel>
            </DockPanel>
          </Border>

          <!-- Split: Left = repositories  Right = custom catalog -->
          <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="5"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- LEFT panel: registered repos -->
            <Grid Grid.Column="0">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" Background="#060616" Padding="10,5" BorderBrush="#181838" BorderThickness="0,0,0,1">
                <TextBlock Text="REGISTERED REPOSITORIES" Foreground="#404080" FontSize="10" FontWeight="Bold" FontFamily="Consolas"/>
              </Border>
              <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Background="#050512">
                <StackPanel Name="pnlRepos" Margin="6"/>
              </ScrollViewer>
            </Grid>

            <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch"
                          Background="#181838" ResizeBehavior="PreviousAndNext"/>

            <!-- RIGHT panel: custom catalog modules -->
            <Grid Grid.Column="2">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" Background="#060616" Padding="8,5" BorderBrush="#181838" BorderThickness="0,0,0,1">
                <DockPanel>
                  <Button Name="btnInstallFromRepo" Content="Install Selected" Visibility="Collapsed"
                          DockPanel.Dock="Right" Padding="8,2" FontSize="10"
                          Background="#0A3020" Foreground="#30C060" BorderBrush="#1A5030" Cursor="Hand"
                          ToolTip="Install checked modules from repository"/>
                  <Button Name="btnClearBrowse" Content="&#x2715; Clear" Visibility="Collapsed"
                          DockPanel.Dock="Right" Padding="6,2" FontSize="10" Margin="0,0,4,0"
                          Background="#1A0808" Foreground="#804040" BorderBrush="#3A1818" Cursor="Hand"/>
                  <TextBlock Name="lblRightPanelTitle" Text="CUSTOM CATALOG MODULES" Foreground="#404080" FontSize="10" FontWeight="Bold" FontFamily="Consolas" VerticalAlignment="Center"/>
                  <TextBlock Name="lblCustomCount" Text="" Foreground="#2A2A60" FontSize="10" FontFamily="Consolas" Margin="8,0,0,0" VerticalAlignment="Center"/>
                </DockPanel>
              </Border>
              <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Background="#050512">
                <StackPanel Name="pnlCustomMods" Margin="6"/>
              </ScrollViewer>
            </Grid>
          </Grid>

          <!-- Bottom status bar -->
          <Border Grid.Row="2" Background="#07071C" Padding="10,5" BorderBrush="#181838" BorderThickness="0,1,0,0">
            <TextBlock Name="lblRepoStatus" Foreground="#282860" FontSize="11"
                       Text="Click Refresh to load registered repositories. Use '+ Add Module to Catalog' to extend the scan list."/>
          </Border>
        </Grid>
      </TabItem>

    </TabControl>

    <!-- DRAG SPLITTER - resize between module list and log panel -->
    <GridSplitter Grid.Row="4" Height="5" HorizontalAlignment="Stretch"
                  VerticalAlignment="Stretch" Background="#252555"
                  ResizeBehavior="PreviousAndNext" ResizeDirection="Rows"
                  Cursor="SizeNS" ToolTip="Drag up/down to resize"/>

    <!-- BOTTOM SECTION (Grid.Row=5): Action bar + Log panel -->
    <Grid Grid.Row="5">
      <Grid.RowDefinitions>
        <RowDefinition Height="44"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <!-- ACTION BAR (Row 0 of bottom Grid) -->
      <Border Grid.Row="0" Background="#07071A" BorderBrush="#181838" BorderThickness="0,1,0,1" Padding="8,5">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Button Grid.Column="0" Name="btnAll"    Content="Select All"       Style="{StaticResource B0}" Margin="0,0,4,0" Width="88"/>
          <Button Grid.Column="1" Name="btnNone"   Content="Clear"            Style="{StaticResource B0}" Margin="0,0,4,0" Width="58"/>
          <Button Grid.Column="2" Name="btnSelUpd" Content="Select Updatable" Style="{StaticResource BO}" Margin="0,0,4,0" Width="128"/>
          <Button Grid.Column="3" Name="btnSelMis" Content="Select Missing"   Style="{StaticResource BO}" Margin="0,0,4,0" Width="118"/>
          <TextBlock Grid.Column="4" Name="lblSel" Foreground="#282860"
                     VerticalAlignment="Center" Margin="8,0" Text="0 selected" FontSize="12"/>
          <Button Grid.Column="5" Name="btnInst"   Content="Install / Update Selected"
                  Style="{StaticResource BG}" Width="200" Margin="0,0,4,0" IsEnabled="False"/>
          <Button Grid.Column="6" Name="btnUpdAll" Content="Update ALL Installed"
                  Style="{StaticResource BO}" Width="160" Margin="0,0,4,0"/>
          <Button Grid.Column="7" Name="btnRemove"  Content="Remove Selected"
                  Style="{StaticResource BR}" Width="130" Margin="0,0,4,0" IsEnabled="False"/>
          <Button Grid.Column="8" Name="btnCancel" Content="Cancel"
                  Style="{StaticResource BR}" Width="68" IsEnabled="False"/>
        </Grid>
      </Border>

      <!-- LOG PANEL (Row 1 of bottom Grid): Operation Log | Terminal Output -->
      <Border Grid.Row="1" Background="#030310" BorderBrush="#121230" BorderThickness="0,0,0,1">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" MinWidth="200"/>
            <ColumnDefinition Width="5"/>
            <ColumnDefinition Width="*" MinWidth="200"/>
          </Grid.ColumnDefinitions>

          <!-- LEFT: Operation Log -->
          <Grid Grid.Column="0">
            <Grid.RowDefinitions>
              <RowDefinition Height="20"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Border Grid.Row="0" Background="#07071C" Padding="8,2">
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="Operation Log" Foreground="#202050"
                           FontWeight="SemiBold" FontSize="11"/>
                <ProgressBar Name="pbar" Width="150" Height="8" Margin="10,0,0,0"
                             VerticalAlignment="Center" Background="#0E0E30"
                             Foreground="#2040C0" Minimum="0" Maximum="100"
                             Value="0" Visibility="Collapsed"/>
                <TextBlock Name="lblProg" Foreground="#303070" FontSize="11"
                           Margin="6,0,0,0" VerticalAlignment="Center"/>
              </StackPanel>
            </Border>
            <ScrollViewer Grid.Row="1" Name="svMini" VerticalScrollBarVisibility="Auto">
              <TextBlock Name="txtMini" FontFamily="Consolas" FontSize="11"
                         Foreground="#5060A0" Padding="6,3" TextWrapping="Wrap"/>
            </ScrollViewer>
          </Grid>

          <!-- SPLITTER between log and terminal -->
          <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch"
                        Background="#181838" ResizeBehavior="PreviousAndNext"/>

          <!-- RIGHT: Terminal / Verbose Output -->
          <Grid Grid.Column="2">
            <Grid.RowDefinitions>
              <RowDefinition Height="20"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Border Grid.Row="0" Background="#07071C" Padding="8,2">
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="Terminal Output" Foreground="#304070"
                           FontWeight="SemiBold" FontSize="11"/>
                <Button Name="btnClearTerm" Content="Clear" FontSize="10"
                        Padding="4,0" Margin="8,1,0,1" Height="16"
                        Background="#151530" Foreground="#5060A0" BorderThickness="1"
                        BorderBrush="#202050"/>
              </StackPanel>
            </Border>
            <ScrollViewer Grid.Row="1" Name="svTerm" VerticalScrollBarVisibility="Auto"
                          HorizontalScrollBarVisibility="Auto" Background="#020210">
              <TextBlock Name="txtTerm" FontFamily="Consolas" FontSize="11"
                         Foreground="#4A8A4A" Padding="6,3" TextWrapping="NoWrap"/>
            </ScrollViewer>
          </Grid>
        </Grid>
      </Border>
    </Grid><!-- end bottom section -->


    <!-- STATUS BAR -->
    <Border Grid.Row="6" Background="#040410" Padding="10,0">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Name="lblStatus" Grid.Column="0"
                   Foreground="#222252" VerticalAlignment="Center" FontSize="11"/>
        <TextBlock Name="lblExe"    Grid.Column="1"
                   Foreground="#182080" VerticalAlignment="Center" FontSize="11"/>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

    # - Parse XAML -----------------
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $win    = [System.Windows.Markup.XamlReader]::Load($reader)
    $script:MainWindow = $win  # Store for dialog Owner access

    # - Get controls (all by name) ------------
    function G([string]$n){ $win.FindName($n) }
    $cmbEng     = G 'cmbEng';    $cmbScope   = G 'cmbScope'
    $txtFilter  = G 'txtFilter'; $chkInst    = G 'chkInst'
    $chkUpd     = G 'chkUpd';    $btnScan    = G 'btnScan'
    $pnlCats    = G 'pnlCats';   $dg         = G 'dg'
    $btnAll     = G 'btnAll';    $btnNone    = G 'btnNone'
    $btnSelUpd  = G 'btnSelUpd'; $btnSelMis  = G 'btnSelMis'
    $btnInst    = G 'btnInst';   $btnUpdAll  = G 'btnUpdAll'
    $btnRemove  = G 'btnRemove'; $btnCancel  = G 'btnCancel'
    $lblSel     = G 'lblSel'
    $btnStopScan   = G 'btnStopScan'
    $btnSetPath    = G 'btnSetPath'
    $btnMoveModule = G 'btnMoveModule'; $btnOpenFolder = G 'btnOpenFolder'
    $btnRefreshPaths = G 'btnRefreshPaths'; $btnOpenPathExplorer = G 'btnOpenPathExplorer'
    $btnAddToPath    = G 'btnAddToPath';   $btnRemFromPath = G 'btnRemFromPath'
    $pnlPaths = G 'pnlPaths'; $lblPathInfo = G 'lblPathInfo'
    $btnDisable    = G 'btnDisable';    $btnEnable     = G 'btnEnable'
    $btnExport     = G 'btnExport';     $btnCleanAll   = G 'btnCleanAll'
    $btnCleanOld   = G 'btnCleanOld'
    $pnlRepos          = G 'pnlRepos';       $pnlCustomMods     = G 'pnlCustomMods'
    $lblRepoStatus     = G 'lblRepoStatus';  $lblCustomCount    = G 'lblCustomCount'
    $lblRightPanelTitle= G 'lblRightPanelTitle'
    $btnInstallFromRepo= G 'btnInstallFromRepo'
    $btnClearBrowse    = G 'btnClearBrowse'
    $script:BrowseMode = $false
    $script:BrowseCheckboxes = [System.Collections.Generic.List[object]]::new()
    $btnRepoRefresh= G 'btnRepoRefresh';$btnRepoRegister= G 'btnRepoRegister'
    $btnRepoUnreg  = G 'btnRepoUnreg';  $btnAddCatalogMod=G 'btnAddCatalogMod'
    $btnRemCatalogMod=G 'btnRemCatalogMod'
    $txtRepoSearch = G 'txtRepoSearch'; $btnRepoFind   = G 'btnRepoFind'
    $cmbTheme      = G 'cmbTheme'
    $btnInfo       = G 'btnInfo'
    $btnAbout      = G 'btnAbout'
    $lblStatus  = G 'lblStatus'; $lblExe     = G 'lblExe'
    $txtMini    = G 'txtMini';   $svMini     = G 'svMini'
    $txtLog     = G 'txtLog';    $svLog      = G 'svLog'
    $lblLogPath = G 'lblLogPath';$btnClearLog= G 'btnClearLog'
    $btnOpenLog = G 'btnOpenLog';$btnReloadLog=G 'btnReloadLog'
    $svTerm     = G 'svTerm';    $txtTerm     = G 'txtTerm'
    $btnClearTerm = G 'btnClearTerm'
    $pbar       = G 'pbar';      $lblProg    = G 'lblProg'
    $pnlEng         = G 'pnlEng';     $lblPs7         = G 'lblPs7'
    $svEngLog       = G 'svEngLog';   $txtEngLog      = G 'txtEngLog'
    $btnCheckEngVer = G 'btnCheckEngVer'; $btnInstPS7 = G 'btnInstPS7'
    $btnInstPS5     = G 'btnInstPS5'; $btnCheckWinget = G 'btnCheckWinget'
    $btnInstWinget  = G 'btnInstWinget'

    $lblLogPath.Text = "Log: $logFile"

    # - Mini log (always on UI thread already) --------
    function ML([string]$t, [string]$lvl = 'INFO') {
        $ts = Get-Date -Format 'HH:mm:ss'
        $line = "[$ts] $t"
        $script:MiniLines.Add($line) | Out-Null
        while ($script:MiniLines.Count -gt 2000) { $script:MiniLines.RemoveAt(0) }
        $txtMini.Text = ($script:MiniLines -join "`n")
        $svMini.ScrollToEnd()
        UiLog $t $lvl
    }
    $script:MiniLines = [System.Collections.Generic.List[string]]::new()

    function ReloadLogView {
        $txtLog.Text = ($logLines -join "`n")
        $svLog.ScrollToEnd()
    }

    # Terminal output - verbose live output (green on black style)
    $script:TermLines = [System.Collections.Generic.List[string]]::new()
    function TL([string]$Msg, [string]$Prefix = '') {
        $ts   = Get-Date -Format 'HH:mm:ss.fff'
        $line = if ($Prefix) { "[$ts][$Prefix] $Msg" } else { "[$ts] $Msg" }
        $script:TermLines.Add($line)
        while ($script:TermLines.Count -gt 5000) { $script:TermLines.RemoveAt(0) }
        $txtTerm.Text = ($script:TermLines -join "`n")
        $svTerm.ScrollToEnd()
        # Mirror to log file for full verbose in notepad
        try { Add-Content -Path $logFile -Value "[TERM]$line" -Encoding UTF8 } catch {}
    }

    function TLHead([string]$Title) {
        $sep = '=' * 60
        TL $sep
        TL "  $Title"
        TL $sep
    }

    # - Engines into ComboBox -------------
    foreach ($e in $engines) {
        $it = [System.Windows.Controls.ComboBoxItem]::new()
        $it.Content   = "$($e.Label)  [$($e.Ver)]"
        $it.Tag       = $e
        $it.IsEnabled = $e.OK
        $cmbEng.Items.Add($it) | Out-Null
    }
    $cmbEng.SelectedIndex = 0

    $e7 = $engines | Where-Object { $_.Tag -eq 'PS7' -and $_.OK } | Select-Object -First 1
    if ($e7) {
        $lblPs7.Text       = "PS7 detected: $($e7.Ver)"
        $lblPs7.Foreground = '#50D090'
    } else {
        $lblPs7.Text = 'PS7 not installed'
    }

    # - Engine info panel ---------------
    foreach ($e in $engines) {
        $brd = [System.Windows.Controls.Border]::new()
        $brd.Background = if ($e.OK) { '#0C1A3A' } else { '#1C0808' }
        $brd.Margin     = [System.Windows.Thickness]::new(0,0,0,8)
        $brd.Padding    = [System.Windows.Thickness]::new(14,10)
        $brd.Tag        = $e.Tag  # e.g. 'PS5' or 'PS7'

        $sp = [System.Windows.Controls.StackPanel]::new()

        # Status + name header
        $icon = if ($e.OK) { '[ONLINE]' } else { '[OFFLINE]' }
        $t = [System.Windows.Controls.TextBlock]::new()
        $t.Text       = "$icon  $($e.Label)"
        $t.FontSize   = 14; $t.FontWeight = 'Bold'
        $t.Foreground = if ($e.OK) { '#50D090' } else { '#D05050' }
        $t.Margin     = [System.Windows.Thickness]::new(0,0,0,6)
        $sp.Children.Add($t) | Out-Null

        # Properties grid
        $propsGrid = [System.Windows.Controls.Grid]::new()
        $propsGrid.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())
        $propsGrid.ColumnDefinitions.Add([System.Windows.Controls.ColumnDefinition]::new())
        $col1def = [System.Windows.Controls.ColumnDefinition]::new()
        $col1def.Width = [System.Windows.GridLength]::new(200)
        $propsGrid.ColumnDefinitions[0].Width = [System.Windows.GridLength]::new(200)

        $props = [ordered]@{
            'Installed Version' = $e.Ver
            'Executable Path'   = $e.Path
            'Apartment State'   = 'STA (required for WPF)'
        }
        $row = 0
        foreach ($pk in $props.Keys) {
            $rd = [System.Windows.Controls.RowDefinition]::new(); $rd.Height = [System.Windows.GridLength]::Auto
            $propsGrid.RowDefinitions.Add($rd)

            $lbl = [System.Windows.Controls.TextBlock]::new()
            $lbl.Text = "  $pk"; $lbl.Foreground = '#404070'; $lbl.FontSize = 11
            $lbl.Margin = [System.Windows.Thickness]::new(0,1,0,1)
            [System.Windows.Controls.Grid]::SetRow($lbl, $row); [System.Windows.Controls.Grid]::SetColumn($lbl, 0)
            $propsGrid.Children.Add($lbl) | Out-Null

            $val = [System.Windows.Controls.TextBlock]::new()
            $val.Text = $props[$pk]; $val.Foreground = '#A0A0D0'; $val.FontFamily = 'Consolas'; $val.FontSize = 11
            $val.TextWrapping = 'Wrap'; $val.Margin = [System.Windows.Thickness]::new(0,1,0,1)
            [System.Windows.Controls.Grid]::SetRow($val, $row); [System.Windows.Controls.Grid]::SetColumn($val, 1)
            $propsGrid.Children.Add($val) | Out-Null
            $row++
        }

        # Latest version row (populated after Check for Updates)
        $rdLatest = [System.Windows.Controls.RowDefinition]::new(); $rdLatest.Height = [System.Windows.GridLength]::Auto
        $propsGrid.RowDefinitions.Add($rdLatest)
        $lblLatest = [System.Windows.Controls.TextBlock]::new()
        $lblLatest.Text = "  Latest Available"; $lblLatest.Foreground = '#404070'; $lblLatest.FontSize = 11
        $lblLatest.Margin = [System.Windows.Thickness]::new(0,1,0,1)
        [System.Windows.Controls.Grid]::SetRow($lblLatest, $row); [System.Windows.Controls.Grid]::SetColumn($lblLatest, 0)
        $propsGrid.Children.Add($lblLatest) | Out-Null

        $valLatest = [System.Windows.Controls.TextBlock]::new()
        $valLatest.Text = '(click Check for Updates)'; $valLatest.Foreground = '#303060'
        $valLatest.FontFamily = 'Consolas'; $valLatest.FontSize = 11; $valLatest.Tag = 'latest_ver'
        $valLatest.Margin = [System.Windows.Thickness]::new(0,1,0,1)
        [System.Windows.Controls.Grid]::SetRow($valLatest, $row); [System.Windows.Controls.Grid]::SetColumn($valLatest, 1)
        $propsGrid.Children.Add($valLatest) | Out-Null

        $sp.Children.Add($propsGrid) | Out-Null

        if (-not $e.OK) {
            $h = [System.Windows.Controls.TextBlock]::new()
            $h.Text = "`n  Not installed. Use [Install / Update PS7] button above."
            $h.Foreground = '#F0A030'; $h.FontSize = 11
            $sp.Children.Add($h) | Out-Null
        }
        $brd.Child = $sp
        $pnlEng.Children.Add($brd) | Out-Null
    }

    # Add winget card
    $wbrd = [System.Windows.Controls.Border]::new()
    $wbrd.Background = '#0A1C0A'; $wbrd.Margin = [System.Windows.Thickness]::new(0,0,0,8)
    $wbrd.Padding = [System.Windows.Thickness]::new(14,10); $wbrd.Tag = 'WINGET'
    $wsp = [System.Windows.Controls.StackPanel]::new()

    $wt = [System.Windows.Controls.TextBlock]::new()
    $wt.Text = "[TOOL]  winget - Windows Package Manager"; $wt.FontSize = 14; $wt.FontWeight = 'Bold'
    $wt.Foreground = '#50A050'; $wt.Margin = [System.Windows.Thickness]::new(0,0,0,4)
    $wsp.Children.Add($wt) | Out-Null

    $wver = & winget.exe --version 2>$null
    $wval = if ($wver) { $wver.Trim() } else { 'Not installed' }
    $wtb = [System.Windows.Controls.TextBlock]::new()
    $wtb.Text = "  Installed Version  :  $wval"; $wtb.Foreground = '#A0A0D0'
    $wtb.FontFamily = 'Consolas'; $wtb.FontSize = 11
    $wsp.Children.Add($wtb) | Out-Null

    $wlat = [System.Windows.Controls.TextBlock]::new()
    $wlat.Text = "  Latest Available   :  (click Check winget)"; $wlat.Foreground = '#303060'
    $wlat.FontFamily = 'Consolas'; $wlat.FontSize = 11; $wlat.Tag = 'winget_latest'
    $wsp.Children.Add($wlat) | Out-Null

    $wbrd.Child = $wsp
    $pnlEng.Children.Add($wbrd) | Out-Null

    # - Bind rows -----------------
    $rowIndex = @{}
    foreach ($r in $allRows) { $rowIndex[$r.Name] = $r }

    # - Selection tracking --------------
    function UpdateSel {
        $n = ($allRows | Where-Object { $_.Sel }).Count
        $lblSel.Text = "$n module(s) selected"
        $btnInst.IsEnabled   = ($n -gt 0)
        $btnRemove.IsEnabled = ($n -gt 0)
    }
    foreach ($r in $allRows) {
        $r.add_PropertyChanged([System.ComponentModel.PropertyChangedEventHandler]{
            param($s,$e)
            if ($e.PropertyName -eq 'Sel') { UpdateSel }
        })
    }

    # - Category chips ----------------
    $script:CatFilter = 'All'
    # allCats rebuilds from LIVE catalog (includes custom modules added at runtime)
    function GetLiveCats {
        @('All') + ($script:LiveCatalog | Select-Object -ExpandProperty Cat -Unique | Sort-Object)
    }
    # LiveCatalog starts as the passed-in catalog and grows when user adds custom modules
    $script:LiveCatalog = [System.Collections.Generic.List[object]]::new()
    foreach ($m in $catalog) { $script:LiveCatalog.Add($m) }
    $allCats = GetLiveCats

    function MakeChips {
        $pnlCats.Children.Clear()
        $allCats = GetLiveCats
        foreach ($cat in $allCats) {
            $b = [System.Windows.Controls.Button]::new()
            $b.Content         = $cat
            $b.Tag             = $cat
            $b.Margin          = [System.Windows.Thickness]::new(0,0,5,0)
            $b.Padding         = [System.Windows.Thickness]::new(10,2,10,2)
            $b.FontSize        = 11
            $b.Cursor          = [System.Windows.Input.Cursors]::Hand
            $b.BorderThickness = [System.Windows.Thickness]::new(1)
            $isAct = ($cat -eq $script:CatFilter)
            $isLight = ($script:CurrentTheme -eq 'Light')
            $fgCol  = if ($isAct)   { '#FFFFFF' } elseif ($isLight) { '#2A2A6A'  } else { '#555590' }
            $bgCol  = if ($isAct)   { '#1A3CC0' } elseif ($isLight) { '#D8D8F0'  } else { '#0E0E28' }
            $brCol  = if ($isAct)   { '#3A5AFF' } elseif ($isLight) { '#A0A0D0'  } else { '#1A1A40' }
            $conv2  = [System.Windows.Media.BrushConverter]::new()
            try { $b.Foreground  = $conv2.ConvertFromString($fgCol) } catch {}
            try { $b.Background  = $conv2.ConvertFromString($bgCol) } catch {}
            try { $b.BorderBrush = $conv2.ConvertFromString($brCol) } catch {}
            # Capture cat value in closure by assigning to button Tag
            $b.add_Click({
                param($sender, $evArgs)
                $script:CatFilter = $sender.Tag
                MakeChips
                DoFilter
            })
            $pnlCats.Children.Add($b) | Out-Null
        }
    }

    # - Filter ------------------
    function DoFilter {
        $ft  = $txtFilter.Text.Trim()
        $oi  = [bool]$chkInst.IsChecked
        $ou  = [bool]$chkUpd.IsChecked
        $cat = $script:CatFilter

        $col = [System.Collections.ObjectModel.ObservableCollection[PSMod.Row]]::new()
        foreach ($r in $allRows) {
            if ($cat -ne 'All' -and $r.Cat -ne $cat) { continue }
            if ($ft -and $r.Name -notlike "*$ft*" -and $r.Desc -notlike "*$ft*") { continue }
            if ($oi -and -not $r.V5 -and -not $r.V7) { continue }
            if ($ou -and -not $r.Up) { continue }
            $col.Add($r)
        }
        $dg.ItemsSource = $col
    }

    MakeChips
    DoFilter

    # - Theme engine -
    $Global:ThemeColors = @{
        Dark  = @{ WinBg='#0F0F1A'; TitleBg='#07071A'; ToolBg='#0C0C1E'; RibbonBg='#09091F'
                   GridBg='#0A0A18'; RowBg='#0E0E22'; AltBg='#131328'; Fg='#C0C0E0'
                   Border='#181838'; AccentFg='#5A88FF' }
        Light = @{ WinBg='#F0F0F8'; TitleBg='#E0E0F0'; ToolBg='#E8E8F8'; RibbonBg='#DCDCF0'
                   GridBg='#FFFFFF'; RowBg='#F8F8FF'; AltBg='#EEEEFF'; Fg='#1A1A3A'
                   Border='#C0C0E0'; AccentFg='#1A3CC0' }
    }

    function ApplyTheme([string]$ThemeName) {
        $t = if ($ThemeName -eq 'Auto') {
            try {
                $reg = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -EA SilentlyContinue
                if ($reg -and $reg.AppsUseLightTheme -eq 0) { 'Dark' } else { 'Light' }
            } catch { 'Dark' }
        } else { $ThemeName }

        $col = $Global:ThemeColors[$t]
        if (-not $col) { return }

        $conv = [System.Windows.Media.BrushConverter]::new()
        function Br([string]$hex) { $conv.ConvertFromString($hex) }

        # Apply to all named panels - walk the visual tree
        # Helper to set bg on any panel/border by name
        function SetBg([string]$name, [string]$hex) {
            $el = $win.FindName($name)
            if ($el) {
                try { $el.Background = Br $hex } catch {}
            }
        }

        $win.Background = Br $col.WinBg

        # Title row
        $titleBorder = $win.FindName('lblPs7')
        if ($titleBorder) {
            # Walk up to the border
            $p = $titleBorder.Parent
            while ($p -and $p -isnot [System.Windows.Controls.Border]) { $p = $p.Parent }
            if ($p) { $p.Background = Br $col.TitleBg }
        }

        # DataGrid
        $dg.Background               = Br $col.GridBg
        $dg.RowBackground            = Br $col.RowBg
        $dg.AlternatingRowBackground = Br $col.AltBg
        $dg.Foreground               = Br $col.Fg

        # Tab control + tab items
        $tabC.Background    = Br $col.ToolBg
        $tabC.BorderBrush   = Br $col.Border

        # Operation log pane
        $txtMini.Foreground = Br $col.Fg
        $txtMini.Background = Br $col.GridBg
        $svMini.Background  = Br $col.GridBg

        # Log viewer
        $txtLog.Foreground  = Br $col.Fg
        $txtLog.Background  = Br $col.GridBg
        $svLog.Background   = Br $col.GridBg

        # Label colors
        $lblStatus.Foreground = Br $col.Fg
        $lblSel.Foreground    = Br $col.Fg
        $lblPs7.Foreground    = if ($t -eq 'Light') { Br '#1A6030' } else { Br '#F0A030' }

        # Category chips - recreate with new theme colors (avoids template property errors)
        MakeChips

        # Engine info panel
        $pnlEng.Background = Br $col.GridBg

        # Toolbar border - walk all direct Border children of grid
        $grid = $win.Content
        if ($grid -is [System.Windows.Controls.Grid]) {
            foreach ($child in $grid.Children) {
                if ($child -is [System.Windows.Controls.Border]) {
                    $row = [System.Windows.Controls.Grid]::GetRow($child)
                    switch ($row) {
                        0 { $child.Background = Br $col.TitleBg; $child.BorderBrush = Br $col.Border }
                        1 { $child.Background = Br $col.RibbonBg; $child.BorderBrush = Br $col.Border }
                        2 { $child.Background = Br $col.ToolBg; $child.BorderBrush = Br $col.Border }
                        4 { $child.Background = Br $col.GridBg }
                        5 { $child.Background = Br $col.TitleBg }
                    }
                }
            }
        }

        $script:CurrentTheme = $t
        ML "Theme applied: $t"
    }

    $cmbTheme.add_SelectionChanged({
        $sel = ($cmbTheme.SelectedItem).Content
        ApplyTheme $sel
    })

    # Apply initial theme (Dark)
    ApplyTheme 'Dark'


    # - Row click toggles checkbox ------------
    $dg.add_MouseLeftButtonUp({
        param($sender, $evArgs)
        $pos = $evArgs.GetPosition($dg)
        $hit = $dg.InputHitTest($pos)
        if ($null -eq $hit) { return }
        # Walk up to DataGridRow
        $el = $hit
        $maxWalk = 15
        while ($el -ne $null -and $maxWalk -gt 0) {
            $maxWalk--
            if ($el -is [System.Windows.Controls.DataGridRow]) {
                $row = $el.Item
                if ($row -is [PSMod.Row]) {
                    # Check we didn't click the checkbox column itself
                    $cell = $hit
                    $cellWalk = 10
                    while ($cell -ne $null -and $cellWalk -gt 0) {
                        $cellWalk--
                        if ($cell -is [System.Windows.Controls.DataGridCell]) { break }
                        $cell = [System.Windows.Media.VisualTreeHelper]::GetParent($cell)
                    }
                    # Only toggle if NOT clicking on checkbox cell (col index 0)
                    $skipToggle = $false
                    if ($cell -is [System.Windows.Controls.DataGridCell]) {
                        if ($cell.Column.DisplayIndex -eq 0) { $skipToggle = $true }
                    }
                    if (-not $skipToggle) { $row.Sel = -not $row.Sel }
                }
                break
            }
            $el = [System.Windows.Media.VisualTreeHelper]::GetParent($el)
        }
    })

    # - Helper: current engine/scope -----------
    function GetEng {
        $it = $cmbEng.SelectedItem
        if ($it -and $it.Tag) { return $it.Tag }
        return ($engines | Where-Object { $_.OK } | Select-Object -First 1)
    }
    function GetScope {
        if (($cmbScope.SelectedItem).Content -like '*AllUsers*') { 'AllUsers' } else { 'CurrentUser' }
    }

    $lblStatus.Text = "Ready  |  Log: $logFile"
    $cmbEng.add_SelectionChanged({
        $e = GetEng; if ($e) { $lblExe.Text = $e.Exe }
    })

    # - SCAN  (Runspace + DispatcherTimer polling) ------
    $script:ScanTimer  = $null
    $script:ScanCancel = $false
    $script:ScanRS     = $null
    $script:ScanPS     = $null
    $script:ScanAR     = $null
    $script:ScanTotal  = 0
    $script:ScanDone   = 0

    $btnScan.add_Click({
        $eng = GetEng
        if (-not $eng.OK) {
            [System.Windows.MessageBox]::Show('Selected engine is not available.','Engine Error',
                [System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        $btnScan.IsEnabled  = $false
        $btnStopScan.IsEnabled = $true
        $pbar.Visibility    = 'Visible'
        $pbar.Value         = 0
        $script:ScanDone    = 0
        $script:ScanTotal   = $allRows.Count
        $script:ScanCancel  = $false
        $lblStatus.Text     = 'Scanning installed modules...'
        foreach ($r in $allRows) { $r.St = '...' }

        $e5obj = $engines | Where-Object { $_.Tag -eq 'PS5' -and $_.OK } | Select-Object -First 1
        $e7obj = $engines | Where-Object { $_.Tag -eq 'PS7' -and $_.OK } | Select-Object -First 1
        $ex5   = if ($e5obj) { $e5obj.Exe } else { '' }
        $ex7   = if ($e7obj) { $e7obj.Exe } else { '' }
        $nameList = ($allRows | ForEach-Object { $_.Name }) -join '|'

        $script:ScanRS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:ScanRS.ApartmentState = 'MTA'
        $script:ScanRS.ThreadOptions  = 'ReuseThread'
        $script:ScanRS.Open()
        $script:ScanPS = [System.Management.Automation.PowerShell]::Create()
        $script:ScanPS.Runspace = $script:ScanRS

        [void]$script:ScanPS.AddScript({
            param($ex5, $ex7, $nameList, $resultQ, $termQ, $logF, $cancelRef)
            $ep = 'SilentlyContinue'
            $ErrorActionPreference = $ep
            $ProgressPreference    = $ep

            $names = $nameList -split '\|'

            function LogIt([string]$m, [string]$tag = 'SCAN') {
                $lts = Get-Date -f 'HH:mm:ss.fff'
                try { Add-Content -Path $logF -Value "[$lts][$tag] $m" -Encoding UTF8 } catch {}
                try { $termQ.Enqueue("[$tag] $m") } catch {}
            }

            # - STEP 1: Get ALL installed modules in ONE call per engine -
            # Use Get-Module -ListAvailable (fast, local only) + Get-InstalledModule for PSGallery scope
            # This avoids 65 separate process spawns
            function GetInstalledBatch([string]$Exe, [string]$Tag) {
                if (-not $Exe) { return @{} }
                LogIt "[$Tag] Batch scan via $Exe ..."

                # Build PS command that:
                # 1. Gets ALL installed modules at once with Get-Module -ListAvailable
                # 2. Also checks Get-InstalledModule for accurate scope
                # Output format: NAME|VER|SCOPE
                $cmd = (
                    '$ep="SilentlyContinue";$ErrorActionPreference=$ep;$ProgressPreference=$ep;' +
                    # PSGet is authoritative: -AllVersions then pick highest per module
                    '$psg=@{};' +
                    'Get-InstalledModule -AllVersions -EA $ep 2>$null|' +
                    'Sort-Object Name,Version|Group-Object Name|' +
                    'ForEach-Object{$pi=$_.Group|Sort-Object Version -Desc|Select-Object -First 1;$psg[$pi.Name]=$pi};' +
                    # Get-Module -ListAvailable for built-in / WinPS modules not in PSGet
                    '$lav=Get-Module -ListAvailable -EA $ep 2>$null|' +
                    'Sort-Object Name,Version|Group-Object Name|' +
                    'ForEach-Object{$_.Group|Sort-Object Version -Desc|Select-Object -First 1};' +
                    '$seen=@{};' +
                    # Output PSGet modules first (accurate version after updates)
                    'foreach($pi in $psg.Values){' +
                    '  $sc="Unknown";$base=$pi.InstalledLocation;' +
                    '  if($base -like "*\Windows\System32\*" -or $base -like "*\Windows\SysWOW64\*"){$sc="System"}' +
                    '  elseif($base -like "*\Program Files\*"){$sc="AllUsers"}' +
                    '  elseif($base -like "*\Documents\*" -or $base -like "*\CurrentUser\*" -or $base -like "*\AppData\*"){$sc="CurrentUser"}' +
                    '  elseif($base -like "*\AllUsers\*"){$sc="AllUsers"}' +
                    '  elseif($base){' +
                    '    $u=[System.Environment]::GetFolderPath("UserProfile");' +
                    '    if($base.StartsWith($u)){$sc="CurrentUser"}else{$sc="AllUsers"}' +
                    '  };' +
                    '  $seen[$pi.Name]=1;' +
                    '  Write-Output ($pi.Name+"|"+$pi.Version.ToString()+"|"+$sc+"|"+$base)' +
                    '};' +
                    # Then built-in WinPS/PS7 modules not tracked by PSGet
                    'foreach($m in $lav){' +
                    '  if($seen.ContainsKey($m.Name)){continue};' +
                    '  $sc="Unknown";$base=$m.ModuleBase;' +
                    '  if($base -like "*\Windows\System32\*" -or $base -like "*\Windows\SysWOW64\*"){$sc="System"}' +
                    '  elseif($base -like "*\Program Files\WindowsPowerShell\*"){$sc="WinPS-System"}' +
                    '  elseif($base -like "*\Program Files\PowerShell\*"){$sc="PS7-System"}' +
                    '  elseif($base -like "*\WindowsPowerShell\Modules*"){$sc="AllUsers"}' +
                    '  elseif($base -like "*\Documents\*" -or $base -like "*\CurrentUser\*"){$sc="CurrentUser"};' +
                    '  Write-Output ($m.Name+"|"+$m.Version.ToString()+"|"+$sc+"|"+$base)' +
                    '}'

                )
                $res = @{}
                try {
                    $out = & $Exe -NoProfile -NonInteractive -Command $cmd 2>$null
                    foreach ($ln in $out) {
                        $x = $ln.ToString().Trim()
                        $parts = $x.Split('|', 4)
                        if ($parts.Count -ge 3 -and $parts[0] -and $parts[1]) {
                            $res[$parts[0]] = @{ Ver=$parts[1]; Scope=$parts[2]; Base=if($parts.Count -ge 4){$parts[3]}else{''} }
                        }
                    }
                } catch { LogIt "[$Tag] ERROR: $_" }
                LogIt "[$Tag] Found $($res.Count) installed modules total."
                return $res
            }

            # - STEP 2: Gallery check for installed modules only -
            function GetGalleryBatch([string]$Exe, [string[]]$ModNames) {
                if (-not $Exe -or $ModNames.Count -eq 0) { return @{} }
                LogIt "Gallery: checking $($ModNames.Count) modules via $Exe ..."
                $nameArr = ($ModNames | ForEach-Object { '"' + ($_ -replace '"','') + '"' }) -join ','
                $cmd = (
                    '$ep="SilentlyContinue";$ErrorActionPreference=$ep;$ProgressPreference=$ep;' +
                    '$names=@(' + $nameArr + ');' +
                    # Use Find-Module with batch - much faster than one-by-one
                    'try{$found=Find-Module -Name $names -EA $ep 2>$null;' +
                    'foreach($m in $found){Write-Output ($m.Name+"|"+$m.Version.ToString())}}' +
                    'catch{' +
                    # Fallback: one by one if batch fails
                    'foreach($n in $names){' +
                    '  $g=Find-Module -Name $n -EA $ep 2>$null|Select-Object -First 1;' +
                    '  if($g){Write-Output ($n+"|"+$g.Version.ToString())}' +
                    '}}'
                )
                $gal = @{}
                try {
                    $out = & $Exe -NoProfile -NonInteractive -Command $cmd 2>$null
                    foreach ($ln in $out) {
                        $x = $ln.ToString().Trim()
                        $p = $x.Split('|', 2)
                        if ($p.Count -eq 2 -and $p[0] -and $p[1]) { $gal[$p[0]] = $p[1] }
                    }
                } catch { LogIt "Gallery ERROR: $_" }
                LogIt "Gallery: got versions for $($gal.Count)/$($ModNames.Count) modules."
                return $gal
            }

            if ($cancelRef.Value) { LogIt "Cancelled before start."; return }

            LogIt "=== SCAN START: $($names.Count) catalog modules ==="

            $res5 = GetInstalledBatch $ex5 'PS5'
            if ($cancelRef.Value) { LogIt "Cancelled after PS5 scan."; return }

            $res7 = GetInstalledBatch $ex7 'PS7'
            if ($cancelRef.Value) { LogIt "Cancelled after PS7 scan."; return }

            # Only gallery-check modules that are installed in at least one engine
            $installedNames = @($names | Where-Object {
                ($res5.ContainsKey($_) -and $res5[$_].Ver) -or
                ($res7.ContainsKey($_) -and $res7[$_].Ver)
            })
            LogIt "Installed in catalog: $($installedNames.Count) modules"

            $galExe = if ($ex7) { $ex7 } else { $ex5 }
            $gal = @{}
            if (-not $cancelRef.Value) {
                $gal = GetGalleryBatch $galExe $installedNames
            }

            LogIt "Building results..."
            foreach ($n in $names) {
                if ($cancelRef.Value) { LogIt "Cancelled during result build."; break }

                $v5  = if ($res5.ContainsKey($n)) { $res5[$n].Ver   } else { '' }
                $sc5 = if ($res5.ContainsKey($n)) { $res5[$n].Scope } else { '' }
                $v7  = if ($res7.ContainsKey($n)) { $res7[$n].Ver   } else { '' }
                $sc7 = if ($res7.ContainsKey($n)) { $res7[$n].Scope } else { '' }
                $gv  = if ($gal.ContainsKey($n))  { $gal[$n]        } else { '' }
                $up  = $false

                # Determine base path (for Open Folder)
                $base = ''
                if ($res7.ContainsKey($n) -and $res7[$n].Base) { $base = $res7[$n].Base }
                elseif ($res5.ContainsKey($n) -and $res5[$n].Base) { $base = $res5[$n].Base }

                $instV = if ($v7) { $v7 } elseif ($v5) { $v5 } else { $null }
                if ($instV -and $gv) {
                    try { $up = ([Version]($gv -replace '[^0-9.]','')) -gt ([Version]($instV -replace '[^0-9.]','')) } catch {}
                }

                $st = if (-not $v5 -and -not $v7) { 'Not Installed' }
                      elseif ($up) { 'Update Available' }
                      else { 'Up to Date' }

                $pth = if($b5){" base=$b5"}else{if($b7){" base=$b7"}else{""}}
                LogIt "$n | PS5=$v5($sc5) PS7=$v7($sc7) Gal=$gv | $st$pth" "RESULT"

                $resultQ.Enqueue([PSCustomObject]@{
                    Name=$n; V5=$v5; S5=$sc5; V7=$v7; S7=$sc7
                    GV=$gv; St=$st; Up=$up; Base=$base
                })
            }
            LogIt "=== SCAN COMPLETE ==="
        })
        [void]$script:ScanPS.AddParameters(@{
            ex5        = $ex5
            ex7        = $ex7
            nameList   = $nameList
            resultQ    = $scanQ
            termQ      = $termQ
            logF       = $logFile
            cancelRef  = ([ref]$script:ScanCancel)
        })

        $script:ScanAR = $script:ScanPS.BeginInvoke()
        $script:LogReloadCount = 0

        $script:ScanTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:ScanTimer.Interval = [System.TimeSpan]::FromMilliseconds(300)
        $script:ScanTimer.add_Tick({
            $script:LogReloadCount++
            if ($script:LogReloadCount -ge 6) { $script:LogReloadCount = 0; ReloadLogView }

            # Drain terminal queue - shows verbose output from scan runspace
            $tLine = $null
            $tCount = 0
            while ($tCount -lt 50 -and $termQ.TryDequeue([ref]$tLine)) { TL $tLine 'SCAN'; $tCount++ }

            $item = $null
            while ($scanQ.TryDequeue([ref]$item)) {
                $row = $rowIndex[$item.Name]
                if ($row) {
                    $row.V5 = $item.V5; $row.S5 = $item.S5
                    $row.V7 = $item.V7; $row.S7 = $item.S7
                    $row.GV = $item.GV; $row.St = $item.St
                    $row.Up = $item.Up
                    if ($item.Base) { $rowIndex['__base__' + $item.Name] = $item.Base }
                }
                $script:ScanDone++
                $pbar.Value   = [int](($script:ScanDone / $script:ScanTotal) * 100)
                $lblProg.Text = "$($script:ScanDone) / $($script:ScanTotal)"
            }

            if ($script:ScanAR -ne $null -and $script:ScanAR.IsCompleted) {
                $script:ScanTimer.Stop()
                try { $script:ScanPS.EndInvoke($script:ScanAR) | Out-Null; $script:ScanPS.Dispose()
                      $script:ScanRS.Close(); $script:ScanRS.Dispose() } catch {}

                DoFilter; ReloadLogView
                $btnScan.IsEnabled     = $true
                $btnStopScan.IsEnabled = $false
                $pbar.Visibility       = 'Collapsed'
                $lblProg.Text          = ''

                $inst = ($allRows | Where-Object { $_.V5 -or $_.V7 }).Count
                $upd  = ($allRows | Where-Object { $_.Up }).Count
                $ni   = $allRows.Count - $inst
                $msg  = "Done  Installed:$inst  Updates:$upd  Missing:$ni"
                TLHead "SCAN COMPLETE"
                TL "Installed: $inst  Updates Available: $upd  Not Installed: $ni"
                TL "Gallery versions checked for: $($allRows | Where-Object {$_.GV} | Measure-Object | Select-Object -ExpandProperty Count) modules"
                $lblStatus.Text = $msg
                ML $msg
            }
        })
        $script:ScanTimer.Start()
    })

    $btnStopScan.add_Click({
        $script:ScanCancel = $true
        $btnStopScan.IsEnabled = $false
        ML "Stop requested - finishing current phase..." 'WARN'
        $lblStatus.Text = 'Stopping scan...'
    })


    # - Filter events ----------------
    $txtFilter.add_TextChanged({ DoFilter })
    $chkInst.add_Checked({ DoFilter });   $chkInst.add_Unchecked({ DoFilter })
    $chkUpd.add_Checked({  DoFilter });   $chkUpd.add_Unchecked({  DoFilter })

    # - Selection buttons ---------------
    $btnAll.add_Click({
        foreach ($r in $dg.ItemsSource) { $r.Sel = $true }
        UpdateSel
    })
    $btnNone.add_Click({
        foreach ($r in $allRows) { $r.Sel = $false }
        UpdateSel
    })
    $btnSelUpd.add_Click({
        foreach ($r in $allRows) { $r.Sel = $r.Up }
        UpdateSel
    })
    $btnSelMis.add_Click({
        foreach ($r in $allRows) { $r.Sel = (-not $r.V5 -and -not $r.V7) }
        UpdateSel
    })

    # - Install / Update batch -------------
    $script:BatchTimer = $null
    $script:BatchRS    = $null
    $script:BatchPS    = $null
    $script:BatchAR    = $null
    $script:BatchTotal = 0
    $script:BatchDone  = 0

    function StartBatch([object[]]$rows, [string]$exe, [string]$scope) {
        $btnInst.IsEnabled   = $false
        $btnUpdAll.IsEnabled = $false
        $btnRemove.IsEnabled = $false
        $btnCancel.IsEnabled = $true
        $pbar.Visibility     = 'Visible'
        $pbar.Value          = 0
        $script:BatchTotal   = $rows.Count
        $script:BatchDone    = 0

        ML "Batch: $($rows.Count) modules via $exe  scope=$scope"
        TLHead "BATCH INSTALL/UPDATE: $($rows.Count) modules  scope=$scope"
        TL "Engine: $exe"
        foreach ($r in $rows) { TL "  -> $($r.Name)" 'QUEUE' }

        $script:BatchRS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:BatchRS.ApartmentState = 'MTA'
        $script:BatchRS.ThreadOptions  = 'ReuseThread'
        $script:BatchRS.Open()

        $script:BatchPS = [System.Management.Automation.PowerShell]::Create()
        $script:BatchPS.Runspace = $script:BatchRS

        [void]$script:BatchPS.AddScript({
            param($rows, $exe, $scope, $resultQ, $termQ2, $cancelRef, $logF)
            $ErrorActionPreference = 'SilentlyContinue'
            $ProgressPreference    = 'SilentlyContinue'
            function BLog([string]$m) {
                $ts = Get-Date -f 'HH:mm:ss.fff'
                try { Add-Content -Path $logF -Value "[$ts][BATCH] $m" -Encoding UTF8 } catch {}
                try { $termQ2.Enqueue("[BATCH] $m") } catch {}
            }
            BLog "=== BATCH START: $($rows.Count) items | exe=$exe | scope=$scope ==="
            foreach ($row in $rows) {
                if ($cancelRef.Value) {
                    try { Add-Content -Path $logF -Value "[$(Get-Date -f 'HH:mm:ss')] Cancelled." -Encoding UTF8 } catch {}
                    break
                }
                $sc = (
                    '$ep="Stop";$n="' + $row.Name + '";$s="' + $scope + '";' +
                    '$ErrorActionPreference=$ep;$ProgressPreference="SilentlyContinue";' +
                    '$ep2="SilentlyContinue";$n2=$n;$s2=$s;' +
                    'try{' +
                    '  $all=Get-InstalledModule -Name $n2 -AllVersions -EA $ep2 2>$null;' +
                    '  $i=$all|Sort-Object Version -Desc|Select-Object -First 1;' +
                    '  if($i){' +
                    '    $g=Find-Module -Name $n2 -EA $ep2 2>$null|Select-Object -First 1;' +
                    '    if($g -and ([Version]$g.Version -gt [Version]$i.Version)){' +
                    '      Install-Module -Name $n2 -Scope $s2 -Force -AllowClobber -SkipPublisherCheck;' +
                    '      $newVer=(Get-InstalledModule -Name $n2 -EA $ep2|Sort-Object Version -Desc|Select-Object -First 1).Version;' +
                    '      $old=$all|Where-Object{[Version]$_.Version -lt [Version]$newVer};' +
                    '      foreach($o in $old){try{Uninstall-Module -Name $n2 -RequiredVersion $o.Version -Force -EA $ep2}catch{}};' +
                    '      Write-Output ("UPDATED|"+$newVer)' +
                    '    } else{Write-Output ("UPTODATE|"+$i.Version)}' +
                    '  } else{' +
                    '    Install-Module -Name $n2 -Scope $s2 -Force -AllowClobber -SkipPublisherCheck;' +
                    '    $v=(Get-InstalledModule -Name $n2 -EA $ep2|Sort-Object Version -Desc|Select-Object -First 1).Version;' +
                    '    Write-Output ("INSTALLED|"+$v)' +
                    '  }' +
                    '} catch{Write-Output ("ERROR|"+$_.Exception.Message.Replace("`n"," "))}'
                )
                $res = 'ERROR|Unknown'
                try {
                    $out = & $exe -NoProfile -NonInteractive -Command $sc 2>&1
                    foreach ($ln in $out) {
                        $x = $ln.ToString().Trim()
                        if ($x) { BLog "  RAW: $x" }
                        if ($x -match '^\w+\|') { $res = $x; break }
                    }
                } catch { $res = "ERROR|$_" }

                $parts  = $res.Split('|', 2)
                $status = $parts[0]
                $detail = if ($parts.Count -ge 2) { $parts[1] } else { '' }
                BLog "  RESULT: $status $($row.Name) $detail"
                $resultQ.Enqueue([PSCustomObject]@{Name=$row.Name;Status=$status;Detail=$detail;Exe=$exe})
            }
        })
        [void]$script:BatchPS.AddParameters(@{
            rows      = $rows
            exe       = $exe
            scope     = $scope
            resultQ   = $batchQ
            termQ2    = $termQ
            cancelRef = $cancelBRef
            logF      = $logFile
        })

        $script:BatchAR = $script:BatchPS.BeginInvoke()

        $script:BatchTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:BatchTimer.Interval = [System.TimeSpan]::FromMilliseconds(300)
        $script:BatchTimer.add_Tick({
            # Drain terminal queue from batch subprocess
            $tLine = $null
            $tCount = 0
            while ($tCount -lt 100 -and $termQ.TryDequeue([ref]$tLine)) {
                if ($tLine -match '^\[BATCH\] (.+)$') { TL $Matches[1] 'BATCH' }
                else { TL $tLine 'BATCH' }
                $tCount++
            }

            $item = $null
            while ($batchQ.TryDequeue([ref]$item)) {
                $script:BatchDone++
                $pbar.Value   = [int](($script:BatchDone / $script:BatchTotal) * 100)
                $lblProg.Text = "$($script:BatchDone) / $($script:BatchTotal)"

                $row = $rowIndex[$item.Name]
                $msg = ''
                switch ($item.Status) {
                    'INSTALLED' {
                        $msg = "[OK] INSTALLED $($item.Name) v$($item.Detail)"
                        TL "INSTALLED $($item.Name)  ->  v$($item.Detail)  [scope=$scope]" 'OK'
                        if ($row) {
                            $row.St = 'Up to Date'; $row.Up = $false
                            if ($item.Exe -like '*pwsh*') { $row.V7 = $item.Detail; $row.S7 = $scope }
                            else { $row.V5 = $item.Detail; $row.S5 = $scope }
                        }
                    }
                    'UPDATED' {
                        $msg = "[UP] UPDATED $($item.Name) v$($item.Detail)"
                        TL "UPDATED  $($item.Name)  ->  v$($item.Detail)" 'UP'
                        if ($row) {
                            $row.St = 'Up to Date'; $row.Up = $false
                            if ($item.Exe -like '*pwsh*') { $row.V7 = $item.Detail; $row.S7 = $scope }
                            else { $row.V5 = $item.Detail; $row.S5 = $scope }
                        }
                    }
                    'UPTODATE' {
                        $msg = "[==] UP TO DATE $($item.Name) v$($item.Detail)"
                        TL "SKIP  $($item.Name)  already v$($item.Detail)" '=='
                        if ($row) { $row.Up = $false; $row.St = 'Up to Date' }
                    }
                    default    {
                        $msg = "[!!] $($item.Status) $($item.Name): $($item.Detail)"
                        TL "ERROR  $($item.Name): $($item.Detail)" 'ERR'
                        if ($row) { $row.St = "ERROR: $($item.Detail.Substring(0,[Math]::Min(40,$item.Detail.Length)))" }
                    }
                }
                if ($msg) { ML $msg }
            }

            if ($script:BatchAR -ne $null -and $script:BatchAR.IsCompleted) {
                $script:BatchTimer.Stop()
                try {
                    $script:BatchPS.EndInvoke($script:BatchAR) | Out-Null
                    $script:BatchPS.Dispose()
                    $script:BatchRS.Close(); $script:BatchRS.Dispose()
                } catch {}
                $btnInst.IsEnabled   = $true
                $btnUpdAll.IsEnabled = $true
                $btnRemove.IsEnabled = $true
                $btnCancel.IsEnabled = $false
                $pbar.Visibility     = 'Collapsed'
                $lblProg.Text        = ''
                $lblStatus.Text      = "Batch done: $($script:BatchDone)/$($script:BatchTotal)"
                ML "Batch complete: $($script:BatchDone)/$($script:BatchTotal)"
                DoFilter; ReloadLogView
                # Auto-rescan to confirm actual installed versions
                TL '--- Batch complete. Auto-refreshing scan...' 'BATCH'
                $window.Dispatcher.BeginInvoke([System.Action]{
                    $btnRefresh.RaiseEvent([System.Windows.RoutedEventArgs]::new(
                        [System.Windows.Controls.Button]::ClickEvent))
                }, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
            }
        })
        $script:BatchTimer.Start()
    }

    # Smart engine routing: use PS7 exe for all (it can see all scopes)
    # Fall back to PS5 for modules with PS5 version only when PS7 unavailable
    function GetBestExe([object]$row) {
        $e7 = $engines | Where-Object { $_.Tag -eq 'PS7' -and $_.OK } | Select-Object -First 1
        $e5 = $engines | Where-Object { $_.Tag -eq 'PS5' -and $_.OK } | Select-Object -First 1
        # Prefer PS7 (handles both AllUsers and CurrentUser in PS7+ modules)
        if ($e7) { return $e7.Exe }
        if ($e5) { return $e5.Exe }
        return (GetEng).Exe
    }

    # RunSmartBatch: run PS7 pass (for V7 modules) + PS5 pass (for V5-only modules)
    function RunSmartBatch([object[]]$rows, [string]$scope) {
        $e7 = $engines | Where-Object { $_.Tag -eq 'PS7' -and $_.OK } | Select-Object -First 1
        $e5 = $engines | Where-Object { $_.Tag -eq 'PS5' -and $_.OK } | Select-Object -First 1

        # Separate: modules with PS7 version go to PS7, PS5-only go to PS5
        $ps7rows = @($rows | Where-Object { $_.V7 })    # has PS7 version -> update via PS7
        $ps5only = @($rows | Where-Object { $_.V5 -and -not $_.V7 })  # PS5 only -> update via PS5
        # Modules with BOTH V5 and V7: also need PS5 update
        $ps5also = @($rows | Where-Object { $_.V5 -and $_.V7 })

        TL "=== SMART BATCH: PS7=$($ps7rows.Count) PS5-only=$($ps5only.Count) Both=$($ps5also.Count) ===" 'BATCH'

        if ($ps7rows.Count -gt 0 -and $e7) {
            StartBatch $ps7rows $e7.Exe $scope
        } elseif ($ps7rows.Count -gt 0 -and $e5) {
            StartBatch $ps7rows $e5.Exe $scope  # fallback
        }
        if ($ps5only.Count -gt 0 -and $e5) {
            StartBatch $ps5only $e5.Exe $scope
        }
        if ($ps5also.Count -gt 0 -and $e5 -and $e7) {
            # Run PS5 update for the "both" modules too (PS7 pass already queued above)
            $script:PendingPS5Batch = @{ rows=$ps5also; exe=$e5.Exe; scope=$scope }
        }
    }

    $btnInst.add_Click({
        $scope = GetScope
        $sel = @($allRows | Where-Object { $_.Sel })
        if ($sel.Count -eq 0) { [System.Windows.MessageBox]::Show('No modules selected.') | Out-Null; return }
        $e7 = $engines | Where-Object { $_.Tag -eq 'PS7' -and $_.OK } | Select-Object -First 1
        $e5 = $engines | Where-Object { $_.Tag -eq 'PS5' -and $_.OK } | Select-Object -First 1
        $engDesc = if ($e7) { "PS7: $($e7.Exe)" } else { "PS5: $($e5.Exe)" }
        $ans = [System.Windows.MessageBox]::Show(
            "Install/Update $($sel.Count) module(s)`n`nWill update each module via its matching engine:`n  $engDesc`nScope: $scope`n`nContinue?",
            'Confirm', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($ans -eq 'Yes') { RunSmartBatch $sel $scope }
    })

    $btnUpdAll.add_Click({
        $scope = GetScope
        $upd = @($allRows | Where-Object { $_.Up })
        if ($upd.Count -eq 0) {
            [System.Windows.MessageBox]::Show('No updates found. Run Refresh first.') | Out-Null; return
        }
        $e7 = $engines | Where-Object { $_.Tag -eq 'PS7' -and $_.OK } | Select-Object -First 1
        $e5 = $engines | Where-Object { $_.Tag -eq 'PS5' -and $_.OK } | Select-Object -First 1
        $engDesc = if ($e7) { "PS7: $($e7.Exe)" } else { "PS5: $($e5.Exe)" }
        $ans = [System.Windows.MessageBox]::Show(
            "Update ALL $($upd.Count) module(s)?`n`nWill update each module via its matching engine:`n  $engDesc`nScope: $scope",
            'Confirm All', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($ans -eq 'Yes') {
            foreach ($r in $upd) { $r.Sel = $true }
            RunSmartBatch $upd $scope
        }
    })

    $btnCancel.add_Click({
        $cancelBRef.Value = $true
        ML 'Cancel requested...' 'WARN'
    })


    $btnInfo.add_Click({
        # Create a styled info/help window
        $infoXaml = [xml]@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PS Module Manager v7.0 - Feature Guide"
        Width="820" Height="680" MinWidth="600" MinHeight="400"
        Background="#0A0A18" WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="13">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <Border Grid.Row="0" Background="#07071A" Padding="20,14" BorderBrush="#181838" BorderThickness="0,0,0,1">
      <StackPanel>
        <TextBlock FontSize="22" FontWeight="Bold" Foreground="#5A88FF">
          <Run Text="PS "/>
          <Run Text="Module Manager" Foreground="#FFFFFF"/>
          <Run Text="  v7.0" FontSize="14" Foreground="#303070"/>
        </TextBlock>
        <TextBlock Text="PowerShell module management GUI for IT professionals" Foreground="#404080" FontSize="12" Margin="0,4,0,0"/>
      </StackPanel>
    </Border>

    <!-- Scrollable content -->
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Background="#0A0A18" Padding="20,10,20,10">
      <StackPanel>

        <!-- SCAN -->
        <Border Background="#07071F" CornerRadius="6" Padding="16,12" Margin="0,6,0,0" BorderBrush="#181848" BorderThickness="1">
          <StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
              <TextBlock Text="REFRESH STATUS  " Foreground="#5A88FF" FontSize="14" FontWeight="Bold"/>
              <Border Background="#0E1A3A" CornerRadius="3" Padding="6,1"><TextBlock Text="Scan" Foreground="#3A6AFF" FontSize="11" FontFamily="Consolas"/></Border>
            </StackPanel>
            <TextBlock TextWrapping="Wrap" Foreground="#9090C0" LineHeight="20">
              <Run Text="Scans all 65 catalog modules across " Foreground="#9090C0"/>
              <Run Text="PS 5.1 (Windows PowerShell)" Foreground="#F0A030" FontWeight="SemiBold"/>
              <Run Text=" and " Foreground="#9090C0"/>
              <Run Text="PS 7 (PowerShell Core)" Foreground="#50D090" FontWeight="SemiBold"/>
              <Run Text=" simultaneously using batch process calls &#8212; just 3 process spawns total for the entire scan. Checks the PSGallery for latest available versions." Foreground="#9090C0"/>
            </TextBlock>
            <TextBlock Margin="0,8,0,0" TextWrapping="Wrap" Foreground="#505080">
              <Run Text="Scope detection: " Foreground="#606090" FontWeight="SemiBold"/>
              <Run Text="AllUsers " Foreground="#5080E0" FontFamily="Consolas"/>
              <Run Text="CurrentUser " Foreground="#50C080" FontFamily="Consolas"/>
              <Run Text="System " Foreground="#606090" FontFamily="Consolas"/>
              <Run Text="WinPS-System " Foreground="#505085" FontFamily="Consolas"/>
              <Run Text="PS7-System " Foreground="#505085" FontFamily="Consolas"/>
              <Run Text="Unknown " Foreground="#906030" FontFamily="Consolas"/>
              <Run Text="&#8212; empty = not installed in that engine" Foreground="#303060"/>
            </TextBlock>
          </StackPanel>
        </Border>

        <!-- INSTALL/UPDATE -->
        <Border Background="#07071F" CornerRadius="6" Padding="16,12" Margin="0,6,0,0" BorderBrush="#181848" BorderThickness="1">
          <StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
              <TextBlock Text="INSTALL / UPDATE  " Foreground="#50C080" FontSize="14" FontWeight="Bold"/>
              <Border Background="#0A200A" CornerRadius="3" Padding="6,1"><TextBlock Text="Smart Dual-Engine Batch" Foreground="#30A060" FontSize="11" FontFamily="Consolas"/></Border>
            </StackPanel>
            <TextBlock TextWrapping="Wrap" Foreground="#9090C0" LineHeight="20">
              <Run Text="Install or update selected modules using " Foreground="#9090C0"/>
              <Run Text="Smart Batch Routing" Foreground="#50C080" FontWeight="SemiBold"/>
              <Run Text=": modules with a PS7 version are updated via " Foreground="#9090C0"/>
              <Run Text="pwsh.exe" Foreground="#50D090" FontFamily="Consolas"/>
              <Run Text=", PS5-only modules via " Foreground="#9090C0"/>
              <Run Text="powershell.exe" Foreground="#F0A030" FontFamily="Consolas"/>
              <Run Text=". This ensures each module is installed to the correct engine path &#8212; fixing the old bug where PS7 modules were updated to the PS5 AllUsers path." Foreground="#9090C0"/>
            </TextBlock>
            <TextBlock Margin="0,8,0,0" TextWrapping="Wrap" Foreground="#505080" LineHeight="18">
              <Run Text="After install: " Foreground="#606090" FontWeight="SemiBold"/>
              <Run Text="old versions are automatically uninstalled and leftover folders deleted. Auto-refresh scan confirms the result." Foreground="#505080"/>
            </TextBlock>
          </StackPanel>
        </Border>

        <!-- MOVE MODULE -->
        <Border Background="#07071F" CornerRadius="6" Padding="16,12" Margin="0,6,0,0" BorderBrush="#181848" BorderThickness="1">
          <StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
              <TextBlock Text="MOVE MODULE  " Foreground="#F0A030" FontSize="14" FontWeight="Bold"/>
              <Border Background="#201800" CornerRadius="3" Padding="6,1"><TextBlock Text="Clean Scope Transfer" Foreground="#A06020" FontSize="11" FontFamily="Consolas"/></Border>
            </StackPanel>
            <TextBlock TextWrapping="Wrap" Foreground="#9090C0" LineHeight="20">
              <Run Text="Moves selected modules between " Foreground="#9090C0"/>
              <Run Text="CurrentUser" Foreground="#50C080" FontFamily="Consolas" FontWeight="SemiBold"/>
              <Run Text=" and " Foreground="#9090C0"/>
              <Run Text="AllUsers" Foreground="#5080E0" FontFamily="Consolas" FontWeight="SemiBold"/>
              <Run Text=" scope. Installs in the target scope, then " Foreground="#9090C0"/>
              <Run Text="uninstalls from the old scope AND deletes leftover folder remnants" Foreground="#F0A030" FontWeight="SemiBold"/>
              <Run Text=" &#8212; no stale files left behind." Foreground="#9090C0"/>
            </TextBlock>
          </StackPanel>
        </Border>

        <!-- SET PSMODULEPATH -->
        <Border Background="#07071F" CornerRadius="6" Padding="16,12" Margin="0,6,0,0" BorderBrush="#181848" BorderThickness="1">
          <StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
              <TextBlock Text="SET PSMODULEPATH  " Foreground="#A060F0" FontSize="14" FontWeight="Bold"/>
              <Border Background="#150A28" CornerRadius="3" Padding="6,1"><TextBlock Text="Persistent" Foreground="#7040C0" FontSize="11" FontFamily="Consolas"/></Border>
            </StackPanel>
            <TextBlock TextWrapping="Wrap" Foreground="#9090C0" LineHeight="20">
              <Run Text="Choose a folder to prepend to " Foreground="#9090C0"/>
              <Run Text="$env:PSModulePath" Foreground="#A060F0" FontFamily="Consolas"/>
              <Run Text=". Sets it for the current session AND persists it to the " Foreground="#9090C0"/>
              <Run Text="User" Foreground="#A060F0" FontWeight="SemiBold"/>
              <Run Text=" environment variable &#8212; new module installs will default to this location." Foreground="#9090C0"/>
            </TextBlock>
          </StackPanel>
        </Border>

        <!-- MODULE PATHS TAB -->
        <Border Background="#07071F" CornerRadius="6" Padding="16,12" Margin="0,6,0,0" BorderBrush="#181848" BorderThickness="1">
          <StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
              <TextBlock Text="MODULE PATHS TAB  " Foreground="#40C0C0" FontSize="14" FontWeight="Bold"/>
              <Border Background="#062020" CornerRadius="3" Padding="6,1"><TextBlock Text="Path Explorer" Foreground="#208080" FontSize="11" FontFamily="Consolas"/></Border>
            </StackPanel>
            <TextBlock TextWrapping="Wrap" Foreground="#9090C0" LineHeight="20">
              <Run Text="Shows all paths registered in " Foreground="#9090C0"/>
              <Run Text="$env:PSModulePath" Foreground="#40C0C0" FontFamily="Consolas"/>
              <Run Text=" as cards. Each card shows the module count, module names, and has " Foreground="#9090C0"/>
              <Run Text="Open in Explorer" Foreground="#40C0C0" FontWeight="SemiBold"/>
              <Run Text=" and " Foreground="#9090C0"/>
              <Run Text="Remove from PATH" Foreground="#E05050" FontWeight="SemiBold"/>
              <Run Text=" buttons. You can also add new paths and open all paths at once." Foreground="#9090C0"/>
            </TextBlock>
          </StackPanel>
        </Border>

        <!-- ENGINE INFO -->
        <Border Background="#07071F" CornerRadius="6" Padding="16,12" Margin="0,6,0,0" BorderBrush="#181848" BorderThickness="1">
          <StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
              <TextBlock Text="ENGINE INFO TAB  " Foreground="#F05050" FontSize="14" FontWeight="Bold"/>
              <Border Background="#200808" CornerRadius="3" Padding="6,1"><TextBlock Text="Updates via GitHub API" Foreground="#903030" FontSize="11" FontFamily="Consolas"/></Border>
            </StackPanel>
            <TextBlock TextWrapping="Wrap" Foreground="#9090C0" LineHeight="20">
              <Run Text="Checks " Foreground="#9090C0"/>
              <Run Text="github.com/PowerShell/PowerShell" Foreground="#F05050" FontFamily="Consolas"/>
              <Run Text=" for the latest PS7 release and " Foreground="#9090C0"/>
              <Run Text="github.com/microsoft/winget-cli" Foreground="#F05050" FontFamily="Consolas"/>
              <Run Text=" for winget. One-click install/update via winget. Shows installed vs latest version." Foreground="#9090C0"/>
            </TextBlock>
          </StackPanel>
        </Border>

        <!-- LOGGING -->
        <Border Background="#07071F" CornerRadius="6" Padding="16,12" Margin="0,6,0,0" BorderBrush="#181848" BorderThickness="1">
          <StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
              <TextBlock Text="VERBOSE LOGGING  " Foreground="#909090" FontSize="14" FontWeight="Bold"/>
              <Border Background="#121212" CornerRadius="3" Padding="6,1"><TextBlock Text="3-channel output" Foreground="#606060" FontSize="11" FontFamily="Consolas"/></Border>
            </StackPanel>
            <TextBlock TextWrapping="Wrap" Foreground="#9090C0" LineHeight="20">
              <Run Text="All events go to 3 places simultaneously: " Foreground="#9090C0"/>
              <Run Text="Operation Log" Foreground="#8080B0" FontWeight="SemiBold"/>
              <Run Text=" (high-level summary), " Foreground="#9090C0"/>
              <Run Text="Terminal Output" Foreground="#8080B0" FontWeight="SemiBold"/>
              <Run Text=" (verbose subprocess output including RAW install lines), and a " Foreground="#9090C0"/>
              <Run Text="log file" Foreground="#8080B0" FontWeight="SemiBold"/>
              <Run Text=" openable via 'Open in Notepad' in the Log Viewer tab. Tagged by: " Foreground="#9090C0"/>
              <Run Text="[SCAN] [RESULT] [BATCH] [ENGINE] [OK] [UP] [ERR]" Foreground="#505080" FontFamily="Consolas"/>
            </TextBlock>
          </StackPanel>
        </Border>

        <!-- OTHER FEATURES -->
        <Border Background="#07071F" CornerRadius="6" Padding="16,12" Margin="0,6,0,0" BorderBrush="#181848" BorderThickness="1">
          <StackPanel>
            <TextBlock Text="OTHER FEATURES" Foreground="#606090" FontSize="14" FontWeight="Bold" Margin="0,0,0,8"/>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <StackPanel Grid.Column="0">
                <TextBlock TextWrapping="Wrap" Foreground="#7070A0" LineHeight="22" FontSize="12">
                  <Run Text="&#x2022; " Foreground="#404070"/><Run Text="Category filters" Foreground="#9090C0" FontWeight="SemiBold"/><Run Text=" (Azure, Graph, Security...)&#x0A;" Foreground="#7070A0"/>
                  <Run Text="&#x2022; " Foreground="#404070"/><Run Text="Disable / Enable" Foreground="#9090C0" FontWeight="SemiBold"/><Run Text=" module (renames folder to .disabled)&#x0A;" Foreground="#7070A0"/>
                  <Run Text="&#x2022; " Foreground="#404070"/><Run Text="Export Module List" Foreground="#9090C0" FontWeight="SemiBold"/><Run Text=" to CSV / HTML / TXT&#x0A;" Foreground="#7070A0"/>
                  <Run Text="&#x2022; " Foreground="#404070"/><Run Text="Remove Selected" Foreground="#9090C0" FontWeight="SemiBold"/><Run Text=" &#8212; uninstall + delete files" Foreground="#7070A0"/>
                </TextBlock>
              </StackPanel>
              <StackPanel Grid.Column="1">
                <TextBlock TextWrapping="Wrap" Foreground="#7070A0" LineHeight="22" FontSize="12">
                  <Run Text="&#x2022; " Foreground="#404070"/><Run Text="Clean ALL Modules" Foreground="#9090C0" FontWeight="SemiBold"/><Run Text=" &#8212; full wipe with double confirm&#x0A;" Foreground="#7070A0"/>
                  <Run Text="&#x2022; " Foreground="#404070"/><Run Text="Clean Old Versions" Foreground="#9090C0" FontWeight="SemiBold"/><Run Text=" &#8212; scan and remove duplicate old versions&#x0A;" Foreground="#7070A0"/>
                  <Run Text="&#x2022; " Foreground="#404070"/><Run Text="Light / Dark theme" Foreground="#9090C0" FontWeight="SemiBold"/><Run Text=" + Auto (follows Windows)&#x0A;" Foreground="#7070A0"/>
                  <Run Text="&#x2022; " Foreground="#404070"/><Run Text="Filter / Search" Foreground="#9090C0" FontWeight="SemiBold"/><Run Text=" by name + Installed/Updates only&#x0A;" Foreground="#7070A0"/>
                  <Run Text="&#x2022; " Foreground="#404070"/><Run Text="Open Module Folder" Foreground="#9090C0" FontWeight="SemiBold"/><Run Text=" in Explorer&#x0A;" Foreground="#7070A0"/>
                  <Run Text="&#x2022; " Foreground="#404070"/><Run Text="Repositories tab" Foreground="#9090C0" FontWeight="SemiBold"/><Run Text=" &#8212; manage repos, add custom modules" Foreground="#7070A0"/>
                </TextBlock>
              </StackPanel>
            </Grid>
          </StackPanel>
        </Border>

        <TextBlock Text="github.com/karanik  |  karanik.gr" Foreground="#1E1E40"
                   FontSize="11" HorizontalAlignment="Center" Margin="0,16,0,8"/>

      </StackPanel>
    </ScrollViewer>

    <!-- Footer -->
    <Border Grid.Row="2" Background="#07071A" Padding="20,10" BorderBrush="#181838" BorderThickness="0,1,0,0">
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
        <Button Name="btnInfoClose" Content="Close" Width="100" Height="32"
                Background="#1A3CC0" Foreground="White" FontWeight="SemiBold"
                BorderThickness="0" Cursor="Hand"/>
      </StackPanel>
    </Border>
  </Grid>
</Window>
'@
        try {
            $infoReader = [System.Xml.XmlNodeReader]::new($infoXaml)
            $infoWin = [System.Windows.Markup.XamlReader]::Load($infoReader)
            $closeBtn = $infoWin.FindName('btnInfoClose')
            if ($closeBtn) { $closeBtn.add_Click({ $infoWin.Close() }) }
            $infoWin.ShowDialog() | Out-Null
        } catch {
            [System.Windows.MessageBox]::Show("Info window error: $_") | Out-Null
        }
    })

    $btnAbout.add_Click({
        try { Start-Process 'https://karanik.gr' } catch {}
        [System.Windows.MessageBox]::Show(
            "PowerShell Module Manager v7.0`n`nDeveloped by karanik.gr`n`nA GUI tool for managing PS modules across`nWindows PowerShell 5.1 and PowerShell 7.`n`nClick OK to open karanik.gr in browser.",
            'About', [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information) | Out-Null
    })


    $btnOpenFolder.add_Click({
        $sel = @($allRows | Where-Object { $_.Sel }) | Select-Object -First 1
        if (-not $sel) {
            [System.Windows.MessageBox]::Show('Select one module first.','Open Folder') | Out-Null
            return
        }

        # Try cached base path from last scan first
        $path = $rowIndex['__base__' + $sel.Name]

        # If not cached or path gone, do fresh lookup
        if (-not $path -or -not (Test-Path $path)) {
            $eng = GetEng
            $sc = (
                '$n="' + $sel.Name + '";$ep="SilentlyContinue";' +
                '$m=Get-Module -ListAvailable -Name $n -EA $ep 2>$null|Sort-Object Version -Desc|Select-Object -First 1;' +
                'if(-not $m){$m=Get-InstalledModule -Name $n -EA $ep 2>$null|Select-Object -First 1};' +
                'if($m){Write-Output $m.ModuleBase}'
            )
            $raw = & $eng.Exe -NoProfile -NonInteractive -Command $sc 2>$null
            # Filter to first non-empty path line
            $path = @($raw | Where-Object { $_.ToString().Trim() -and $_.ToString().Trim() -ne '' } | Select-Object -First 1)
            $path = if ($path) { $path[0].ToString().Trim() } else { '' }
        }

        # Path may be version subfolder - open parent (module root) if it looks like a version
        if ($path -and (Test-Path $path)) {
            $leaf = Split-Path $path -Leaf
            if ($leaf -match '^\d+\.\d+') {
                # It's a version folder - show module root (parent)
                $parentPath = Split-Path $path -Parent
                if (Test-Path $parentPath) { $path = $parentPath }
            }
            Start-Process 'explorer.exe' -ArgumentList $path
            ML "Opened folder: $path"
        } else {
            [System.Windows.MessageBox]::Show(
                "Module folder not found for: $($sel.Name)`n`nTip: Run Refresh Status first to populate paths.",
                'Open Folder') | Out-Null
        }
    })


    $btnMoveModule.add_Click({
        $sel = @($allRows | Where-Object { $_.Sel })
        if ($sel.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Select at least one module first.','Move Module') | Out-Null
            return
        }
        $eng = GetEng

        # Ask direction
        $ans = [System.Windows.MessageBox]::Show(
            "Move $($sel.Count) module(s) for: $($sel.Name -join ', ')`n`nYes  = CurrentUser  ->  AllUsers (System)`nNo   = AllUsers      ->  CurrentUser",
            'Move Module', [System.Windows.MessageBoxButton]::YesNoCancel,
            [System.Windows.MessageBoxImage]::Question)

        if ($ans -eq 'Cancel') { return }
        $toScope = if ($ans -eq 'Yes') { 'AllUsers' } else { 'CurrentUser' }
        $fromScope = if ($toScope -eq 'AllUsers') { 'CurrentUser' } else { 'AllUsers' }

        foreach ($mod in $sel) {
            ML "Moving $($mod.Name) from $fromScope to $toScope..."
            TL "MOVE $($mod.Name)  $fromScope -> $toScope" 'BATCH'
            $sc = (
                '$ep="SilentlyContinue";$n="' + $mod.Name + '";$s="' + $toScope + '";$sf="' + $fromScope + '";' +
                '$ErrorActionPreference="Stop";$ProgressPreference="SilentlyContinue";' +
                'try{' +
                '  # 1. Install in target scope' +
                '  Install-Module -Name $n -Scope $s -Force -AllowClobber -SkipPublisherCheck;' +
                '  $newMod=Get-InstalledModule $n -EA $ep|Sort-Object Version -Desc|Select-Object -First 1;' +
                '  $v=$newMod.Version;' +
                '  # 2. Get old scope module paths (all versions) before uninstall' +
                '  $oldMods=Get-InstalledModule -Name $n -AllVersions -EA $ep 2>$null|Where-Object{$_.InstalledLocation -notlike "*$s*"};' +
                '  $oldPaths=@($oldMods|Select-Object -ExpandProperty InstalledLocation);' +
                '  # 3. Uninstall from old scope via PSGet' +
                '  foreach($o in $oldMods){try{Uninstall-Module -Name $n -RequiredVersion $o.Version -Force -EA $ep}catch{}}; ' +
                '  # 4. Delete leftover folders (PSGet sometimes leaves files)' +
                '  foreach($p in $oldPaths){if($p -and (Test-Path $p)){try{Remove-Item $p -Recurse -Force -EA $ep}catch{}}}; ' +
                '  Write-Output ("OK|"+$v)' +
                '}catch{Write-Output ("ERR|"+$_.Exception.Message)}'
            )
            $out = & $eng.Exe -NoProfile -NonInteractive -Command $sc 2>&1
            $res = ($out | Where-Object { $_ -match '^\w+\|' } | Select-Object -First 1)
            if ($res -match '^OK\|(.+)$') {
                ML "[OK] Moved $($mod.Name) v$($Matches[1]) to $toScope (old scope cleaned)"
                TL "OK   $($mod.Name) v$($Matches[1]) moved + old scope cleaned" 'OK'
                if ($toScope -eq 'AllUsers') { $mod.S5 = 'AllUsers'; $mod.S7 = 'AllUsers' }
                else { $mod.S5 = 'CurrentUser'; $mod.S7 = 'CurrentUser' }
            } else {
                ML "[!!] Move failed for $($mod.Name): $res"
                TL "FAIL $($mod.Name): $res" 'ERR'
            }
        }
        ML "Move operation complete."
        TL "=== Move operation complete ==="
    })

    # -------------------------------------------------------------------------
    # SET PSMODULEPATH button
    # -------------------------------------------------------------------------
    $btnSetPath.add_Click({
        $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
        $dlg.Description    = "Select default module install folder"
        $dlg.ShowNewFolderButton = $true
        # Pre-select first path in PSModulePath
        $first = ($env:PSModulePath -split ';' | Where-Object { Test-Path $_ } | Select-Object -First 1)
        if ($first) { $dlg.SelectedPath = $first }
        $dlg.RootFolder = [System.Environment+SpecialFolder]::Desktop

        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $chosen = $dlg.SelectedPath
            $ans = [System.Windows.MessageBox]::Show(
                "Set as default module path:`n$chosen`n`nThis will prepend the folder to `$env:PSModulePath for the current session AND set it permanently in the user environment.`n`nContinue?",
                'Set PSModulePath', [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question)
            if ($ans -eq 'Yes') {
                # Add to front of PSModulePath (session)
                $existing = $env:PSModulePath -split ';' | Where-Object { $_ -ne $chosen }
                $newPath  = ($chosen + ';' + ($existing -join ';')).Trim(';')
                $env:PSModulePath = $newPath
                # Persist to user environment
                [System.Environment]::SetEnvironmentVariable('PSModulePath', $newPath, 'User')
                ML "[OK] PSModulePath updated. New default: $chosen"
                TL "PSModulePath set: $chosen" 'OK'
                [System.Windows.MessageBox]::Show(
                    "PSModulePath updated!`n`nNew path:`n$chosen`n`nNew installs will default here.`nRestart PS to apply globally.",
                    'Done', [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information) | Out-Null
            }
        }
    })

    # -------------------------------------------------------------------------
    # MODULE PATHS tab functions
    # -------------------------------------------------------------------------
    function RefreshPathsTab {
        $pnlPaths.Children.Clear()
        $allPaths = $env:PSModulePath -split ';' | Where-Object { $_ -ne '' }

        foreach ($p in $allPaths) {
            $exists = Test-Path $p
            $count  = if ($exists) { @(Get-ChildItem $p -Directory -EA SilentlyContinue).Count } else { 0 }
            # Module names list
            $mods = if ($exists) {
                @(Get-ChildItem $p -Directory -EA SilentlyContinue | Select-Object -ExpandProperty Name) -join ', '
            } else { '' }
            if ($mods.Length -gt 120) { $mods = $mods.Substring(0, 117) + '...' }

            $bg    = if ($exists) { '#07071F' } else { '#1F0707' }
            $fg    = if ($exists) { '#3030A0' } else { '#803030' }
            $badge = if ($exists) { "[$count modules]" } else { '[NOT FOUND]' }

            $card = [System.Windows.Controls.Border]::new()
            $card.Background     = [System.Windows.Media.BrushConverter]::new().ConvertFromString($bg)
            $card.BorderBrush    = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#181848')
            $card.BorderThickness = [System.Windows.Thickness]::new(1)
            $card.CornerRadius   = [System.Windows.CornerRadius]::new(4)
            $card.Margin         = [System.Windows.Thickness]::new(0,0,0,6)
            $card.Padding        = [System.Windows.Thickness]::new(10,8,10,8)
            $card.Tag            = $p

            $grid = [System.Windows.Controls.Grid]::new()
            $c0   = [System.Windows.Controls.ColumnDefinition]::new(); $c0.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $c1   = [System.Windows.Controls.ColumnDefinition]::new(); $c1.Width = [System.Windows.GridLength]::Auto
            $grid.ColumnDefinitions.Add($c0); $grid.ColumnDefinitions.Add($c1)

            $left = [System.Windows.Controls.StackPanel]::new()
            $left.Orientation = [System.Windows.Controls.Orientation]::Vertical

            # Path header
            $hdr = [System.Windows.Controls.TextBlock]::new()
            $hdr.Text       = "$badge  $p"
            $hdr.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fg)
            $hdr.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
            $hdr.FontSize   = 12
            $hdr.FontWeight = [System.Windows.FontWeights]::SemiBold
            $left.Children.Add($hdr)

            # Modules line
            if ($mods) {
                $ml = [System.Windows.Controls.TextBlock]::new()
                $ml.Text       = $mods
                $ml.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1E1E5A')
                $ml.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
                $ml.FontSize   = 10
                $ml.Margin     = [System.Windows.Thickness]::new(0,3,0,0)
                $ml.TextWrapping = [System.Windows.TextWrapping]::Wrap
                $left.Children.Add($ml)
            }

            [System.Windows.Controls.Grid]::SetColumn($left, 0)
            $grid.Children.Add($left)

            # Right: buttons
            $btnSP = [System.Windows.Controls.StackPanel]::new()
            $btnSP.Orientation  = [System.Windows.Controls.Orientation]::Horizontal
            $btnSP.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
            [System.Windows.Controls.Grid]::SetColumn($btnSP, 1)

            if ($exists) {
                $bOpen = [System.Windows.Controls.Button]::new()
                $bOpen.Content = "Open"
                $bOpen.Padding = [System.Windows.Thickness]::new(10, 3, 10, 3)
                $bOpen.Margin  = [System.Windows.Thickness]::new(4, 0, 0, 0)
                $bOpen.FontSize = 11
                $bOpen.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#0E1A3A')
                $bOpen.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#4060D0')
                $bOpen.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#181858')
                $pathCapture = $p
                $bOpen.add_Click({ Start-Process 'explorer.exe' $pathCapture }.GetNewClosure())
                $btnSP.Children.Add($bOpen)
            }

            $bRem = [System.Windows.Controls.Button]::new()
            $bRem.Content = "Remove from PATH"
            $bRem.Padding = [System.Windows.Thickness]::new(10, 3, 10, 3)
            $bRem.Margin  = [System.Windows.Thickness]::new(4, 0, 0, 0)
            $bRem.FontSize = 11
            $bRem.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#200808')
            $bRem.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#A03030')
            $bRem.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#401010')
            $pathCapture = $p
            $bRem.add_Click({
                $ans2 = [System.Windows.MessageBox]::Show(
                    "Remove from PSModulePath?`n$pathCapture`n`nThis removes it from the user environment variable.",
                    'Remove Path', [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Warning)
                if ($ans2 -eq 'Yes') {
                    $cur = $env:PSModulePath -split ';' | Where-Object { $_ -ne $pathCapture }
                    $env:PSModulePath = $cur -join ';'
                    [System.Environment]::SetEnvironmentVariable('PSModulePath', ($cur -join ';'), 'User')
                    ML "[OK] Removed from PSModulePath: $pathCapture"
                    RefreshPathsTab
                }
            }.GetNewClosure())
            $btnSP.Children.Add($bRem)

            $grid.Children.Add($btnSP)
            $card.Child = $grid
            $pnlPaths.Children.Add($card)
        }

        $lblPathInfo.Text = "Showing $($allPaths.Count) paths from `$env:PSModulePath  |  Click 'Open' to browse in Explorer"
    }

    $btnRefreshPaths.add_Click({ RefreshPathsTab })

    $btnOpenPathExplorer.add_Click({
        $allPaths = $env:PSModulePath -split ';' | Where-Object { Test-Path $_ }
        foreach ($p in $allPaths) { Start-Process 'explorer.exe' $p }
    })

    $btnAddToPath.add_Click({
        $dlg2 = [System.Windows.Forms.FolderBrowserDialog]::new()
        $dlg2.Description = "Select folder to add to PSModulePath"
        $dlg2.ShowNewFolderButton = $true
        if ($dlg2.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $chosen2 = $dlg2.SelectedPath
            if (($env:PSModulePath -split ';') -notcontains $chosen2) {
                $env:PSModulePath = $chosen2 + ';' + $env:PSModulePath
                [System.Environment]::SetEnvironmentVariable('PSModulePath', $env:PSModulePath, 'User')
                ML "[OK] Added to PSModulePath: $chosen2"
                RefreshPathsTab
            } else {
                [System.Windows.MessageBox]::Show("Already in PSModulePath: $chosen2") | Out-Null
            }
        }
    })

    $btnRemFromPath.add_Click({ RefreshPathsTab })  # refresh shows current state


    $btnDisable.add_Click({
        $sel = @($allRows | Where-Object { $_.Sel })
        if ($sel.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Select at least one module.','Disable') | Out-Null; return
        }
        foreach ($mod in $sel) {
            try {
                # Rename module folder to .disabled to prevent autoloading
                $eng = GetEng
                $sc = '$n="' + $mod.Name + '";$m=Get-Module -ListAvailable -Name $n|Select-Object -First 1;if($m){Write-Output $m.ModuleBase}'
                $base = (& $eng.Exe -NoProfile -NonInteractive -Command $sc 2>$null | Select-Object -First 1)
                if ($base -and (Test-Path $base)) {
                    $parent = Split-Path $base -Parent
                    $leaf   = Split-Path $base -Leaf
                    $newPath = Join-Path $parent ($leaf + '.disabled')
                    Rename-Item -Path $base -NewName ($leaf + '.disabled') -Force
                    $mod.St = 'Disabled'; $mod.V5 = ''; $mod.V7 = ''
                    ML "[DISABLED] $($mod.Name) at $base"
                } else {
                    ML "[!!] Could not find path for $($mod.Name)"
                }
            } catch { ML "[!!] Disable error $($mod.Name): $_" }
        }
    })

    $btnEnable.add_Click({
        $sel = @($allRows | Where-Object { $_.Sel })
        if ($sel.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Select at least one module.','Enable') | Out-Null; return
        }
        foreach ($mod in $sel) {
            try {
                $eng = GetEng
                # Look for .disabled folder
                $sc = (
                    '$n="' + $mod.Name + '";' +
                    '$paths=@("$env:ProgramFiles\PowerShell\Modules","$env:ProgramFiles\WindowsPowerShell\Modules","$HOME\Documents\PowerShell\Modules","$HOME\Documents\WindowsPowerShell\Modules");' +
                    'foreach($p in $paths){$d=Get-ChildItem $p -Filter ($n+".disabled") -EA SilentlyContinue|Select-Object -First 1;if($d){Write-Output $d.FullName;break}}'
                )
                $disPath = (& $eng.Exe -NoProfile -NonInteractive -Command $sc 2>$null | Select-Object -First 1)
                if ($disPath -and (Test-Path $disPath)) {
                    $newPath = $disPath -replace '\.disabled$',''
                    Rename-Item -Path $disPath -NewName (Split-Path $newPath -Leaf) -Force
                    $mod.St = 'Pending'
                    ML "[ENABLED] $($mod.Name) - click Refresh Status to verify"
                } else {
                    ML "[!!] No .disabled folder found for $($mod.Name)"
                }
            } catch { ML "[!!] Enable error $($mod.Name): $_" }
        }
    })


    $btnRemove.add_Click({
        $sel = @($allRows | Where-Object { $_.Sel })
        if ($sel.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Select at least one module to remove.','Remove') | Out-Null
            return
        }
        $names = ($sel | ForEach-Object { $_.Name }) -join ", "
        $ans = [System.Windows.MessageBox]::Show(
            "REMOVE $($sel.Count) module(s)?`n`n$names`n`nThis will Uninstall-Module AND delete files.",
            'Confirm Remove', [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
        if ($ans -ne 'Yes') { return }

        $eng = GetEng
        foreach ($mod in $sel) {
            ML "Removing $($mod.Name)..."
            $sc = (
                '$n="' + $mod.Name + '";$ErrorActionPreference="SilentlyContinue";' +
                # Step 1: Uninstall via PowerShellGet
                'try{$all=Get-InstalledModule -Name $n -AllVersions -EA SilentlyContinue;' +
                'if($all){$all|ForEach-Object{Uninstall-Module -Name $n -RequiredVersion $_.Version -Force -EA SilentlyContinue};Write-Output "UNINSTALLED"}' +
                'else{Write-Output "NOT_INST"}}catch{Write-Output ("ERR1|"+$_)};' +
                # Step 2: Delete leftover files from all known paths
                '$paths=@(' +
                '"$env:ProgramFiles\PowerShell\Modules"+$n,' +
                '"$env:ProgramFiles\WindowsPowerShell\Modules"+$n,' +
                '"$HOME\Documents\PowerShell\Modules"+$n,' +
                '"$HOME\Documents\WindowsPowerShell\Modules"+$n,' +
                '"$env:ProgramFiles(x86)\WindowsPowerShell\Modules"+$n);' +
                'foreach($p in $paths){if(Test-Path $p){Remove-Item $p -Recurse -Force -EA SilentlyContinue;Write-Output ("DELETED|"+$p)}}'
            )
            $out = & $eng.Exe -NoProfile -NonInteractive -Command $sc 2>&1
            foreach ($ln in $out) {
                $x = $ln.ToString().Trim()
                if ($x) { ML "  $x" }
            }
            $mod.V5 = ''; $mod.S5 = ''; $mod.V7 = ''; $mod.S7 = ''
            $mod.St = 'Not Installed'; $mod.Up = $false; $mod.Sel = $false
            ML "[REMOVED] $($mod.Name)"
        }
        DoFilter
        ML "Remove operation complete."
    })


    $btnExport.add_Click({
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Title      = 'Export Module List'
        $dlg.Filter     = 'CSV (*.csv)|*.csv|HTML (*.html)|*.html|Text (*.txt)|*.txt'
        $dlg.FileName   = "PSModules_$(Get-Date -Format 'yyyyMMdd_HHmm')"
        if ($dlg.ShowDialog() -ne 'OK') { return }

        $path = $dlg.FileName
        $ext  = [System.IO.Path]::GetExtension($path).ToLower()

        $data = $allRows | Where-Object { $_.V5 -or $_.V7 } | Sort-Object Cat, Name

        try {
            switch ($ext) {
                '.csv' {
                    $lines = @('"Name","Category","Description","PS5_Ver","PS5_Scope","PS7_Ver","PS7_Scope","Gallery","Status"')
                    foreach ($r in $data) {
                        $lines += '"'+$r.Name+'","'+$r.Cat+'","'+$r.Desc+'","'+$r.V5+'","'+$r.S5+'","'+$r.V7+'","'+$r.S7+'","'+$r.GV+'","'+$r.St+'"'
                    }
                    $lines | Set-Content $path -Encoding UTF8
                }
                '.html' {
                    $rows = ($data | ForEach-Object {
                        $color = if ($_.Up) { '#3A2800' } elseif ($_.St -eq 'Not Installed') { '#1A1A1A' } else { '#0A1A0A' }
                        "<tr style='background:$color'><td>$($_.Name)</td><td>$($_.Cat)</td><td>$($_.Desc)</td><td>$($_.V5)</td><td>$($_.S5)</td><td>$($_.V7)</td><td>$($_.S7)</td><td>$($_.GV)</td><td>$($_.St)</td></tr>"
                    }) -join "`n"
                    $html = @"
<!DOCTYPE html><html><head><meta charset='utf-8'>
<title>PS Module List - $(Get-Date -Format 'yyyy-MM-dd')</title>
<style>body{font-family:Segoe UI,sans-serif;background:#0F0F1A;color:#C0C0E0}
table{border-collapse:collapse;width:100%}th{background:#07071A;color:#5A88FF;padding:8px;text-align:left;border-bottom:2px solid #1A1A3A}
td{padding:5px 8px;border-bottom:1px solid #181830;font-size:12px}</style></head>
<body><h2 style='color:#5A88FF'>PowerShell Module Manager - Export</h2>
<p style='color:#606080'>Generated: $(Get-Date)  |  karanik.gr</p>
<table><tr><th>Name</th><th>Category</th><th>Description</th><th>PS5 Ver</th><th>PS5 Scope</th><th>PS7 Ver</th><th>PS7 Scope</th><th>Gallery</th><th>Status</th></tr>
$rows</table></body></html>
"@
                    $html | Set-Content $path -Encoding UTF8
                }
                default {
                    $lines = @("PowerShell Module Manager Export - $(Get-Date)", "=" * 80)
                    foreach ($r in $data) {
                        $lines += "$($r.Name.PadRight(45)) Cat:$($r.Cat.PadRight(12)) PS5:$($r.V5.PadRight(12)) PS7:$($r.V7.PadRight(12)) Status:$($r.St)"
                    }
                    $lines | Set-Content $path -Encoding UTF8
                }
            }
            ML "Exported $($data.Count) modules to: $path"
            $ans = [System.Windows.MessageBox]::Show("Exported $($data.Count) modules.`n`nOpen file now?", 'Export Done',
                [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
            if ($ans -eq 'Yes') { Start-Process $path }
        } catch {
            ML "[!!] Export error: $_"
            [System.Windows.MessageBox]::Show("Export error: $_", 'Error') | Out-Null
        }
    })


    $btnCleanAll.add_Click({
        $ans = [System.Windows.MessageBox]::Show(
            "WARNING: This will REMOVE ALL modules from ALL scopes!`n`n" +
            "- Uninstall-Module for every installed module`n" +
            "- Delete files from all PS module paths`n" +
            "- Affects both AllUsers and CurrentUser`n`n" +
            "This cannot be undone. Are you absolutely sure?",
            'CLEAN ALL MODULES',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning) | Out-Null
        if ($ans -ne 'Yes') { return }

        # Double-confirm
        $ans2 = [System.Windows.MessageBox]::Show(
            "FINAL CONFIRMATION`n`nRemove ALL PowerShell modules?",
            'Confirm', [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Stop)
        if ($ans2 -ne 'Yes') { return }

        $eng = GetEng
        ML "=== CLEAN ALL MODULES STARTED ===" 'WARN'

        # Get list of all installed modules
        $sc = (
            '$ep="SilentlyContinue";$ErrorActionPreference=$ep;$ProgressPreference=$ep;' +
            '$all = Get-InstalledModule -EA $ep 2>$null;' +
            'if(-not $all){$all = Get-Module -ListAvailable -EA $ep 2>$null};' +
            'foreach($m in $all){ Write-Output ($m.Name+"|"+$m.ModuleBase) }'
        )
        $installed = & $eng.Exe -NoProfile -NonInteractive -Command $sc 2>$null

        $total = @($installed).Count
        $done  = 0
        $pbar.Visibility = 'Visible'; $pbar.Value = 0

        foreach ($ln in $installed) {
            $parts = $ln.ToString().Split('|', 2)
            $mName = $parts[0].Trim()
            $mBase = if ($parts.Count -ge 2) { $parts[1].Trim() } else { '' }
            if (-not $mName) { continue }

            ML "Removing: $mName"

            # Step 1: Uninstall via PSGet (all versions)
            $usc = (
                '$n="' + ($mName -replace '"','') + '";$ep="SilentlyContinue";' +
                '$all=Get-InstalledModule -Name $n -AllVersions -EA $ep 2>$null;' +
                'if($all){$all|ForEach-Object{Uninstall-Module -Name $n -RequiredVersion $_.Version -Force -EA $ep}}; ' +
                'Write-Output "DONE"'
            )
            & $eng.Exe -NoProfile -NonInteractive -Command $usc 2>$null | Out-Null

            # Step 2: Delete leftover folders from ALL known PS module paths
            $allModPaths = @(
                "$env:ProgramFiles\PowerShell\Modules\$mName",
                "$env:ProgramFiles\WindowsPowerShell\Modules\$mName",
                "$env:ProgramFiles(x86)\WindowsPowerShell\Modules\$mName",
                "$env:USERPROFILE\Documents\PowerShell\Modules\$mName",
                "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\$mName",
                "$env:APPDATA\Local\Microsoft\PowerShell\Modules\$mName"
            )
            foreach ($mp in $allModPaths) {
                if (Test-Path $mp) {
                    try { Remove-Item $mp -Recurse -Force -EA SilentlyContinue; ML "  Deleted: $mp" }
                    catch { ML "  Error deleting $mp : $_" }
                }
            }

            $done++
            $pbar.Value = [int](($done / [math]::Max($total,1)) * 100)
        }

        # Also clean up any remaining folders not caught by module list
        $cleanPaths = @(
            "$env:ProgramFiles\PowerShell\Modules",
            "$env:ProgramFiles\WindowsPowerShell\Modules",
            "$env:USERPROFILE\Documents\PowerShell\Modules",
            "$env:USERPROFILE\Documents\WindowsPowerShell\Modules"
        )
        foreach ($cp in $cleanPaths) {
            if (Test-Path $cp) {
                $orphans = Get-ChildItem $cp -Directory -EA SilentlyContinue | Where-Object {
                    $n = $_.Name -replace '\.disabled$',''
                    -not (Get-InstalledModule -Name $n -EA SilentlyContinue) -and
                    # Keep system modules (they start with Microsoft. or are built-in)
                    $n -notmatch '^(Microsoft\.PowerShell|PSReadLine|PowerShellGet|PackageManagement)$'
                }
            }
        }

        # Reset all rows
        foreach ($r in $allRows) {
            $r.V5=''; $r.S5=''; $r.V7=''; $r.S7=''; $r.GV=''; $r.St='Not Installed'; $r.Up=$false; $r.Sel=$false
        }

        $pbar.Visibility = 'Collapsed'
        DoFilter; ReloadLogView
        ML "=== CLEAN ALL DONE: $done modules removed ===" 'WARN'
        [System.Windows.MessageBox]::Show("Removed $done modules.`n`nRun Refresh Status to verify.", 'Clean Done') | Out-Null
    })


    # -------------------------------------------------------------------------
    # CLEAN OLD VERSIONS
    # -------------------------------------------------------------------------
    $btnCleanOld.add_Click({
        $eng  = GetEng
        if (-not $eng.OK) {
            [System.Windows.MessageBox]::Show('No engine available.','Clean Old Versions') | Out-Null; return
        }
        ML "Scanning for old module versions..."
        $lblStatus.Text = "Scanning for duplicate versions..."
        $pbar.Visibility = 'Visible'; $pbar.Value = 0

        $script:OldVerRS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:OldVerRS.ApartmentState = 'MTA'; $script:OldVerRS.ThreadOptions = 'ReuseThread'
        $script:OldVerRS.Open()
        $script:OldVerPS = [System.Management.Automation.PowerShell]::Create()
        $script:OldVerPS.Runspace = $script:OldVerRS

        [void]$script:OldVerPS.AddScript({
            param($termQ, $logF, $ex5, $ex7)
            $ep = 'SilentlyContinue'
            function Push([string]$m) {
                $ts = Get-Date -f 'HH:mm:ss.fff'
                try { Add-Content -Path $logF -Value "[$ts][OLDVER] $m" -Encoding UTF8 } catch {}
                try { $termQ.Enqueue("[OLDVER] $m") } catch {}
            }
            function ScanOldVers([string]$exe, [string]$tag) {
                if (-not $exe) { return }
                Push "[$tag] Scanning for duplicate module versions via $exe..."
                $cmd = (
                    '$ep="SilentlyContinue";$ErrorActionPreference=$ep;$ProgressPreference=$ep;' +
                    '$all=Get-InstalledModule -AllVersions -EA $ep 2>$null;' +
                    'if(-not $all){Write-Output "NONE"; return};' +
                    '$groups=$all|Group-Object Name;' +
                    'foreach($g in $groups){' +
                    '  if($g.Count -lt 2){continue};' +
                    '  $sorted=$g.Group|Sort-Object Version -Desc;' +
                    '  $latest=$sorted[0];' +
                    '  $old=$sorted|Select-Object -Skip 1;' +
                    '  foreach($o in $old){' +
                    '    Write-Output ($o.Name+"|"+$o.Version.ToString()+"|"+$latest.Version.ToString()+"|"+$o.InstalledLocation)' +
                    '  }' +
                    '}'
                )
                try {
                    $out = & $exe -NoProfile -NonInteractive -Command $cmd 2>$null
                    $count = 0
                    foreach ($ln in $out) {
                        $x = $ln.ToString().Trim()
                        if ($x -and $x -ne 'NONE') {
                            $termQ.Enqueue("OLDVER_ITEM|$tag|$x")
                            $count++
                        }
                    }
                    Push "[$tag] Found $count old version(s)."
                } catch { Push "[$tag] Scan error: $_" }
            }
            ScanOldVers $ex5 'PS5'
            ScanOldVers $ex7 'PS7'
            $termQ.Enqueue("OLDVER_SCAN_DONE")
        })
        $e5obj = $engines | Where-Object { $_.Tag -eq 'PS5' -and $_.OK } | Select-Object -First 1
        $e7obj = $engines | Where-Object { $_.Tag -eq 'PS7' -and $_.OK } | Select-Object -First 1
        [void]$script:OldVerPS.AddParameters(@{
            termQ = $termQ; logF = $logFile
            ex5   = if ($e5obj) { $e5obj.Exe } else { '' }
            ex7   = if ($e7obj) { $e7obj.Exe } else { '' }
        })
        $script:OldVerAR = $script:OldVerPS.BeginInvoke()

        # Collect results
        $script:OldItems = [System.Collections.Generic.List[object]]::new()

        $script:OldVerTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:OldVerTimer.Interval = [System.TimeSpan]::FromMilliseconds(300)
        $script:OldVerTimer.add_Tick({
            $tLine = $null
            while ($termQ.TryDequeue([ref]$tLine)) {
                if ($tLine -match '^OLDVER_ITEM\|([^|]+)\|(.+)$') {
                    $tag  = $Matches[1]
                    $data = $Matches[2].Split('|', 4)
                    if ($data.Count -ge 3) {
                        $script:OldItems.Add([PSCustomObject]@{
                            Eng=$tag; Name=$data[0]; OldVer=$data[1]
                            LatestVer=$data[2]; Path=if($data.Count -ge 4){$data[3]}else{''}
                        })
                        TL "  OLD: $($data[0]) v$($data[1]) (latest: v$($data[2])) [$tag]" 'OLDVER'
                    }
                } elseif ($tLine -match '^\[OLDVER\] (.+)$') {
                    TL $Matches[1] 'OLDVER'
                }
            }

            if ($script:OldVerAR -ne $null -and $script:OldVerAR.IsCompleted) {
                $script:OldVerTimer.Stop()
                try { $script:OldVerPS.EndInvoke($script:OldVerAR) | Out-Null
                      $script:OldVerPS.Dispose(); $script:OldVerRS.Close(); $script:OldVerRS.Dispose() } catch {}

                $pbar.Visibility = 'Collapsed'
                $items = $script:OldItems

                if ($items.Count -eq 0) {
                    $lblStatus.Text = "Clean Old Versions: No old versions found."
                    ML "[OK] No old module versions found - everything is clean!"
                    [System.Windows.MessageBox]::Show(
                        "No old versions found!`n`nAll installed modules are already on their latest version.",
                        'Clean Old Versions', [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information) | Out-Null
                    return
                }

                # Show summary dialog
                $summary = ($items | ForEach-Object {
                    "  $($_.Name.PadRight(40)) v$($_.OldVer.PadRight(12)) -> keep v$($_.LatestVer)  [$($_.Eng)]"
                }) -join "`n"

                $ans = [System.Windows.MessageBox]::Show(
                    "Found $($items.Count) old module version(s) to remove:`n`n$summary`n`nRemove all old versions now?",
                    'Clean Old Versions',
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Question)

                if ($ans -ne 'Yes') {
                    $lblStatus.Text = "Clean Old Versions: Cancelled."
                    return
                }

                # Remove old versions
                $eng = GetEng
                $removed = 0; $errors = 0
                $pbar.Visibility = 'Visible'; $pbar.Value = 0
                $total = $items.Count

                foreach ($item in $items) {
                    $exe = if ($item.Eng -eq 'PS7') {
                        ($engines | Where-Object { $_.Tag -eq 'PS7' -and $_.OK } | Select-Object -First 1).Exe
                    } else {
                        ($engines | Where-Object { $_.Tag -eq 'PS5' -and $_.OK } | Select-Object -First 1).Exe
                    }
                    if (-not $exe) { $exe = $eng.Exe }

                    $sc = (
                        '$ep="SilentlyContinue";$n="' + ($item.Name -replace '"','') + '";$v="' + $item.OldVer + '";' +
                        '$ErrorActionPreference="Stop";$ProgressPreference="SilentlyContinue";' +
                        'try{' +
                        '  Uninstall-Module -Name $n -RequiredVersion $v -Force -EA $ep;' +
                        '  $p=(Get-InstalledModule $n -RequiredVersion $v -EA $ep 2>$null);' +
                        '  if(-not $p){Write-Output "OK"} else{Write-Output "STILL_THERE"}' +
                        '}catch{Write-Output ("ERR|"+$_.Exception.Message)}'
                    )
                    try {
                        $out = & $exe -NoProfile -NonInteractive -Command $sc 2>&1
                        $res = ($out | Where-Object { $_ -match '^(OK|ERR|STILL)' } | Select-Object -First 1)
                        if ($res -match '^OK') {
                            # Also try to delete folder if it remains
                            if ($item.Path -and (Test-Path $item.Path)) {
                                try { Remove-Item $item.Path -Recurse -Force -EA SilentlyContinue } catch {}
                            }
                            ML "[OK] Removed $($item.Name) v$($item.OldVer) (kept v$($item.LatestVer))"
                            TL "OK  $($item.Name) v$($item.OldVer) removed" 'OK'
                            $removed++
                        } else {
                            ML "[!!] Could not remove $($item.Name) v$($item.OldVer): $res"
                            $errors++
                        }
                    } catch {
                        ML "[!!] Error removing $($item.Name) v$($item.OldVer): $_"; $errors++
                    }
                    $pbar.Value = [int](($removed + $errors) / $total * 100)
                }

                $pbar.Visibility = 'Collapsed'
                $msg = "Clean Old Versions complete: $removed removed"
                if ($errors -gt 0) { $msg += ", $errors errors (may need admin)" }
                $lblStatus.Text = $msg
                ML "=== CLEAN OLD VERSIONS DONE: $removed/$total removed ==="
                [System.Windows.MessageBox]::Show(
                    "$msg`n`nRun Refresh Status to confirm.",
                    'Clean Old Versions', [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information) | Out-Null
                DoFilter
            }
        })
        $script:OldVerTimer.Start()
    })


    # -------------------------------------------------------------------------
    # REPOSITORIES TAB
    # -------------------------------------------------------------------------
    # Known/popular repos that can be toggled
    # =========================================================================
    # REPOSITORIES TAB  - stateless architecture (no stale closure issues)
    # =========================================================================
    $script:KnownRepos = @(
        [PSCustomObject]@{
            Name='PSGallery'; CanRemove=$false
            Url='https://www.powershellgallery.com/api/v2'
            DefaultTrust='Trusted'
            Color='#0A180A'; BorderColor='#0A2F0A'
            BadgeColor='#30B050'; DescColor='#204020'
            Desc='Official Microsoft PowerShell Gallery - the primary source for PS modules. Maintained and scanned by Microsoft. Cannot be unregistered.'
        }
        [PSCustomObject]@{
            Name='NuGet'; CanRemove=$true
            Url='https://www.nuget.org/api/v2'
            DefaultTrust='Untrusted'
            Color='#07071A'; BorderColor='#18183A'
            BadgeColor='#5060D0'; DescColor='#202048'
            Desc='.NET package repository - primarily .NET libraries, not PS modules. Use with caution for PS workflows.'
        }
        [PSCustomObject]@{
            Name='Chocolatey'; CanRemove=$true
            Url='https://chocolatey.org/api/v2'
            DefaultTrust='Untrusted'
            Color='#120A00'; BorderColor='#2A1800'
            BadgeColor='#B06020'; DescColor='#2A1800'
            Desc='Windows package manager for applications - NOT for PowerShell modules. Non-standard as a PSRepository.'
        }
    )

    # State dict: tracks registration status of known repos (populated by RefreshReposTab)
    $script:RepoRegState   = @{}   # Name -> $true/$false
    $script:RepoTrustState = @{}   # Name -> 'Trusted'/'Untrusted'
    $script:RepoTabBusy    = $false

    function Get-RepoEngine {
        $e7 = $engines | Where-Object { $_.Tag -eq 'PS7' -and $_.OK } | Select-Object -First 1
        if ($e7) { return $e7 }
        return $engines | Where-Object { $_.OK } | Select-Object -First 1
    }

    # Run repo operation async - fires callback on completion via DispatcherTimer
    function Invoke-RepoOp([string]$PsCmd, [scriptblock]$OnDone) {
        $eng = Get-RepoEngine
        if (-not $eng) { & $OnDone 'ERR|No engine'; return }
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = 'MTA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs
        $q  = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        $sb = {
            param($exe, $cmd, $q)
            try {
                $out = & $exe -NoProfile -NonInteractive -Command $cmd
                $last = ($out | Select-Object -Last 1)
                $val = if ($last) { $last.ToString() } else { 'OK' }
                $q.Enqueue($val)
            } catch { $q.Enqueue("ERR|$_") }
        }
        [void]$ps.AddScript($sb)
        [void]$ps.AddParameters(@{ exe = $eng.Exe; cmd = $PsCmd; q = $q })
        $ar = $ps.BeginInvoke()
        $exeCapture = $eng.Exe
        $t = [System.Windows.Threading.DispatcherTimer]::new()
        $t.Interval = [System.TimeSpan]::FromMilliseconds(250)
        $cbCapture = $OnDone; $qCapture = $q
        $t.add_Tick({
            if ($ar.IsCompleted) {
                $t.Stop()
                try { $ps.EndInvoke($ar) | Out-Null; $ps.Dispose(); $rs.Close(); $rs.Dispose() } catch {}
                $res = $null; $qCapture.TryDequeue([ref]$res) | Out-Null
                & $cbCapture ($res)
            }
        }.GetNewClosure())
        $t.Start()
    }

    function RefreshReposTab {
        if ($script:RepoTabBusy) { return }
        $script:RepoTabBusy = $true
        $pnlRepos.Children.Clear()
        $pnlCustomMods.Children.Clear()
        $lblRepoStatus.Text = "Loading repositories..."

        $eng = Get-RepoEngine
        if (-not $eng) {
            $lblRepoStatus.Text = "No engine available."
            $script:RepoTabBusy = $false; return
        }

        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = 'MTA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs
        $q  = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        $sb2 = {
            param($exe, $q)
            $cmd = (
                '$ep="SilentlyContinue";$ErrorActionPreference=$ep;$ProgressPreference=$ep;' +
                'if(-not(Get-Command Get-PSRepository -EA $ep)){$q2="NOMOD";Write-Output $q2;return};' +
                'try{$repos=Get-PSRepository -EA $ep 2>$null}catch{$repos=@()};' +
                'if(-not $repos){Write-Output "NONE";return};' +
                'foreach($r in $repos){' +
                '  Write-Output ("REPO^"+$r.Name+"^"+$r.SourceLocation+"^"+$r.InstallationPolicy+"^"+$r.PackageManagementProvider)' +
                '}'
            )
            try {
                $out = & $exe -NoProfile -NonInteractive -Command $cmd 2>$null
                foreach ($ln in $out) { $q.Enqueue($ln.ToString().Trim()) }
            } catch { $q.Enqueue("ERR^$_") }
            $q.Enqueue("DONE")
        }
        [void]$ps.AddScript($sb2)
        [void]$ps.AddParameters(@{ exe = $eng.Exe; q = $q })
        $ar = $ps.BeginInvoke()
        $script:LiveRepos = @()

        $t = [System.Windows.Threading.DispatcherTimer]::new()
        $t.Interval = [System.TimeSpan]::FromMilliseconds(300)
        $t.add_Tick({
            $raw = $null
            while ($q.TryDequeue([ref]$raw)) {
                if ($raw -eq 'DONE') {
                    $t.Stop()
                    try { $ps.EndInvoke($ar) | Out-Null; $ps.Dispose(); $rs.Close(); $rs.Dispose() } catch {}
                    $script:RepoRegState   = @{}
                    $script:RepoTrustState = @{}
                    foreach ($r in $script:LiveRepos) {
                        $script:RepoRegState[$r.Name]   = $true
                        $script:RepoTrustState[$r.Name] = $r.Policy
                    }
                    $script:RepoTabBusy = $false
                    BuildRepoUI
                    return
                } elseif ($raw -match '^REPO\^') {
                    $p = $raw.Substring(5).Split('^')
                    if ($p.Count -ge 3) {
                        $provVal = if ($p.Count -ge 4) { $p[3].Trim() } else { '' }
                        $script:LiveRepos += [PSCustomObject]@{
                            Name=$p[0].Trim(); Url=$p[1].Trim()
                            Policy=$p[2].Trim(); Provider=$provVal
                            IsRegistered=$true
                        }
                    }
                } elseif ($raw -eq 'NONE' -or $raw -eq 'NOMOD') {
                    # handled at DONE
                } elseif ($raw -match '^ERR') {
                    $t.Stop()
                    try { $ps.EndInvoke($ar) | Out-Null; $ps.Dispose(); $rs.Close(); $rs.Dispose() } catch {}
                    $script:RepoTabBusy = $false
                    BuildRepoUI
                    return
                }
            }
        }.GetNewClosure())
        $t.Start()
    }

    function BuildRepoUI {
        $pnlRepos.Children.Clear()
        $conv = [System.Windows.Media.BrushConverter]::new()
        $regNames = @{}
        foreach ($r in $script:LiveRepos) { $regNames[$r.Name] = $r }

        # ---- Section header ----
        $hdr = [System.Windows.Controls.TextBlock]::new()
        $hdr.Text = 'WELL-KNOWN REPOSITORIES'
        $hdr.Foreground = $conv.ConvertFromString('#252560')
        $hdr.FontSize = 9; $hdr.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
        $hdr.Margin = [System.Windows.Thickness]::new(2,0,0,8)
        $pnlRepos.Children.Add($hdr) | Out-Null

        foreach ($kr in $script:KnownRepos) {
            $isReg    = $regNames.ContainsKey($kr.Name)
            $isTrust  = if ($isReg) { $regNames[$kr.Name].Policy -eq 'Trusted' } else { $false }
            $krName   = $kr.Name   # capture for closures

            $card = [System.Windows.Controls.Border]::new()
            $card.Background      = $conv.ConvertFromString($kr.Color)
            $card.BorderBrush     = $conv.ConvertFromString($kr.BorderColor)
            $card.BorderThickness = [System.Windows.Thickness]::new(1)
            $card.CornerRadius    = [System.Windows.CornerRadius]::new(5)
            $card.Margin          = [System.Windows.Thickness]::new(0,0,0,8)
            $card.Padding         = [System.Windows.Thickness]::new(12,10,12,10)

            $vsp = [System.Windows.Controls.StackPanel]::new()

            # ---- Toggle + Name row ----
            $row = [System.Windows.Controls.DockPanel]::new()

            # Toggle (CheckBox with ToggleSwitch style)
            $chk = [System.Windows.Controls.CheckBox]::new()
            $chk.Style     = try { $win.FindResource('ToggleSwitch') } catch { $null }
            $chk.IsChecked = $isReg
            $chk.FontSize  = 10
            $chk.VerticalAlignment = 'Center'
            $chk.Margin    = [System.Windows.Thickness]::new(0,0,8,0)
            [System.Windows.Controls.DockPanel]::SetDock($chk, 'Right')

            # Use click event instead of Checked/Unchecked to avoid re-entry issues
            $urlCapture  = $kr.Url
            $trustCapture = $kr.DefaultTrust
            $canRemCapture = $kr.CanRemove
            $chk.add_Click({
                param($sender, $e)
                $e.Handled = $true
                $sender.IsEnabled = $false
                $nowChecked = [bool]$sender.IsChecked

                if ($nowChecked) {
                    # Register
                    $pol = $trustCapture
                    $cmd = '$ep="SilentlyContinue";try{Set-PSRepository -Name "' + $krName + '" -InstallationPolicy ' + $pol + ' -EA Stop 2>$null;Write-Output "OK"}catch{try{Register-PSRepository -Name "' + $krName + '" -SourceLocation "' + $urlCapture + '" -InstallationPolicy ' + $pol + ' -EA Stop 2>$null;Write-Output "OK"}catch{Write-Output "ERR|$_"}}'
                    $senderCapture = $sender
                    Invoke-RepoOp $cmd {
                        param($res)
                        if ($res -match '^OK') {
                            ML "[OK] Registered: $krName"
                        } else {
                            ML "[!!] Register failed: $krName - $res"
                        }
                        $senderCapture.IsEnabled = $true
                        $script:RepoTabBusy = $false
                        RefreshReposTab
                    }.GetNewClosure()
                } else {
                    # Unregister
                    if (-not $canRemCapture) {
                        [System.Windows.MessageBox]::Show("$krName cannot be unregistered.", 'Cannot Unregister', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
                        $sender.IsChecked = $true; $sender.IsEnabled = $true; return
                    }
                    $cmd = "Unregister-PSRepository -Name '$krName' -EA SilentlyContinue; Write-Output 'OK'"
                    $senderCapture2 = $sender
                    Invoke-RepoOp $cmd {
                        param($res)
                        ML "[OK] Unregistered: $krName"
                        $senderCapture2.IsEnabled = $true
                        $script:RepoTabBusy = $false
                        RefreshReposTab
                    }.GetNewClosure()
                }
            }.GetNewClosure())

            $row.Children.Add($chk) | Out-Null

            # Name
            $ntb = [System.Windows.Controls.TextBlock]::new()
            $ntb.Text = $kr.Name; $ntb.FontSize = 13
            $ntb.FontWeight = [System.Windows.FontWeights]::SemiBold
            $ntb.Foreground = $conv.ConvertFromString('#C0C0E0')
            $ntb.VerticalAlignment = 'Center'
            [System.Windows.Controls.DockPanel]::SetDock($ntb, 'Left')
            $row.Children.Add($ntb) | Out-Null

            # Official badge
            $offBrd = [System.Windows.Controls.Border]::new()
            $offBrd.Background    = $conv.ConvertFromString('#080818')
            $offBrd.CornerRadius  = [System.Windows.CornerRadius]::new(3)
            $offBrd.Padding       = [System.Windows.Thickness]::new(5,1,5,1)
            $offBrd.Margin        = [System.Windows.Thickness]::new(8,0,0,0)
            $offBrd.VerticalAlignment = 'Center'
            $offTb = [System.Windows.Controls.TextBlock]::new()
            $offTb.Text       = 'OFFICIAL'
            $offTb.FontSize   = 9
            $offTb.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
            $offTb.Foreground = $conv.ConvertFromString($kr.BadgeColor)
            $offBrd.Child = $offTb
            [System.Windows.Controls.DockPanel]::SetDock($offBrd, 'Left')
            $row.Children.Add($offBrd) | Out-Null

            # Trusted badge (only when registered)
            if ($isReg) {
                $tbrd = [System.Windows.Controls.Border]::new()
                $tBg  = if ($isTrust) {'#041504'} else {'#150A00'}
                $tFg  = if ($isTrust) {'#30B050'} else {'#C07020'}
                $tTxt = if ($isTrust) {'TRUSTED'}  else {'UNTRUSTED'}
                $tbrd.Background   = $conv.ConvertFromString($tBg)
                $tbrd.CornerRadius = [System.Windows.CornerRadius]::new(3)
                $tbrd.Padding      = [System.Windows.Thickness]::new(5,1,5,1)
                $tbrd.Margin       = [System.Windows.Thickness]::new(6,0,0,0)
                $tbrd.VerticalAlignment = 'Center'
                $ttb = [System.Windows.Controls.TextBlock]::new()
                $ttb.Text       = $tTxt
                $ttb.FontSize   = 9
                $ttb.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
                $ttb.Foreground = $conv.ConvertFromString($tFg)
                $tbrd.Child = $ttb
                [System.Windows.Controls.DockPanel]::SetDock($tbrd, 'Left')
                $row.Children.Add($tbrd) | Out-Null
            }

            $vsp.Children.Add($row) | Out-Null

            # URL
            $utb = [System.Windows.Controls.TextBlock]::new()
            $utb.Text        = $kr.Url
            $utb.FontFamily  = [System.Windows.Media.FontFamily]::new('Consolas')
            $utb.FontSize    = 10
            $utb.Foreground  = $conv.ConvertFromString('#252550')
            $utb.Margin      = [System.Windows.Thickness]::new(0,5,0,0)
            $utb.TextWrapping = 'Wrap'
            $vsp.Children.Add($utb) | Out-Null

            # Description
            $dtb = [System.Windows.Controls.TextBlock]::new()
            $dtb.Text         = $kr.Desc
            $dtb.FontSize     = 10
            $dtb.Foreground   = $conv.ConvertFromString($kr.DescColor)
            $dtb.Margin       = [System.Windows.Thickness]::new(0,4,0,0)
            $dtb.TextWrapping = 'Wrap'
            $vsp.Children.Add($dtb) | Out-Null

            $card.Child = $vsp
            $pnlRepos.Children.Add($card) | Out-Null
        }

        # ---- Custom registered repos (not in known list) ----
        $others = @($script:LiveRepos | Where-Object {
            $n = $_.Name
            -not ($script:KnownRepos | Where-Object { $_.Name -eq $n })
        })
        if ($others.Count -gt 0) {
            $sep = [System.Windows.Controls.TextBlock]::new()
            $sep.Text = 'CUSTOM REGISTERED REPOSITORIES'
            $sep.Foreground = $conv.ConvertFromString('#252560')
            $sep.FontSize = 9; $sep.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
            $sep.Margin = [System.Windows.Thickness]::new(2,8,0,6)
            $pnlRepos.Children.Add($sep) | Out-Null

            foreach ($repo in $others) {
                $isTrust2 = ($repo.Policy -eq 'Trusted')
                $card2 = [System.Windows.Controls.Border]::new()
                $card2.Background    = $conv.ConvertFromString('#07071F')
                $card2.BorderBrush   = $conv.ConvertFromString('#181848')
                $card2.BorderThickness = [System.Windows.Thickness]::new(1)
                $card2.CornerRadius  = [System.Windows.CornerRadius]::new(5)
                $card2.Margin = [System.Windows.Thickness]::new(0,0,0,6)
                $card2.Padding = [System.Windows.Thickness]::new(12,10,12,10)
                $card2.Cursor = [System.Windows.Input.Cursors]::Hand
                $rname2 = $repo.Name
                $card2.add_MouseLeftButtonUp({ $script:RepoSelected = $rname2; $lblRepoStatus.Text = "Selected: $rname2" }.GetNewClosure())

                $vsp2 = [System.Windows.Controls.StackPanel]::new()
                $row2  = [System.Windows.Controls.DockPanel]::new()
                $n2tb  = [System.Windows.Controls.TextBlock]::new()
                $n2tb.Text = $repo.Name; $n2tb.FontSize = 12
                $n2tb.FontWeight = [System.Windows.FontWeights]::SemiBold
                $n2tb.Foreground = $conv.ConvertFromString('#C0C0E0'); $n2tb.VerticalAlignment = 'Center'
                [System.Windows.Controls.DockPanel]::SetDock($n2tb, 'Left'); $row2.Children.Add($n2tb) | Out-Null

                $tbrd2 = [System.Windows.Controls.Border]::new()
                $t2Bg  = if ($isTrust2) {'#041504'} else {'#150A00'}
                $t2Fg  = if ($isTrust2) {'#30B050'} else {'#C07020'}
                $t2Txt = if ($isTrust2) {'TRUSTED'}  else {'UNTRUSTED'}
                $tbrd2.Background = $conv.ConvertFromString($t2Bg)
                $tbrd2.CornerRadius = [System.Windows.CornerRadius]::new(3)
                $tbrd2.Padding = [System.Windows.Thickness]::new(5,1,5,1); $tbrd2.Margin = [System.Windows.Thickness]::new(8,0,0,0); $tbrd2.VerticalAlignment = 'Center'
                $ttb2 = [System.Windows.Controls.TextBlock]::new()
                $ttb2.Text = $t2Txt; $ttb2.FontSize = 9
                $ttb2.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
                $ttb2.Foreground = $conv.ConvertFromString($t2Fg)
                $tbrd2.Child = $ttb2
                [System.Windows.Controls.DockPanel]::SetDock($tbrd2, 'Left'); $row2.Children.Add($tbrd2) | Out-Null

                $btnT2 = [System.Windows.Controls.Button]::new()
                $btnT2.Content = if($isTrust2){'Set Untrusted'}else{'Set Trusted'}
                $btnT2.FontSize = 9; $btnT2.Padding = [System.Windows.Thickness]::new(6,1,6,1)
                $btnT2.Margin = [System.Windows.Thickness]::new(4,0,0,0); $btnT2.VerticalAlignment = 'Center'
                $btnT2.Background = $conv.ConvertFromString('#0A0A1E')
                $btnT2.Foreground = $conv.ConvertFromString('#3050B0')
                $btnT2.BorderBrush = $conv.ConvertFromString('#181840')
                $rC2 = $repo
                $btnT2.add_Click({
                    $newPol = if($rC2.Policy -eq 'Trusted'){'Untrusted'}else{'Trusted'}
                    $cmd3 = "Set-PSRepository -Name '$($rC2.Name)' -InstallationPolicy $newPol -EA SilentlyContinue; Write-Output 'OK'"
                    Invoke-RepoOp $cmd3 {
                        ML "[OK] $($rC2.Name) set to $newPol"
                        $script:RepoTabBusy = $false; RefreshReposTab
                    }.GetNewClosure()
                }.GetNewClosure())
                [System.Windows.Controls.DockPanel]::SetDock($btnT2, 'Right'); $row2.Children.Add($btnT2) | Out-Null

                $vsp2.Children.Add($row2) | Out-Null
                $u2tb = [System.Windows.Controls.TextBlock]::new()
                $u2tb.Text = $repo.Url; $u2tb.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
                $u2tb.FontSize = 10; $u2tb.Foreground = $conv.ConvertFromString('#252550')
                $u2tb.Margin = [System.Windows.Thickness]::new(0,4,0,0); $u2tb.TextWrapping = 'Wrap'
                $vsp2.Children.Add($u2tb) | Out-Null
                $card2.Child = $vsp2
                $pnlRepos.Children.Add($card2) | Out-Null
            }
        }

        # ---- Custom catalog modules panel ----
        $custom = @($script:LiveCatalog | Where-Object { $_.Custom -eq $true })
        $lblCustomCount.Text = "  ($($custom.Count) custom)"
        foreach ($m in $custom) {
            $mc = [System.Windows.Controls.Border]::new()
            $mc.Background = $conv.ConvertFromString('#07070F')
            $mc.BorderBrush = $conv.ConvertFromString('#181848')
            $mc.BorderThickness = [System.Windows.Thickness]::new(1)
            $mc.CornerRadius = [System.Windows.CornerRadius]::new(4)
            $mc.Margin = [System.Windows.Thickness]::new(0,0,0,4)
            $mc.Padding = [System.Windows.Thickness]::new(10,6,10,6)
            $mdp = [System.Windows.Controls.DockPanel]::new()
            $mtb = [System.Windows.Controls.TextBlock]::new()
            $mtb.Text = $m.Name; $mtb.FontSize = 11
            $mtb.FontWeight = [System.Windows.FontWeights]::SemiBold
            $mtb.Foreground = $conv.ConvertFromString('#8080D0')
            [System.Windows.Controls.DockPanel]::SetDock($mtb, 'Left'); $mdp.Children.Add($mtb) | Out-Null
            $catBrd = [System.Windows.Controls.Border]::new()
            $catBrd.Background = $conv.ConvertFromString('#0A0A20')
            $catBrd.CornerRadius = [System.Windows.CornerRadius]::new(3)
            $catBrd.Padding = [System.Windows.Thickness]::new(5,1,5,1); $catBrd.Margin = [System.Windows.Thickness]::new(6,0,0,0); $catBrd.VerticalAlignment = 'Center'
            $catTb = [System.Windows.Controls.TextBlock]::new()
            $catTb.Text = $m.Cat; $catTb.FontSize = 9; $catTb.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
            $catTb.Foreground = $conv.ConvertFromString('#3040A0')
            $catBrd.Child = $catTb
            [System.Windows.Controls.DockPanel]::SetDock($catBrd, 'Left'); $mdp.Children.Add($catBrd) | Out-Null
            $mDesc2 = [System.Windows.Controls.TextBlock]::new()
            $mDesc2.Text = $m.Desc; $mDesc2.FontSize = 10
            $mDesc2.Foreground = $conv.ConvertFromString('#303060')
            $mDesc2.Margin = [System.Windows.Thickness]::new(0,3,0,0)
            $mSp2 = [System.Windows.Controls.StackPanel]::new()
            $mSp2.Children.Add($mdp) | Out-Null; $mSp2.Children.Add($mDesc2) | Out-Null
            $mc.Child = $mSp2; $pnlCustomMods.Children.Add($mc) | Out-Null
        }

        $regCount    = $script:LiveRepos.Count
        $customCount = $custom.Count
        $lblRepoStatus.Text = "$regCount registered repositories  |  $customCount custom catalog modules  |  Toggle switch to register/unregister"
    }

    $btnBrowseRepo = G 'btnBrowseRepo'
    $btnBrowseRepo.add_Click({
        $eng2 = Get-RepoEngine
        if (-not $eng2) { [System.Windows.MessageBox]::Show('No engine available.') | Out-Null; return }
        # Get current repos directly (don't depend on cached $script:LiveRepos)
        $eng3 = Get-RepoEngine
        if (-not $eng3) { [System.Windows.MessageBox]::Show('No engine available.') | Out-Null; return }
        $repoCheckCmd = '$ep="SilentlyContinue";try{(Get-PSRepository -EA $ep).Name -join ","}catch{"NONE"}'
        $repoCheckOut = & $eng3.Exe -NoProfile -NonInteractive -Command $repoCheckCmd 2>$null
        $repoNamesStr = ($repoCheckOut | Select-Object -Last 1).ToString().Trim()
        if (-not $repoNamesStr -or $repoNamesStr -eq 'NONE') {
            [System.Windows.MessageBox]::Show("No repositories registered yet.`nUse the toggle switches to register PSGallery first.", 'Browse Repos', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
            return
        }
        $repoNames = $repoNamesStr
        $lblRepoStatus.Text = "Browsing modules from: $repoNames ..."

        $safeRepos = $repoNamesStr -replace '"',''
        $cmd4 = (
            '$ep="SilentlyContinue";$ErrorActionPreference=$ep;$ProgressPreference=$ep;' +
            '$rnames="' + $safeRepos + '".Split(",") | ForEach-Object{$_.Trim()};' +
            'Find-Module -Repository $rnames -EA $ep 2>$null |' +
            'Select-Object -First 500 |' +
            'ForEach-Object {' +
            '  $d=$_.Description;if($d.Length -gt 80){$d=$d.Substring(0,80)};' +
            '  Write-Output ($_.Name+"^"+$_.Version+"^"+$_.Repository+"^"+$d)' +
            '}'
        )
        $browseRS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $browseRS.ApartmentState='MTA'; $browseRS.ThreadOptions='ReuseThread'; $browseRS.Open()
        $browsePS = [System.Management.Automation.PowerShell]::Create()
        $browsePS.Runspace = $browseRS
        $browseQ = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        $browseSB = {
            param($exe, $cmd, $q)
            try {
                $out = & $exe -NoProfile -NonInteractive -Command $cmd 2>$null
                foreach ($ln in $out) { $q.Enqueue($ln.ToString().Trim()) }
            } catch { $q.Enqueue("ERR^$_") }
            $q.Enqueue("BROWSE_DONE")
        }
        [void]$browsePS.AddScript($browseSB)
        [void]$browsePS.AddParameters(@{ exe=$eng2.Exe; cmd=$cmd4; q=$browseQ })
        $browseAR = $browsePS.BeginInvoke()
        $browseResults = [System.Collections.Generic.List[object]]::new()

        $browseTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $browseTimer.Interval = [System.TimeSpan]::FromMilliseconds(400)
        $browseTimer.add_Tick({
            $raw2 = $null
            while ($browseQ.TryDequeue([ref]$raw2)) {
                if ($raw2 -eq 'BROWSE_DONE') {
                    $browseTimer.Stop()
                    try { $browsePS.EndInvoke($browseAR)|Out-Null; $browsePS.Dispose(); $browseRS.Close(); $browseRS.Dispose() } catch {}
                    $cnt = $browseResults.Count
                    $lblRepoStatus.Text = "Found $cnt modules in registered repositories."
                    if ($cnt -eq 0) {
                        [System.Windows.MessageBox]::Show("No modules found. Make sure PSGallery is registered and try Refresh first.", 'Browse Repos') | Out-Null
                        return
                    }
                    # Show results IN the right panel as checkboxes
                    $script:BrowseMode = $true
                    $script:BrowseCheckboxes.Clear()
                    $pnlCustomMods.Children.Clear()
                    $lblRightPanelTitle.Text = "AVAILABLE REPO MODULES  ($cnt)"
                    $lblCustomCount.Text = ""
                    $btnInstallFromRepo.Visibility = 'Visible'
                    $btnClearBrowse.Visibility = 'Visible'
                    $conv5 = [System.Windows.Media.BrushConverter]::new()

                    # Add filter textbox at top
                    $filterBrd = [System.Windows.Controls.Border]::new()
                    $filterBrd.Margin = [System.Windows.Thickness]::new(0,0,0,6)
                    $filterBrd.Background = $conv5.ConvertFromString('#080818')
                    $filterBrd.BorderBrush = $conv5.ConvertFromString('#181838')
                    $filterBrd.BorderThickness = [System.Windows.Thickness]::new(1)
                    $filterBrd.Padding = [System.Windows.Thickness]::new(4,2,4,2)
                    $filterTb = [System.Windows.Controls.TextBox]::new()
                    $filterTb.Background = $conv5.ConvertFromString('#080818')
                    $filterTb.Foreground = $conv5.ConvertFromString('#8080C0')
                    $filterTb.BorderThickness = [System.Windows.Thickness]::new(0)
                    $filterTb.FontSize = 10
                    $filterTb.Text = "Filter modules..."
                    $filterBrd.Child = $filterTb
                    $pnlCustomMods.Children.Add($filterBrd) | Out-Null

                    # Render all browse results as checkboxes
                    $allChkBrd = [System.Collections.Generic.List[object]]::new()
                    foreach ($mod in $browseResults) {
                        $chkBrd = [System.Windows.Controls.Border]::new()
                        $chkBrd.Background = $conv5.ConvertFromString('#07070F')
                        $chkBrd.BorderBrush = $conv5.ConvertFromString('#111130')
                        $chkBrd.BorderThickness = [System.Windows.Thickness]::new(0,0,0,1)
                        $chkBrd.Padding = [System.Windows.Thickness]::new(4,3,4,3)
                        $chkBrd.Tag = $mod.Name

                        $dp5 = [System.Windows.Controls.DockPanel]::new()
                        $chk5 = [System.Windows.Controls.CheckBox]::new()
                        $chk5.VerticalAlignment = 'Center'
                        $chk5.Margin = [System.Windows.Thickness]::new(0,0,6,0)
                        $chk5.Tag = $mod
                        [System.Windows.Controls.DockPanel]::SetDock($chk5, 'Left')
                        $dp5.Children.Add($chk5) | Out-Null

                        $sp5 = [System.Windows.Controls.StackPanel]::new()
                        $ntb5 = [System.Windows.Controls.TextBlock]::new()
                        $ntb5.Text = $mod.Name
                        $ntb5.FontSize = 11
                        $ntb5.FontWeight = [System.Windows.FontWeights]::SemiBold
                        $ntb5.Foreground = $conv5.ConvertFromString('#8090D0')
                        $vtb5 = [System.Windows.Controls.TextBlock]::new()
                        $vtb5.Text = "v$($mod.Version)  [$($mod.Repo)]"
                        $vtb5.FontSize = 9
                        $vtb5.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
                        $vtb5.Foreground = $conv5.ConvertFromString('#303060')
                        $sp5.Children.Add($ntb5) | Out-Null
                        $sp5.Children.Add($vtb5) | Out-Null
                        $dp5.Children.Add($sp5) | Out-Null
                        $chkBrd.Child = $dp5
                        $pnlCustomMods.Children.Add($chkBrd) | Out-Null
                        $script:BrowseCheckboxes.Add($chk5) | Out-Null
                        $allChkBrd.Add($chkBrd) | Out-Null
                    }

                    # Filter textbox handler
                    $allChkBrdCapture = $allChkBrd
                    $filterTb.add_TextChanged({
                        $ft5 = $filterTb.Text.Trim().ToLower()
                        if ($ft5 -eq 'filter modules...' -or $ft5 -eq '') {
                            foreach ($b5 in $allChkBrdCapture) { $b5.Visibility = 'Visible' }
                            return
                        }
                        foreach ($b5 in $allChkBrdCapture) {
                            $b5.Visibility = if ($b5.Tag.ToLower() -like "*$ft5*") {'Visible'} else {'Collapsed'}
                        }
                    }.GetNewClosure())
                    $filterTb.add_GotFocus({ if ($filterTb.Text -eq 'Filter modules...') { $filterTb.Text = '' } })
                    $filterTb.add_LostFocus({ if ($filterTb.Text -eq '') { $filterTb.Text = 'Filter modules...' } })
                    return
                } elseif ($raw2 -match '\^' -and $raw2 -notmatch '^ERR') {
                    $p2 = $raw2.Split('^', 4)
                    if ($p2.Count -ge 3) {
                        $browseResults.Add([PSCustomObject]@{
                            Name=$p2[0]; Version=$p2[1]; Repo=$p2[2]
                            Desc=if($p2.Count -ge 4){$p2[3]}else{''}
                        })
                    }
                }
            }
        }.GetNewClosure())
        $browseTimer.Start()
    })

    # Install Selected from Repo
    $btnInstallFromRepo.add_Click({
        # Read checkboxes directly from the panel (avoids stale list issues)
        $selected5 = @(
            foreach ($child in $pnlCustomMods.Children) {
                if ($child -is [System.Windows.Controls.Border]) {
                    $dp = $child.Child
                    if ($dp -is [System.Windows.Controls.DockPanel]) {
                        foreach ($item in $dp.Children) {
                            if ($item -is [System.Windows.Controls.CheckBox] -and [bool]$item.IsChecked) {
                                $item.Tag
                            }
                        }
                    }
                }
            }
        )
        if ($selected5.Count -eq 0) {
            # Debug: count all checkboxes in panel
            $allChk = @(foreach ($ch in $pnlCustomMods.Children) {
                if ($ch -is [System.Windows.Controls.Border] -and $ch.Child -is [System.Windows.Controls.DockPanel]) {
                    foreach ($c2 in $ch.Child.Children) { if ($c2 -is [System.Windows.Controls.CheckBox]) { $c2 } }
                }
            })
            $checkedCount = @($allChk | Where-Object { [bool]$_.IsChecked }).Count
            [System.Windows.MessageBox]::Show("No modules selected. ($($allChk.Count) checkboxes found, $checkedCount checked)", 'Install from Repo', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
            return
        }
        # Ask scope
        $modList5 = ($selected5 | Select-Object -First 10 | ForEach-Object { "  - " + $_.Name }) -join "`n"
        $scopeAns = [System.Windows.MessageBox]::Show(
            "Install $($selected5.Count) module(s):`n`n$modList5`n`nYES = AllUsers (requires Admin)`nNO  = CurrentUser",
            "Choose Install Scope",
            [System.Windows.MessageBoxButton]::YesNoCancel,
            [System.Windows.MessageBoxImage]::Question)
        if ($scopeAns -eq 'Cancel') { return }
        $scope5 = if ($scopeAns -eq 'Yes') { 'AllUsers' } else { 'CurrentUser' }
        $eng5 = Get-RepoEngine
        $names5 = ($selected5 | ForEach-Object { '"' + ($_.Name.ToString() -replace '"','') + '"' }) -join ','
        $cmd5 = '$ep="SilentlyContinue";$ErrorActionPreference=$ep;$ProgressPreference=$ep;' +
                '@(' + $names5 + ') | ForEach-Object {' +
                '  $mn=$_;Write-Host "Installing $mn...";' +
                '  try{Install-Module -Name $mn -Scope ' + $scope5 + ' -Force -AllowClobber -EA Stop 2>$null;Write-Host "[OK] $mn installed"}' +
                '  catch{$em=$_.Exception.Message;Write-Host "[ERR] $mn - $em"}' +
                '}'
        ML "Installing $($selected5.Count) module(s) to $scope5 scope..."
        $lblRepoStatus.Text = "Installing... check Terminal Output for progress."
        $script:TermLines.Clear()
        $termQ.Enqueue("START_BATCH")
        $i5RS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $i5RS.ApartmentState='MTA'; $i5RS.ThreadOptions='ReuseThread'; $i5RS.Open()
        $i5PS = [System.Management.Automation.PowerShell]::Create()
        $i5PS.Runspace = $i5RS
        $i5Q = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        $i5SB = {
            param($exe,$cmd,$q)
            try {
                $out = & $exe -NoProfile -NonInteractive -Command $cmd 2>&1
                foreach ($ln in $out) { if($ln){ $q.Enqueue($ln.ToString().Trim()) } }
            } catch { $q.Enqueue("[ERR] $_") }
            $q.Enqueue("INSTALL_DONE")
        }
        [void]$i5PS.AddScript($i5SB)
        [void]$i5PS.AddParameters(@{exe=$eng5.Exe; cmd=$cmd5; q=$i5Q})
        $i5AR = $i5PS.BeginInvoke()
        $i5T = [System.Windows.Threading.DispatcherTimer]::new()
        $i5T.Interval = [System.TimeSpan]::FromMilliseconds(300)
        $i5T.add_Tick({
            $ln5 = $null
            while ($i5Q.TryDequeue([ref]$ln5)) {
                if ($ln5 -eq 'INSTALL_DONE') {
                    $i5T.Stop()
                    try{$i5PS.EndInvoke($i5AR)|Out-Null;$i5PS.Dispose();$i5RS.Close();$i5RS.Dispose()}catch{}

                    # Auto-add newly installed modules to catalog so they appear in Module Catalog scan
                    $addedToLog = [System.Collections.Generic.List[string]]::new()
                    foreach ($m5 in $selected5) {
                        $mName5 = $m5.Name.ToString().Trim()
                        $mRepo5 = $m5.Repo.ToString().Trim()
                        # Only add if not already in allRows
                        $exists5 = $allRows | Where-Object { $_.Name -eq $mName5 }
                        if (-not $exists5) {
                            try {
                                $newEntry5 = [PSCustomObject]@{
                                    Name   = $mName5
                                    Cat    = $mRepo5
                                    Desc   = "Installed from $mRepo5"
                                    Custom = $true
                                }
                                $Global:Catalog += $newEntry5
                                $newRow5 = [PSMod.Row]::new()
                                $newRow5.Name = $mName5
                                $newRow5.Cat  = $mRepo5
                                $newRow5.Desc = "Installed from $mRepo5"
                                $newRow5.St   = 'Pending'
                                $allRows.Add($newRow5)
                                $addedToLog.Add($mName5) | Out-Null
                                # Persist to custom modules JSON
                                Save-CustomModules
                            } catch {}
                        }
                    }
                    if ($addedToLog.Count -gt 0) {
                        ML "[OK] Added to Module Catalog: $($addedToLog -join ', ')"
                        # Rebuild chip tabs to show new category
                        MakeChips
                        # Refresh DataGrid
                        $script:AllRowsView.Refresh()
                    }

                    ML "Install complete."
                    $lblRepoStatus.Text = "Installation complete. New modules added to catalog - click Refresh Status to scan them."
                    return
                }
                TL $ln5
            }
        }.GetNewClosure())
        $i5T.Start()
    })

    $btnClearBrowse.add_Click({
        $script:BrowseMode = $false
        $script:BrowseCheckboxes.Clear()
        $btnInstallFromRepo.Visibility = 'Collapsed'
        $btnClearBrowse.Visibility = 'Collapsed'
        $lblRightPanelTitle.Text = 'CUSTOM CATALOG MODULES'
        $lblCustomCount.Text = ''
        $pnlCustomMods.Children.Clear()
        # Rebuild custom catalog modules display without touching repos
        $conv9 = [System.Windows.Media.BrushConverter]::new()
        $custom9 = @($script:LiveCatalog | Where-Object { $_.Custom -eq $true })
        $lblCustomCount.Text = "  ($($custom9.Count) custom)"
        foreach ($m9 in $custom9) {
            $mc9 = [System.Windows.Controls.Border]::new()
            $mc9.Background = $conv9.ConvertFromString('#07070F')
            $mc9.BorderBrush = $conv9.ConvertFromString('#181848')
            $mc9.BorderThickness = [System.Windows.Thickness]::new(1)
            $mc9.CornerRadius = [System.Windows.CornerRadius]::new(4)
            $mc9.Margin = [System.Windows.Thickness]::new(0,0,0,4)
            $mc9.Padding = [System.Windows.Thickness]::new(10,6,10,6)
            $sp9 = [System.Windows.Controls.StackPanel]::new()
            $nt9 = [System.Windows.Controls.TextBlock]::new()
            $nt9.Text = $m9.Name; $nt9.FontSize = 11
            $nt9.FontWeight = [System.Windows.FontWeights]::SemiBold
            $nt9.Foreground = $conv9.ConvertFromString('#8080D0')
            $dt9 = [System.Windows.Controls.TextBlock]::new()
            $dt9.Text = $m9.Desc; $dt9.FontSize = 10
            $dt9.Foreground = $conv9.ConvertFromString('#303060')
            $sp9.Children.Add($nt9) | Out-Null
            $sp9.Children.Add($dt9) | Out-Null
            $mc9.Child = $sp9
            $pnlCustomMods.Children.Add($mc9) | Out-Null
        }
    })

    $btnRepoRefresh.add_Click({ $script:RepoTabBusy = $false; RefreshReposTab })



    $btnRepoRegister.add_Click({
        # Dialog to register a new repo
        $dlgXaml = [xml]@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Register PSRepository" Width="520" Height="300"
        Background="#0A0A18" WindowStartupLocation="CenterOwner"
        FontFamily="Segoe UI" FontSize="12">
  <Grid Margin="20">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" Text="Register a new PSRepository" Foreground="#5A88FF" FontSize="14" FontWeight="Bold" Margin="0,0,0,16"/>
    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Name:" Foreground="#9090C0" Width="110" VerticalAlignment="Center"/>
      <TextBox Name="txtRegName" Width="280" Padding="4,2" Background="#0A0A20" Foreground="#C0C0E0" BorderBrush="#28286A" ToolTip="Unique name for this repository"/>
    </StackPanel>
    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Source URL:" Foreground="#9090C0" Width="110" VerticalAlignment="Center"/>
      <TextBox Name="txtRegUrl" Width="280" Padding="4,2" Background="#0A0A20" Foreground="#C0C0E0" BorderBrush="#28286A" ToolTip="NuGet feed URL e.g. https://myserver/nuget/v2"/>
    </StackPanel>
    <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Policy:" Foreground="#9090C0" Width="110" VerticalAlignment="Center"/>
      <ComboBox Name="cmbRegPolicy" Width="140" Background="#0A0A20" Foreground="#C0C0E0" BorderBrush="#28286A">
        <ComboBoxItem Content="Trusted" IsSelected="True"/>
        <ComboBoxItem Content="Untrusted"/>
      </ComboBox>
    </StackPanel>
    <TextBlock Grid.Row="4" Text="Note: repository must be a NuGet v2/v3 feed. Internal repos (Nexus, Artifactory, ProGet) are fully supported." Foreground="#303060" FontSize="10" TextWrapping="Wrap" Margin="0,4,0,0"/>
    <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button Name="btnRegOK"     Content="Register"  Width="90" Height="30" Background="#1A3CC0" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand" Margin="0,0,8,0"/>
      <Button Name="btnRegCancel" Content="Cancel"    Width="80" Height="30" Background="#1A1A2E" Foreground="#8080A0" BorderBrush="#282850" Cursor="Hand"/>
    </StackPanel>
  </Grid>
</Window>
'@
        try {
            $dlgR = [System.Xml.XmlNodeReader]::new($dlgXaml)
            $dlgW = $null
            try { $dlgW = [System.Windows.Markup.XamlReader]::Load($dlgR) }
            catch { [System.Windows.MessageBox]::Show("XAML error: $_") | Out-Null; return }
            if (-not $dlgW) { [System.Windows.MessageBox]::Show("Dialog failed to load.") | Out-Null; return }
            try { $dlgW.Owner = $script:MainWindow } catch {}
            $rName = $dlgW.FindName('txtRegName'); $rUrl = $dlgW.FindName('txtRegUrl')
            $rPol  = $dlgW.FindName('cmbRegPolicy')
            $dlgW.FindName('btnRegCancel').add_Click({ $dlgW.Close() })
            $dlgW.FindName('btnRegOK').add_Click({
                $n = $rName.Text.Trim(); $u = $rUrl.Text.Trim()
                $p = ($rPol.SelectedItem).Content
                if (-not $n -or -not $u) {
                    [System.Windows.MessageBox]::Show('Name and URL are required.') | Out-Null; return
                }
                $eng2 = GetEng
                $nSafe = $n -replace '"',''
                $uSafe = $u -replace '"',''
                $sc = '$ep="SilentlyContinue";' +
                      'try{Set-PSRepository -Name "' + $nSafe + '" -SourceLocation "' + $uSafe + '" -InstallationPolicy ' + $p + ' -EA Stop}' +
                      'catch{try{Register-PSRepository -Name "' + $nSafe + '" -SourceLocation "' + $uSafe + '" -InstallationPolicy ' + $p + ' -EA Stop}catch{Write-Output "FAIL:$_";return}};' +
                      'Write-Output "OK"'
                try {
                    $out = & $eng2.Exe -NoProfile -NonInteractive -Command $sc 2>&1
                    if ($out -match 'OK') {
                        ML "[OK] Repository registered: $n"
                        $dlgW.Close()
                        $script:RepoTabBusy = $false
                        RefreshReposTab
                    } else {
                        [System.Windows.MessageBox]::Show("Registration output:`n$($out -join "`n")") | Out-Null
                    }
                } catch {
                    [System.Windows.MessageBox]::Show("Error: $_") | Out-Null
                }
            })
            $dlgW.ShowDialog() | Out-Null
        } catch { [System.Windows.MessageBox]::Show("Dialog error: $_") | Out-Null }
    })

    $btnRepoUnreg.add_Click({
        $sel = $script:RepoSelected
        if (-not $sel) {
            [System.Windows.MessageBox]::Show('Click a repository card to select it first.','Unregister') | Out-Null; return
        }
        if ($sel -eq 'PSGallery') {
            [System.Windows.MessageBox]::Show('PSGallery cannot be unregistered.','Unregister') | Out-Null; return
        }
        $ans = [System.Windows.MessageBox]::Show(
            "Unregister repository: $sel`n`nThis removes it from all PS engines. Continue?",
            'Unregister Repository', [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning)
        if ($ans -eq 'Yes') {
            $eng2 = GetEng
            & $eng2.Exe -NoProfile -NonInteractive -Command "Unregister-PSRepository -Name '$sel' -EA SilentlyContinue" 2>$null
            ML "[OK] Unregistered repository: $sel"
            $script:RepoSelected = $null
            RefreshReposTab
        }
    })

    $btnAddCatalogMod.add_Click({
        $dlgXaml2 = [xml]@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Add Module to Catalog" Width="560" Height="380"
        Background="#0A0A18" WindowStartupLocation="CenterOwner"
        FontFamily="Segoe UI" FontSize="12">
  <Grid Margin="20">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" Text="Add Module to Scan Catalog" Foreground="#5A88FF" FontSize="14" FontWeight="Bold" Margin="0,0,0,16"/>
    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Module Name:" Foreground="#9090C0" Width="120" VerticalAlignment="Center"/>
      <TextBox Name="txtModName" Width="280" Padding="4,2" Background="#0A0A20" Foreground="#C0C0E0" BorderBrush="#28286A" ToolTip="Exact name from PSGallery or your repo"/>
      <Button Name="btnModVerify" Content="Verify" Width="60" Height="24" Background="#0E1A3A" Foreground="#3A6AFF" BorderBrush="#181858" Margin="4,0,0,0" FontSize="10" Cursor="Hand"/>
    </StackPanel>
    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Category:" Foreground="#9090C0" Width="120" VerticalAlignment="Center"/>
      <ComboBox Name="cmbModCat" Width="160" Background="#0A0A20" Foreground="#C0C0E0" BorderBrush="#28286A">
        <ComboBoxItem Content="System"/>
        <ComboBoxItem Content="Azure"/>
        <ComboBoxItem Content="Graph"/>
        <ComboBoxItem Content="ActiveDir"/>
        <ComboBoxItem Content="Database"/>
        <ComboBoxItem Content="Security"/>
        <ComboBoxItem Content="Terminal"/>
        <ComboBoxItem Content="Utilities"/>
        <ComboBoxItem Content="VMware"/>
        <ComboBoxItem Content="Custom" IsSelected="True"/>
        <!-- Dynamic: populated with registered repo names at runtime -->
      </ComboBox>
    </StackPanel>
    <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Description:" Foreground="#9090C0" Width="120" VerticalAlignment="Center"/>
      <TextBox Name="txtModDesc" Width="280" Padding="4,2" Background="#0A0A20" Foreground="#C0C0E0" BorderBrush="#28286A"/>
    </StackPanel>
    <TextBlock Name="lblModVerify" Grid.Row="4" Foreground="#305030" FontSize="10" FontFamily="Consolas" Margin="0,0,0,8" TextWrapping="Wrap" Text=""/>
    <TextBlock Grid.Row="5" Foreground="#303060" FontSize="10" TextWrapping="Wrap" Margin="0,0,0,4"
               Text="The module will be added to the scan catalog and saved to PSModuleManager_custom.json next to this script. It will appear in the Module Catalog tab under its category."/>
    <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button Name="btnModOK"     Content="Add to Catalog" Width="110" Height="30" Background="#1A3CC0" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand" Margin="0,0,8,0"/>
      <Button Name="btnModCancel" Content="Cancel"         Width="80"  Height="30" Background="#1A1A2E" Foreground="#8080A0" BorderBrush="#282850" Cursor="Hand"/>
    </StackPanel>
  </Grid>
</Window>
'@
        try {
            $dlgR2 = [System.Xml.XmlNodeReader]::new($dlgXaml2)
            $dlgW2 = [System.Windows.Markup.XamlReader]::Load($dlgR2)
            $dlgW2.Owner = $script:MainWindow
            $mName  = $dlgW2.FindName('txtModName'); $mDesc = $dlgW2.FindName('txtModDesc')
            $mCat   = $dlgW2.FindName('cmbModCat');  $mVLbl = $dlgW2.FindName('lblModVerify')
            $dlgW2.FindName('btnModCancel').add_Click({ $dlgW2.Close() })
            # Add registered repo names to category dropdown
            $mCatCmb2 = $dlgW2.FindName('cmbModCat')
            foreach ($rr in $script:LiveRepos) {
                if ($rr.Name -ne 'PSGallery') {
                    $rItem = [System.Windows.Controls.ComboBoxItem]::new()
                    $rItem.Content = $rr.Name
                    $mCatCmb2.Items.Add($rItem) | Out-Null
                }
            }
            $dlgW2.FindName('btnModVerify').add_Click({
                $n2 = $mName.Text.Trim()
                if (-not $n2) { $mVLbl.Text = "Enter a module name first."; return }
                $mVLbl.Text = "Searching PSGallery..."
                $eng2 = GetEng
                $sc2 = ('$m=Find-Module -Name "' + ($n2 -replace '"','') +
                        '" -EA SilentlyContinue 2>$null|Select-Object -First 1;' +
                        'if($m){Write-Output ($m.Name+"|"+$m.Version+"|"+$m.Description)}' +
                        'else{Write-Output "NOT_FOUND"}')
                $out2 = & $eng2.Exe -NoProfile -NonInteractive -Command $sc2 2>$null
                $r2 = ($out2 | Select-Object -First 1)
                if ($r2 -and $r2 -ne 'NOT_FOUND') {
                    $parts2 = $r2.Split('|', 3)
                    $mName.Text = $parts2[0]
                    $mVLbl.Foreground = '#30A050'
                    $mVLbl.Text = "Found: v$($parts2[1]) &#8212; $($parts2[2])"
                    if (-not $mDesc.Text -and $parts2.Count -ge 3) {
                        $mDesc.Text = $parts2[2].Substring(0, [Math]::Min(80, $parts2[2].Length))
                    }
                } else {
                    $mVLbl.Foreground = '#A03030'
                    $mVLbl.Text = "Not found in PSGallery &#8212; you can still add it manually."
                }
            })
            $dlgW2.FindName('btnModOK').add_Click({
                $n3  = $mName.Text.Trim()
                $d3  = $mDesc.Text.Trim()
                $c3  = ($mCat.SelectedItem).Content
                if (-not $n3) {
                    [System.Windows.MessageBox]::Show('Module name is required.') | Out-Null; return
                }
                # Check not already in catalog
                if ($script:LiveCatalog | Where-Object { $_.Name -eq $n3 }) {
                    [System.Windows.MessageBox]::Show("'$n3' is already in the catalog.") | Out-Null; return
                }
                # Add to LiveCatalog
                $entry = [PSCustomObject]@{Name=$n3; Cat=$c3; Desc=$d3; Custom=$true}
                $script:LiveCatalog.Add($entry)
                # Add to allRows
                $nr = [PSMod.Row]::new()
                $nr.Name=$n3; $nr.Cat=$c3; $nr.Desc=$d3; $nr.St='Pending'
                $allRows.Add($nr)
                # Save to JSON
                $customOnly = $script:LiveCatalog | Where-Object { $_.Custom -eq $true }
                try {
                    $customOnly | Select-Object Name,Cat,Desc |
                        ConvertTo-Json -Depth 2 |
                        Set-Content $customModsFile -Encoding UTF8
                } catch { ML "[WARN] Could not save custom modules: $_" }
                # Rebuild UI
                MakeChips; DoFilter
                ML "[OK] Added '$n3' (Category: $c3) to catalog. Total: $($allRows.Count) modules."
                $lblStatus.Text = "Added: $n3 &#8212; Run Refresh Status to scan it."
                $dlgW2.Close()
                RefreshReposTab
            })
            $dlgW2.ShowDialog() | Out-Null
        } catch { [System.Windows.MessageBox]::Show("Dialog error: $_") | Out-Null }
    })

    $btnRemCatalogMod.add_Click({
        # Show picker of custom modules only
        $custom = @($script:LiveCatalog | Where-Object { $_.Custom -eq $true })
        if ($custom.Count -eq 0) {
            [System.Windows.MessageBox]::Show('No custom modules in catalog to remove.`nBuilt-in modules cannot be removed.','Remove from Catalog') | Out-Null
            return
        }
        $names = ($custom | ForEach-Object { $_.Name }) -join "`n  "
        $ans = [System.Windows.MessageBox]::Show(
            "Custom modules in catalog:`n  $names`n`nEnter the module name to remove in the next dialog.",
            'Remove from Catalog', [System.Windows.MessageBoxButton]::OKCancel,
            [System.Windows.MessageBoxImage]::Question)
        if ($ans -ne 'OK') { return }
        $toRem = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Type the module name to remove from catalog:`n(Custom only &#8212; built-in modules cannot be removed)",
            "Remove from Catalog", "")
        if (-not $toRem) { return }
        $found = $script:LiveCatalog | Where-Object { $_.Name -eq $toRem -and $_.Custom -eq $true }
        if (-not $found) {
            [System.Windows.MessageBox]::Show("'$toRem' not found in custom catalog.") | Out-Null; return
        }
        $script:LiveCatalog.Remove($found) | Out-Null
        # Remove from allRows too
        $rowToRem = $allRows | Where-Object { $_.Name -eq $toRem }
        if ($rowToRem) { $allRows.Remove($rowToRem) | Out-Null }
        # Save updated JSON
        $customOnly = $script:LiveCatalog | Where-Object { $_.Custom -eq $true }
        if ($customOnly) {
            $customOnly | Select-Object Name,Cat,Desc | ConvertTo-Json -Depth 2 |
                Set-Content $customModsFile -Encoding UTF8
        } else {
            if (Test-Path $customModsFile) { Remove-Item $customModsFile -Force }
        }
        MakeChips; DoFilter
        ML "[OK] Removed '$toRem' from catalog. Total: $($allRows.Count) modules."
        RefreshReposTab
    })

    $btnRepoFind.add_Click({
        $searchTerm = $txtRepoSearch.Text.Trim()
        if (-not $searchTerm) {
            [System.Windows.MessageBox]::Show('Enter a module name or keyword to search.','Search') | Out-Null; return
        }
        $lblRepoStatus.Text = "Searching for '$searchTerm'..."
        $btnRepoFind.IsEnabled = $false
        $engS = Get-RepoEngine
        $stSafe = $searchTerm -replace '"','' -replace "'",''
        $scSearch = (
            '$ep="SilentlyContinue";$ErrorActionPreference=$ep;$ProgressPreference=$ep;' +
            'Find-Module -Name "*' + $stSafe + '*" -EA $ep 2>$null |' +
            'Select-Object -First 20 |' +
            'ForEach-Object {' +
            '  $d=$_.Description;if($d.Length -gt 100){$d=$d.Substring(0,100)};' +
            '  Write-Output ($_.Name+"|"+$_.Version+"|"+$_.Repository+"|"+$d)' +
            '}'
        )
        $srQ = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        $srRS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $srRS.ApartmentState='MTA'; $srRS.ThreadOptions='ReuseThread'; $srRS.Open()
        $srPS = [System.Management.Automation.PowerShell]::Create()
        $srPS.Runspace = $srRS
        $srSB = {
            param($exe,$cmd,$q)
            try {
                $out = & $exe -NoProfile -NonInteractive -Command $cmd 2>$null
                foreach ($ln in $out) { if ($ln) { $q.Enqueue($ln.ToString().Trim()) } }
            } catch { $q.Enqueue("ERR|$_") }
            $q.Enqueue("SEARCH_DONE")
        }
        [void]$srPS.AddScript($srSB)
        [void]$srPS.AddParameters(@{exe=$engS.Exe; cmd=$scSearch; q=$srQ})
        $srAR = $srPS.BeginInvoke()
        $srResults = [System.Collections.Generic.List[string]]::new()
        $srTerm = $searchTerm

        $srTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $srTimer.Interval = [System.TimeSpan]::FromMilliseconds(300)
        $srTimer.add_Tick({
            $raw3 = $null
            while ($srQ.TryDequeue([ref]$raw3)) {
                if ($raw3 -eq 'SEARCH_DONE') {
                    $srTimer.Stop()
                    try { $srPS.EndInvoke($srAR)|Out-Null; $srPS.Dispose(); $srRS.Close(); $srRS.Dispose() } catch {}
                    $btnRepoFind.IsEnabled = $true
                    $cnt3 = $srResults.Count
                    if ($cnt3 -eq 0) {
                        $lblRepoStatus.Text = "No results for '$srTerm'."
                        return
                    }
                    $lblRepoStatus.Text = "Search: $cnt3 result(s) for '$srTerm'. Check boxes and click Install Selected."
                    # Show results in the right panel as checkboxes (same as Browse)
                    $script:BrowseMode = $true
                    $script:BrowseCheckboxes.Clear()
                    $pnlCustomMods.Children.Clear()
                    $lblRightPanelTitle.Text = "SEARCH RESULTS: $srTerm  ($cnt3)"
                    $lblCustomCount.Text = ""
                    $btnInstallFromRepo.Visibility = 'Visible'
                    $btnClearBrowse.Visibility = 'Visible'
                    $convS = [System.Windows.Media.BrushConverter]::new()
                    $allChkBrdS = [System.Collections.Generic.List[object]]::new()
                    foreach ($line3 in $srResults) {
                        $p3 = $line3.Split('|', 4)
                        if ($p3.Count -lt 2) { continue }
                        $chkBrdS = [System.Windows.Controls.Border]::new()
                        $chkBrdS.Background = $convS.ConvertFromString('#07070F')
                        $chkBrdS.BorderBrush = $convS.ConvertFromString('#111130')
                        $chkBrdS.BorderThickness = [System.Windows.Thickness]::new(0,0,0,1)
                        $chkBrdS.Padding = [System.Windows.Thickness]::new(4,3,4,3)
                        $chkBrdS.Tag = $p3[0].Trim()
                        $dpS = [System.Windows.Controls.DockPanel]::new()
                        $chkS = [System.Windows.Controls.CheckBox]::new()
                        $chkS.VerticalAlignment = 'Center'
                        $chkS.Margin = [System.Windows.Thickness]::new(0,0,6,0)
                        $chkS.Tag = [PSCustomObject]@{ Name=$p3[0].Trim(); Version=$p3[1].Trim(); Repo=if($p3.Count -ge 3){$p3[2].Trim()}else{'PSGallery'} }
                        [System.Windows.Controls.DockPanel]::SetDock($chkS, 'Left')
                        $dpS.Children.Add($chkS) | Out-Null
                        $spS = [System.Windows.Controls.StackPanel]::new()
                        $ntbS = [System.Windows.Controls.TextBlock]::new()
                        $ntbS.Text = $p3[0].Trim(); $ntbS.FontSize = 11
                        $ntbS.FontWeight = [System.Windows.FontWeights]::SemiBold
                        $ntbS.Foreground = $convS.ConvertFromString('#8090D0')
                        $vtbS = [System.Windows.Controls.TextBlock]::new()
                        $repoS = if ($p3.Count -ge 3) { $p3[2].Trim() } else { 'PSGallery' }
                        $vtbS.Text = "v$($p3[1].Trim())  [$repoS]"
                        $vtbS.FontSize = 9
                        $vtbS.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
                        $vtbS.Foreground = $convS.ConvertFromString('#303060')
                        $spS.Children.Add($ntbS) | Out-Null
                        $spS.Children.Add($vtbS) | Out-Null
                        $dpS.Children.Add($spS) | Out-Null
                        $chkBrdS.Child = $dpS
                        $pnlCustomMods.Children.Add($chkBrdS) | Out-Null
                        $script:BrowseCheckboxes.Add($chkS) | Out-Null
                        $allChkBrdS.Add($chkBrdS) | Out-Null
                    }
                    return
                } elseif ($raw3 -match '\|' -and $raw3 -notmatch '^ERR') {
                    $srResults.Add($raw3)
                }
            }
        }.GetNewClosure())
        $srTimer.Start()
    })

    # Load repos tab when selected
    $script:LastSelectedTab = ''
    $tabC.add_SelectionChanged({
        # Guard: skip if event bubbled from a child control (DataGrid etc.)
        $src = $args[1].OriginalSource
        if ($src -and $src.GetType().Name -ne 'TabControl') { return }
        $selected = $tabC.SelectedItem
        if (-not $selected) { return }
        $hdr = $selected.Header.ToString()
        # Skip if same tab re-selected (avoids double-fire)
        if ($hdr -eq $script:LastSelectedTab) { return }
        $script:LastSelectedTab = $hdr
        switch ($hdr) {
            'Log Viewer'   { ReloadLogView }
            'Repositories' {
                $script:RepoTabBusy = $false
                RefreshReposTab
            }
            'Module Paths' {
                if ($script:PathsTabTimer -and $script:PathsTabTimer.IsEnabled) {
                    $script:PathsTabTimer.Stop()
                }
                $script:PathsTabTimer = [System.Windows.Threading.DispatcherTimer]::new()
                $script:PathsTabTimer.Interval = [System.TimeSpan]::FromMilliseconds(200)
                $script:PathsTabTimer.add_Tick({ $script:PathsTabTimer.Stop(); RefreshPathsTab })
                $script:PathsTabTimer.Start()
            }
        }
    })

    # - Engine log helper ---------
    $script:EngLogLines = [System.Collections.Generic.List[string]]::new()
    function EL([string]$m) {
        $line = "[$(Get-Date -f 'HH:mm:ss')] $m"
        $script:EngLogLines.Add($line)
        $txtEngLog.Text = ($script:EngLogLines -join "`n")
        $svEngLog.ScrollToEnd()
        TL $m 'ENGINE'
        UiLog $m 'INFO'
    }

    # - Engine version check ------
    $btnCheckEngVer.add_Click({
        $btnCheckEngVer.IsEnabled = $false
        $btnCheckEngVer.Content   = 'Checking...'
        EL "Checking latest PS versions from GitHub..."

        $script:EngCheckRS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:EngCheckRS.ApartmentState = 'MTA'; $script:EngCheckRS.ThreadOptions = 'ReuseThread'
        $script:EngCheckRS.Open()
        $script:EngCheckPS = [System.Management.Automation.PowerShell]::Create()
        $script:EngCheckPS.Runspace = $script:EngCheckRS

        [void]$script:EngCheckPS.AddScript({
            param($termQ)
            $ep = 'SilentlyContinue'
            $ErrorActionPreference = $ep; $ProgressPreference = $ep

            function Push([string]$m) { try { $termQ.Enqueue("[ENGINE] $m") } catch {} }

            $results = @{}

            # Check PS7 latest via GitHub API
            Push "Querying GitHub for latest PowerShell 7..."
            try {
                $rel = Invoke-RestMethod 'https://api.github.com/repos/PowerShell/PowerShell/releases/latest' -EA $ep
                if ($rel) {
                    $ver = $rel.tag_name -replace '^v',''
                    $dl  = ($rel.assets | Where-Object { $_.name -like '*win-x64.msi' } | Select-Object -First 1).browser_download_url
                    $results['PS7_Latest']    = $ver
                    $results['PS7_DLUrl']     = $dl
                    $results['PS7_Published'] = $rel.published_at
                    Push "PS7 latest: v$ver (published: $($rel.published_at))"
                    Push "Download: $dl"
                }
            } catch { Push "PS7 GitHub query failed: $_" }

            # Check winget latest
            Push "Querying GitHub for latest winget..."
            try {
                $wrel = Invoke-RestMethod 'https://api.github.com/repos/microsoft/winget-cli/releases/latest' -EA $ep
                if ($wrel) {
                    $wver = $wrel.tag_name -replace '^v',''
                    $wdl  = ($wrel.assets | Where-Object { $_.name -like '*.msixbundle' } | Select-Object -First 1).browser_download_url
                    $results['Winget_Latest'] = $wver
                    $results['Winget_DLUrl']  = $wdl
                    Push "winget latest: v$wver"
                    Push "Download: $wdl"
                }
            } catch { Push "winget GitHub query failed: $_" }

            # Check current winget version
            try {
                $wv = & winget.exe --version 2>$null
                if ($wv) { $results['Winget_Current'] = $wv.Trim() -replace '^v',''; Push "winget installed: $($wv.Trim())" }
                else { $results['Winget_Current'] = 'Not found'; Push "winget: not found in PATH" }
            } catch { $results['Winget_Current'] = 'Not found' }

            Push "VERSION_CHECK_DONE"
            $termQ.Enqueue("ENG_RESULT|" + ($results.Keys | ForEach-Object { $_ + "=" + $results[$_] }) -join "|")
        })
        [void]$script:EngCheckPS.AddParameters(@{ termQ = $termQ })
        $script:EngCheckAR = $script:EngCheckPS.BeginInvoke()

        $script:EngCheckTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:EngCheckTimer.Interval = [System.TimeSpan]::FromMilliseconds(400)
        $script:EngCheckTimer.add_Tick({
            $tLine = $null
            while ($termQ.TryDequeue([ref]$tLine)) {
                if ($tLine -match '^\[ENGINE\] (.+)$') { EL $Matches[1] }
                elseif ($tLine -match '^ENG_RESULT\|(.+)$') {
                    # Parse results and update engine cards
                    $data = @{}
                    $Matches[1].Split('|') | ForEach-Object {
                        $kv = $_.Split('=',2); if ($kv.Count -eq 2) { $data[$kv[0]] = $kv[1] }
                    }
                    $script:EngCheckData = $data

                    # Find and update engine cards
                    foreach ($brd in $pnlEng.Children) {
                        if ($brd -is [System.Windows.Controls.Border]) {
                            $sp = $brd.Child
                            if ($sp -is [System.Windows.Controls.StackPanel]) {
                                # Find version label and update text
                                foreach ($child in $sp.Children) {
                                    if ($child -is [System.Windows.Controls.TextBlock]) {
                                        if ($child.Tag -eq 'latest_ver') {
                                            if ($data.ContainsKey('PS7_Latest') -and $brd.Tag -eq 'PS7') {
                                                $child.Text = "  Latest Available : v$($data['PS7_Latest'])"
                                                $child.Foreground = '#50D090'
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    EL "Version check complete. PS7 latest: $($data['PS7_Latest'])  winget: $($data['Winget_Latest'])"
                } else {
                    TL $tLine 'ENGINE'
                }
            }

            if ($script:EngCheckAR -ne $null -and $script:EngCheckAR.IsCompleted) {
                $script:EngCheckTimer.Stop()
                try { $script:EngCheckPS.EndInvoke($script:EngCheckAR) | Out-Null
                      $script:EngCheckPS.Dispose(); $script:EngCheckRS.Close(); $script:EngCheckRS.Dispose() } catch {}
                $btnCheckEngVer.IsEnabled = $true
                $btnCheckEngVer.Content   = 'Check for Updates'
            }
        })
        $script:EngCheckTimer.Start()
    })

    # - Install / Update PS7 ------
    $btnInstPS7.add_Click({
        $ans = [System.Windows.MessageBox]::Show(
            "Install or update PowerShell 7 using winget?`n`n" +
            "Equivalent to: winget install --id Microsoft.PowerShell`n`n" +
            "Continue?", 'Install PS7',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question)
        if ($ans -ne 'Yes') { return }

        $btnInstPS7.IsEnabled = $false
        EL "Starting PS7 install/update via winget..."
        TLHead "PS7 INSTALL/UPDATE"

        $script:PS7InstRS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:PS7InstRS.ApartmentState = 'MTA'; $script:PS7InstRS.ThreadOptions = 'ReuseThread'
        $script:PS7InstRS.Open()
        $script:PS7InstPS = [System.Management.Automation.PowerShell]::Create()
        $script:PS7InstPS.Runspace = $script:PS7InstRS
        [void]$script:PS7InstPS.AddScript({
            param($termQ)
            function Push([string]$m) { try { $termQ.Enqueue("[PS7INST] $m") } catch {} }
            Push "Running: winget install --id Microsoft.PowerShell --source winget --accept-package-agreements --accept-source-agreements"
            try {
                $out = & winget.exe install --id Microsoft.PowerShell --source winget --accept-package-agreements --accept-source-agreements 2>&1
                foreach ($ln in $out) { Push $ln.ToString() }
                Push "DONE"
            } catch { Push "ERROR: $_" }
        })
        [void]$script:PS7InstPS.AddParameters(@{ termQ = $termQ })
        $script:PS7InstAR = $script:PS7InstPS.BeginInvoke()

        $script:PS7InstTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:PS7InstTimer.Interval = [System.TimeSpan]::FromMilliseconds(500)
        $script:PS7InstTimer.add_Tick({
            $tLine = $null
            while ($termQ.TryDequeue([ref]$tLine)) {
                if ($tLine -match '^\[PS7INST\] (.+)$') { EL $Matches[1] }
                else { TL $tLine 'PS7' }
            }
            if ($script:PS7InstAR -ne $null -and $script:PS7InstAR.IsCompleted) {
                $script:PS7InstTimer.Stop()
                try { $script:PS7InstPS.EndInvoke($script:PS7InstAR) | Out-Null
                      $script:PS7InstPS.Dispose(); $script:PS7InstRS.Close(); $script:PS7InstRS.Dispose() } catch {}
                $btnInstPS7.IsEnabled = $true
                EL "PS7 install/update process complete. Restart app to detect new version."
            }
        })
        $script:PS7InstTimer.Start()
    })

    # - Install WMF 5.1 -----------
    $btnInstPS5.add_Click({
        $urls = @(
            'https://www.microsoft.com/en-us/download/details.aspx?id=54616',
            'https://aka.ms/wmf51download'
        )
        $msg = "Windows Management Framework 5.1`n`n" +
               "Built into Windows 10/11 - already present.`n" +
               "For Windows 7/8.1/2012R2, download from Microsoft.`n`n" +
               "Opening Microsoft Download Center..."
        [System.Windows.MessageBox]::Show($msg, 'WMF 5.1') | Out-Null
        $opened = $false
        foreach ($url in $urls) {
            try { Start-Process $url; $opened = $true; break } catch {}
        }
        if (-not $opened) {
            [System.Windows.MessageBox]::Show(
                "Could not open browser. Visit manually:`nhttps://www.microsoft.com/download/details.aspx?id=54616",
                'WMF 5.1') | Out-Null
        }
    })

    # - Check winget --------------
    $btnCheckWinget.add_Click({
        $btnCheckWinget.IsEnabled = $false
        EL "Checking winget version..."
        try {
            $v = & winget.exe --version 2>$null
            if ($v) {
                EL "winget installed: $($v.Trim())"
                TL "winget: $($v.Trim())"
                # Get latest from GitHub (reuse check engine)
                $btnCheckEngVer.RaiseEvent([System.Windows.RoutedEventArgs]::new(
                    [System.Windows.Controls.Button]::ClickEvent))
            } else {
                EL "winget not found. It comes with App Installer from the Microsoft Store."
                [System.Windows.MessageBox]::Show(
                    "winget not found in PATH.`n`nwinget comes with 'App Installer' from the Microsoft Store.`n`nOpening Microsoft Store...",
                    'winget Not Found') | Out-Null
                try { Start-Process 'ms-windows-store://pdp/?ProductId=9NBLGGH4NNS1' } catch {
                    Start-Process 'https://aka.ms/getwinget'
                }
            }
        } catch { EL "winget check error: $_" }
        $btnCheckWinget.IsEnabled = $true
    })

    # - Install / Update winget ---
    $btnInstWinget.add_Click({
        $ans = [System.Windows.MessageBox]::Show(
            "Install or update winget (App Installer)?`n`n" +
            "Method: winget upgrade --id Microsoft.AppInstaller`n`n" +
            "If winget is not installed, this will open the Microsoft Store.`n`nContinue?",
            'Install winget',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question)
        if ($ans -ne 'Yes') { return }

        $btnInstWinget.IsEnabled = $false
        EL "Attempting winget self-upgrade..."
        TLHead "WINGET INSTALL/UPDATE"

        $script:WingetRS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:WingetRS.ApartmentState = 'MTA'; $script:WingetRS.ThreadOptions = 'ReuseThread'
        $script:WingetRS.Open()
        $script:WingetPS = [System.Management.Automation.PowerShell]::Create()
        $script:WingetPS.Runspace = $script:WingetRS
        [void]$script:WingetPS.AddScript({
            param($termQ)
            function Push([string]$m) { try { $termQ.Enqueue("[WINGET] $m") } catch {} }
            try {
                $v = & winget.exe --version 2>$null
                if ($v) {
                    Push "Current winget: $($v.Trim())"
                    Push "Running: winget upgrade --id Microsoft.AppInstaller"
                    $out = & winget.exe upgrade --id Microsoft.AppInstaller --accept-package-agreements --accept-source-agreements 2>&1
                    foreach ($ln in $out) { Push $ln.ToString() }
                } else {
                    Push "winget not installed. Opening Microsoft Store..."
                    $termQ.Enqueue("[WINGET_STORE]")
                }
                Push "DONE"
            } catch { Push "ERROR: $_" }
        })
        [void]$script:WingetPS.AddParameters(@{ termQ = $termQ })
        $script:WingetAR = $script:WingetPS.BeginInvoke()

        $script:WingetTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:WingetTimer.Interval = [System.TimeSpan]::FromMilliseconds(400)
        $script:WingetTimer.add_Tick({
            $tLine = $null
            while ($termQ.TryDequeue([ref]$tLine)) {
                if ($tLine -eq '[WINGET_STORE]') {
                    try { Start-Process 'ms-windows-store://pdp/?ProductId=9NBLGGH4NNS1' } catch {}
                } elseif ($tLine -match '^\[WINGET\] (.+)$') { EL $Matches[1] }
                else { TL $tLine 'WINGET' }
            }
            if ($script:WingetAR -ne $null -and $script:WingetAR.IsCompleted) {
                $script:WingetTimer.Stop()
                try { $script:WingetPS.EndInvoke($script:WingetAR) | Out-Null
                      $script:WingetPS.Dispose(); $script:WingetRS.Close(); $script:WingetRS.Dispose() } catch {}
                $btnInstWinget.IsEnabled = $true
                EL "winget operation complete."
            }
        })
        $script:WingetTimer.Start()
    })

    # - Log tab ------------------
    $btnClearTerm.add_Click({
        $script:TermLines.Clear()
        $txtTerm.Text = ''
        TL 'Terminal cleared.'
    })

    $btnClearLog.add_Click({
        $logLines.Clear()
        try { Clear-Content -Path $logFile } catch {}
        $txtMini.Text = ''; $txtLog.Text = ''
        ML 'Log cleared.'
    })
    $btnOpenLog.add_Click({
        try { Start-Process notepad.exe $logFile } catch {}
    })
    $btnReloadLog.add_Click({ ReloadLogView })

    $btnSetLogPath = G 'btnSetLogPath'
    $btnSetLogPath.add_Click({
        Add-Type -AssemblyName System.Windows.Forms -EA SilentlyContinue
        $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
        $fbd.Description = "Select folder for log file"
        $fbd.SelectedPath = [System.IO.Path]::GetDirectoryName($logFile)
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $newFolder = $fbd.SelectedPath
            $logName   = [System.IO.Path]::GetFileName($script:LogFilePath)
            $newPath   = [System.IO.Path]::Combine($newFolder, $logName)
            try {
                $existing = Get-Content $script:LogFilePath -Raw -EA SilentlyContinue
                if ($existing) { Set-Content $newPath -Value $existing -Encoding UTF8 }
            } catch {}
            $script:LogFilePath = $newPath
            $logFile = $newPath
            $lblLogPath.Text = "Log: $newPath"
            ML "[OK] Log path changed to: $newPath"
        }
    })

    # - Init -------------------
    ML "Module Manager v7.0 ready. $($allRows.Count) modules in catalog."
    ML "Select engine + scope, then click [Refresh Status]."
    $lblStatus.Text = "Ready  |  Log: $logFile"
    $cmbEng.SelectedIndex = 0
    $e = GetEng; if ($e) { $lblExe.Text = $e.Exe }
    ReloadLogView

    # - Show window with Application -----------
    # This is the CORRECT pattern: create Application FIRST, then Run
    # We are already on a dedicated STA thread so this is safe
    # Stop all running timers when window closes to prevent post-close errors
    $win.add_Closing({
        try { if ($script:ScanTimer   -and $script:ScanTimer.IsEnabled)   { $script:ScanTimer.Stop() } }   catch {}
        try { if ($script:BatchTimer  -and $script:BatchTimer.IsEnabled)  { $script:BatchTimer.Stop() } }  catch {}
        try { if ($script:OldVerTimer -and $script:OldVerTimer.IsEnabled) { $script:OldVerTimer.Stop() } } catch {}
        try { if ($script:EngCheckTimer -and $script:EngCheckTimer.IsEnabled) { $script:EngCheckTimer.Stop() } } catch {}
        try { if ($script:PS7InstTimer  -and $script:PS7InstTimer.IsEnabled)  { $script:PS7InstTimer.Stop() } }  catch {}
        try { if ($script:WingetTimer   -and $script:WingetTimer.IsEnabled)   { $script:WingetTimer.Stop() } }   catch {}
        try { if ($script:PathsTabTimer -and $script:PathsTabTimer.IsEnabled) { $script:PathsTabTimer.Stop() } } catch {}
        try { if ($script:RepoScanRS)  { $script:RepoScanRS.Close();  $script:RepoScanRS.Dispose() } }  catch {}
        try { if ($script:OldVerRS)    { $script:OldVerRS.Close();    $script:OldVerRS.Dispose() } }    catch {}
    })

    $app = [System.Windows.Application]::new()
    $app.Run($win) | Out-Null
}

# ============================================================
#  LAUNCH GUI - PowerShell Runspace with STA apartment
# ============================================================
# Using PowerShell.Create() with an STA runspace.
# This gives the new thread a proper PS execution context.
# BeginInvoke() runs async; we Wait() on the IAsyncResult.
# ============================================================
Write-Log 'Starting GUI on STA Runspace...' 'INFO'

# Build an InitialSessionState so the runspace has access to all
# necessary types and variables
$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

# Create a Runspace with STA apartment - this is the container
$guiRS = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
$guiRS.ApartmentState = [System.Threading.ApartmentState]::STA
$guiRS.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
$guiRS.Open()

# Set variables the GUI script needs, directly in the runspace
$guiRS.SessionStateProxy.SetVariable('engines',     $Global:Engines)
$guiRS.SessionStateProxy.SetVariable('allRows',     $Global:AllRows)
$guiRS.SessionStateProxy.SetVariable('catalog',     $Global:Catalog)
$guiRS.SessionStateProxy.SetVariable('logFile',     $Global:LogFile)
$guiRS.SessionStateProxy.SetVariable('customModsFile', $Global:CustomModulesFile)
$guiRS.SessionStateProxy.SetVariable('logQ',        $Global:LogQ)
$guiRS.SessionStateProxy.SetVariable('scanQ',       $Global:ScanResults)
$guiRS.SessionStateProxy.SetVariable('batchQ',      $Global:BatchResults)
$guiRS.SessionStateProxy.SetVariable('cancelRef',   ([ref]$Global:CancelScan))
$guiRS.SessionStateProxy.SetVariable('cancelBRef',  ([ref]$Global:CancelBatch))

# Create PowerShell instance bound to our STA runspace
$guiPS = [System.Management.Automation.PowerShell]::Create()
$guiPS.Runspace = $guiRS

# Add the GUI script - variables are already set in runspace session state
# so we call the script without parameters
[void]$guiPS.AddScript($guiScript)
[void]$guiPS.AddParameters(@{
    engines    = $Global:Engines
    allRows    = $Global:AllRows
    catalog    = $Global:Catalog
    logFile    = $Global:LogFile
    customModsFile = $Global:CustomModulesFile
    logQ       = $Global:LogQ
    scanQ      = $Global:ScanResults
    batchQ     = $Global:BatchResults
    termQ      = $Global:TermQueue
    cancelRef  = ([ref]$Global:CancelScan)
    cancelBRef = ([ref]$Global:CancelBatch)
})

# Invoke synchronously on the STA runspace thread
# (the runspace uses ReuseThread, so it runs on its own STA thread)
try {
    $guiPS.Invoke() | Out-Null
    # Non-fatal GUI errors (e.g. theme/brush on close) are suppressed
    # to keep the terminal clean. Fatal errors still show via catch block.
} catch {
    Write-Host "GUI Fatal: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
} finally {
    $guiPS.Dispose()
    $guiRS.Close()
    $guiRS.Dispose()
}

Write-Log 'GUI closed.' 'INFO'
$e = $null
while ($Global:LogQ.TryDequeue([ref]$e)) { Write-Host $e }
