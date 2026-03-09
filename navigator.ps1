<#
.SYNOPSIS
    CybtekSTK SysOp Navigator - Windows Console Navigation System
.VERSION
    3.0.2
.DESCRIPTION
    External config files:
      menus.json  - all menu definitions, options, columns, on-load cmdlets
      style.json  - colors, margins, padding, highlight settings
    Global shortcuts: Alt+M=Main Menu  Alt+S=Settings  Alt+T=Exit
    ESC always returns to the previous menu.
.NOTES
    For support email s0nt3k@protonmail.com
#>

#region === VT / ANSI SETUP ===
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class NativeConsole {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
    public static void EnableVirtualTerminal() {
        IntPtr handle = GetStdHandle(-11);
        uint mode;
        if (GetConsoleMode(handle, out mode))
            SetConsoleMode(handle, mode | 0x0004);
    }
}
"@ -ErrorAction SilentlyContinue
[NativeConsole]::EnableVirtualTerminal()

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Win32Window {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    public static void MoveWindow() {
        IntPtr hwnd = GetConsoleWindow();
        SendMessage(hwnd, 0x0112, (IntPtr)0xF010, IntPtr.Zero); // WM_SYSCOMMAND / SC_MOVE — blocks until move completes
    }
}
"@ -ErrorAction SilentlyContinue

$Script:ESC = [char]27
function fmtBI([string]$t) { "$($Script:ESC)[1;3m$t$($Script:ESC)[0m" }
#endregion

#region === STATE ===
$Script:RootPath      = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
$Script:ConfigPath    = Join-Path $Script:RootPath "data\menus.json"
$Script:StylePath     = Join-Path $Script:RootPath "data\style.json"
$Script:PropsPath     = Join-Path $Script:RootPath "data\properties.json"
$Script:ExitSplashPath    = Join-Path $Script:RootPath "scripts\exitsplash.ps1"
$Script:StartupSplashPath = Join-Path $Script:RootPath "scripts\startupsplash.ps1"
$Script:PathsDataPath         = Join-Path $Script:RootPath "data\paths.json"
$Script:PathsData             = $null       # file & folder path settings (data\paths.json)
$Script:ServiceSettingsPath   = Join-Path $Script:RootPath "data\servicesettings.json"
$Script:ServiceSettings       = $null       # service mode wizard toggle states (data\servicesettings.json)
$Script:DbConfigPath          = Join-Path $Script:RootPath "data\dbconfig.json"
$Script:DbConfig              = $null       # MySQL database connection settings (data\dbconfig.json)
$Script:Version       = "3.0.0"
$Script:UserAccess    = "ADMIN"
$Script:ServiceMode          = $false
$Script:ServiceModeStartTime = $null    # datetime when service mode was activated
$Script:Hostname      = $env:COMPUTERNAME
$Script:Running       = $true
$Script:MenuStack     = [System.Collections.Generic.Stack[string]]::new()
$Script:CurrentMenuId = "main"
$Script:StatusMsg     = "Navigator ready."
$Script:Menus         = $null
$Script:Style         = $null
$Script:Props         = $null       # navigator properties (startup/termination settings)
$Script:StartTime     = [DateTime]::Now
$Script:TimestampRow  = -1          # cursor row of the live timestamp line
$Script:SelectedIndex = 0           # highlighted option index in visual nav order
$Script:LastShownMenuId = ""        # last menu rendered — used to detect menu changes
$Script:ArrowNavActive  = $false    # true only after first arrow key press; hides highlight until then
$Script:ArrowNavTime    = [DateTime]::MinValue  # timestamp of last arrow key for 7-second highlight timeout
#endregion

#region === DEFAULT CONFIGS ===
function Get-DefaultMenuConfig {
    return [pscustomobject]@{
        menus = @(
            [pscustomobject]@{
                id="main"; title="CybtekSTK SysOp Navigator Menu"; columns=2; onLoad=""
                options=@(
                    [pscustomobject]@{key="1";label="Customer Account & Service Settings";column="left";type="menu";target="customer-account";window="current"}
                    [pscustomobject]@{key="2";label="System Administration & Management";column="left";type="menu";target="system-admin";window="current"}
                    [pscustomobject]@{key="3";label="Windows Settings & Configurations";column="left";type="menu";target="windows-settings";window="current"}
                    [pscustomobject]@{key="4";label="Information Security & Privacy";column="left";type="menu";target="info-security";window="current"}
                    [pscustomobject]@{key="A";label="Software Package & Module Management";column="right";type="menu";target="software-mgmt";window="current"}
                    [pscustomobject]@{key="B";label="Miscellaneous Tools & Utilities";column="right";type="menu";target="misc-tools";window="current"}
                    [pscustomobject]@{key="C";label="CybtekSTK SysOp Navigator Settings";column="right";type="menu";target="settings";window="current"}
                    [pscustomobject]@{key="D";label="Diagnostics & Troubleshooting Menu";column="right";type="menu";target="diagnostics";window="current"}
                )
            }
            [pscustomobject]@{
                id="settings"; title="CybtekSTK SysOp Navigator Settings"; columns=2; onLoad=""
                options=@(
                    [pscustomobject]@{key="1";label="Manage Navigation Menus";column="left";type="menu";target="_manage-menus";window="current"}
                    [pscustomobject]@{key="2";label="Manage Menu Options";column="left";type="menu";target="_manage-options";window="current"}
                    [pscustomobject]@{key="3";label="Manage Colors and Styling";column="left";type="menu";target="_colors-styling";window="current"}
                    [pscustomobject]@{key="4";label="File & Folder Path Management";column="left";type="menu";target="_file-paths";window="current"}
                    [pscustomobject]@{key="A";label="Termination & Startup Settings";column="right";type="menu";target="_termination-startup";window="current"}
                    [pscustomobject]@{key="B";label="Navigator Property Settings";column="right";type="run";target="";window="current"}
                    [pscustomobject]@{key="C";label="Configure Database Connection";column="right";type="function";target="Show-DatabaseConfig";window="current"}
                )
            }
            [pscustomobject]@{
                id="customer-account"; title="Customer Account & Service Settings"; columns=1; onLoad=""
                options=@(
                    [pscustomobject]@{key="1";label="Enable Technical Service Mode";column="left";type="function";target="Invoke-EnableServiceMode";window="current"}
                )
            }
            [pscustomobject]@{id="system-admin";title="System Administration & Management";columns=1;onLoad="";options=@()}
            [pscustomobject]@{
                id="windows-settings"; title="Windows Settings & Configurations"; columns=2; onLoad=""
                options=@(
                    [pscustomobject]@{key="1";label="Windows Update Settings";column="left";type="run";target="ms-settings:windowsupdate";window="current"}
                    [pscustomobject]@{key="2";label="Privacy & Security Settings";column="left";type="run";target="ms-settings:privacy";window="current"}
                    [pscustomobject]@{key="3";label="Windows System Settings";column="left";type="run";target="ms-settings:about";window="current"}
                    [pscustomobject]@{key="4";label="Network & Internet Settings";column="left";type="run";target="ms-settings:network";window="current"}
                    [pscustomobject]@{key="5";label="User Accounts Settings";column="left";type="run";target="ms-settings:accounts";window="current"}
                    [pscustomobject]@{key="6";label="Personalization Settings";column="left";type="run";target="ms-settings:personalization";window="current"}
                    [pscustomobject]@{key="7";label="Application Settings";column="left";type="run";target="ms-settings:appsfeatures";window="current"}
                    [pscustomobject]@{key="8";label="Bluetooth & Devices Settings";column="left";type="run";target="ms-settings:bluetooth";window="current"}
                    [pscustomobject]@{key="A";label="Enable Service Mode Settings";column="right";type="menu";target="_enable-service-mode";window="current"}
                    [pscustomobject]@{key="B";label="Disable Service Mode Settings";column="right";type="menu";target="_disable-service-mode";window="current"}
                    [pscustomobject]@{key="C";label="Scheduled Automation Settings";column="right";type="run";target="";window="current"}
                )
            }
            [pscustomobject]@{id="info-security";title="Information Security & Privacy";columns=1;onLoad="";options=@()}
            [pscustomobject]@{id="software-mgmt";title="Software Package & Module Management";columns=1;onLoad="";options=@()}
            [pscustomobject]@{id="misc-tools";title="Miscellaneous Tools & Utilities";columns=1;onLoad="";options=@()}
            [pscustomobject]@{id="diagnostics";title="Diagnostics & Troubleshooting Menu";columns=1;onLoad="";options=@()}
        )
    }
}

function Get-DefaultPropsConfig {
    return [pscustomobject]@{
        StartupMenu             = "Menu"
        StartupDelayMs          = 4500
        ShowStartupSplash       = $false
        StartupSplashMs         = 4000
        StartupScriptDelay      = "Disabled"
        ShowTerminationSplash   = $true
        TerminationSplashMs     = 5000
        TerminateActiveProcess  = $true
    }
}

function Get-DefaultPathsConfig {
    return [pscustomobject]@{
        MenuConfigPath  = Join-Path $Script:RootPath "data\menus.json"
        StyleConfigPath = Join-Path $Script:RootPath "data\style.json"
        PropsConfigPath = Join-Path $Script:RootPath "data\properties.json"
        ExitSplashPath  = Join-Path $Script:RootPath "scripts\exitsplash.ps1"
        DataFolderPath  = Join-Path $Script:RootPath "data"
    }
}

function Get-DefaultServiceSettings {
    # Empty sub-objects — hardcoded defaults are used until the user saves for the first time
    return [pscustomobject]@{
        enable  = [pscustomobject]@{}
        disable = [pscustomobject]@{}
    }
}

function Get-DefaultDbConfig {
    return [pscustomobject]@{
        Host     = "localhost"
        Port     = 3306
        Database = ""
        Username = "root"
        Password = ""
    }
}

function Get-DefaultStyleConfig {
    return [pscustomobject]@{
        colors = [pscustomobject]@{
            BorderColor        = "Cyan"
            BracketHyphenColor = "DarkGray"
            TriggerKeyColor    = "Cyan"
            MenuOptionText     = "DarkYellow"
            MenuTitleText      = "Yellow"
            LiveTimestamp      = "White"
            ConfigValue        = "Red"
            FooterHint         = "DarkYellow"
            StatusMessage      = "White"
        }
        layout = [pscustomobject]@{
            BorderVerticalMargin   = 1
            BorderHorizontalMargin = 1
            BorderPadding          = 2
            HighlightBackground    = "DarkBlue"
            HighlightForeground    = "White"
            ServiceModeBorder      = "Green"
        }
        hexColors = [pscustomobject]@{}
    }
}
#endregion

#region === CONFIG I/O ===
function Initialize-Config {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $dataDir = Join-Path $Script:RootPath "data"
    if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }
    if (Test-Path $Script:ConfigPath) {
        try { $Script:Menus = Get-Content $Script:ConfigPath -Raw | ConvertFrom-Json }
        catch { $Script:Menus = Get-DefaultMenuConfig; Save-MenuConfig }
    } else {
        $Script:Menus = Get-DefaultMenuConfig; Save-MenuConfig
    }
    if (Test-Path $Script:StylePath) {
        try { $Script:Style = Get-Content $Script:StylePath -Raw | ConvertFrom-Json }
        catch { $Script:Style = Get-DefaultStyleConfig; Save-StyleConfig }
    } else {
        $Script:Style = Get-DefaultStyleConfig; Save-StyleConfig
    }
    if (Test-Path $Script:PropsPath) {
        try { $Script:Props = Get-Content $Script:PropsPath -Raw | ConvertFrom-Json }
        catch { $Script:Props = Get-DefaultPropsConfig; Save-PropsConfig }
    } else {
        $Script:Props = Get-DefaultPropsConfig; Save-PropsConfig
    }
    if (Test-Path $Script:PathsDataPath) {
        try { $Script:PathsData = Get-Content $Script:PathsDataPath -Raw | ConvertFrom-Json }
        catch { $Script:PathsData = Get-DefaultPathsConfig; Save-PathsConfig }
    } else {
        $Script:PathsData = Get-DefaultPathsConfig; Save-PathsConfig
    }
    if (Test-Path $Script:ServiceSettingsPath) {
        try { $Script:ServiceSettings = Get-Content $Script:ServiceSettingsPath -Raw | ConvertFrom-Json }
        catch { $Script:ServiceSettings = Get-DefaultServiceSettings; Save-ServiceSettings }
    } else {
        $Script:ServiceSettings = Get-DefaultServiceSettings; Save-ServiceSettings
    }
    if (Test-Path $Script:DbConfigPath) {
        try { $Script:DbConfig = Get-Content $Script:DbConfigPath -Raw | ConvertFrom-Json }
        catch { $Script:DbConfig = Get-DefaultDbConfig; Save-DbConfig }
    } else {
        $Script:DbConfig = Get-DefaultDbConfig; Save-DbConfig
    }
    # Migrate: update settings menu [C] Configure Database Connection from type="run" to type="function"
    $settingsMenu = @($Script:Menus.menus) | Where-Object { $_.id -eq "settings" } | Select-Object -First 1
    if ($settingsMenu -and $settingsMenu.options) {
        $dbOpt = @($settingsMenu.options) | Where-Object { $_.key -eq "C" -and $_.type -eq "run" } | Select-Object -First 1
        if ($dbOpt) { $dbOpt.type = "function"; $dbOpt.target = "Show-DatabaseConfig"; Save-MenuConfig }
    }
    # Migrate: update customer-account "Enable Technical Service Mode" from type="run" to type="function"
    $caMenu = @($Script:Menus.menus) | Where-Object { $_.id -eq "customer-account" } | Select-Object -First 1
    if ($caMenu -and $caMenu.options) {
        $caOpt = @($caMenu.options) | Where-Object { $_.key -eq "1" -and $_.label -like "*Enable Technical Service Mode*" -and $_.type -eq "run" } | Select-Object -First 1
        if ($caOpt) { $caOpt.type = "function"; $caOpt.target = "Invoke-EnableServiceMode"; Save-MenuConfig }
    }
    # Migrate: update windows-settings service mode options from legacy type="run" to type="menu"
    $wsMenu = @($Script:Menus.menus) | Where-Object { $_.id -eq "windows-settings" } | Select-Object -First 1
    if ($wsMenu -and $wsMenu.options) {
        $migrated = $false
        $optA = @($wsMenu.options) | Where-Object { $_.key -eq "A" -and $_.label -like "*Enable Service Mode*" -and $_.type -eq "run" } | Select-Object -First 1
        $optB = @($wsMenu.options) | Where-Object { $_.key -eq "B" -and $_.label -like "*Disable Service Mode*" -and $_.type -eq "run" } | Select-Object -First 1
        if ($optA) { $optA.type = "menu"; $optA.target = "_enable-service-mode";  $migrated = $true }
        if ($optB) { $optB.type = "menu"; $optB.target = "_disable-service-mode"; $migrated = $true }
        if ($migrated) { Save-MenuConfig }
    }
    # On every startup: service mode is always OFF — reset label if it was left as "Disable Technical Service Mode"
    $caMenuReset = @($Script:Menus.menus) | Where-Object { $_.id -eq "customer-account" } | Select-Object -First 1
    if ($caMenuReset -and $caMenuReset.options) {
        $optReset = @($caMenuReset.options) | Where-Object { $_.key -eq "1" -and $_.label -eq "Disable Technical Service Mode" } | Select-Object -First 1
        if ($optReset) { $optReset.label = "Enable Technical Service Mode"; Save-MenuConfig }
    }
    # Resume service mode if a reboot-persistence file exists (written by [F] Launch Script After Reboot)
    $resumePath = Join-Path $Script:RootPath "data\servicemode-resume.json"
    if (Test-Path $resumePath) {
        try {
            $rd = Get-Content $resumePath -Raw | ConvertFrom-Json
            if ($rd.StartTime) {
                $Script:ServiceModeStartTime = [DateTime]::Parse($rd.StartTime)
                $Script:ServiceMode = $true
                $caR = @($Script:Menus.menus) | Where-Object { $_.id -eq "customer-account" } | Select-Object -First 1
                if ($caR -and $caR.options) {
                    $optR = @($caR.options) | Where-Object { $_.key -eq "1" } | Select-Object -First 1
                    if ($optR) { $optR.label = "Disable Technical Service Mode"; Save-MenuConfig }
                }
            }
        } catch {}
    }
}

function Save-MenuConfig     { $Script:Menus           | ConvertTo-Json -Depth 10 | Set-Content $Script:ConfigPath          -Encoding UTF8 }
function Save-StyleConfig    { $Script:Style           | ConvertTo-Json -Depth 10 | Set-Content $Script:StylePath           -Encoding UTF8 }
function Save-ServiceSettings{ $Script:ServiceSettings | ConvertTo-Json -Depth 5  | Set-Content $Script:ServiceSettingsPath -Encoding UTF8 }
function Save-DbConfig {
    $dir = Split-Path $Script:DbConfigPath -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $Script:DbConfig | ConvertTo-Json -Depth 5 | Set-Content $Script:DbConfigPath -Encoding UTF8
}
function Save-PropsConfig  { $Script:Props  | ConvertTo-Json -Depth 10 | Set-Content $Script:PropsPath  -Encoding UTF8 }
function Save-PathsConfig  {
    $dir = Split-Path $Script:PathsDataPath -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $Script:PathsData | ConvertTo-Json -Depth 10 | Set-Content $Script:PathsDataPath -Encoding UTF8
}

function Get-MenuById([string]$Id) {
    return $Script:Menus.menus | Where-Object { $_.id -eq $Id } | Select-Object -First 1
}
function Get-VisualNavOrder($Menu) {
    # Column-major: all left options then all right options (intuitive Up/Down within each column)
    $opts = if ($Menu.options) { @($Menu.options) } else { @() }
    if ([int]$Menu.columns -ne 2) { return $opts }
    $leftO  = @($opts | Where-Object { $_.column -eq "left"  })
    $rightO = @($opts | Where-Object { $_.column -eq "right" })
    return @($leftO) + @($rightO)
}
function Get-StyleColor([string]$n) {   # Get color from style
    $v = $Script:Style.colors.$n
    if ([string]::IsNullOrEmpty($v)) { return "White" }
    return $v
}
function Get-LayoutVal([string]$n) { return $Script:Style.layout.$n }  # Get layout value

function Get-ServiceModeTimer {
    if (-not $Script:ServiceMode -or $null -eq $Script:ServiceModeStartTime) { return "" }
    $elapsed = [DateTime]::Now - $Script:ServiceModeStartTime
    $h = [int]$elapsed.TotalHours
    $m = $elapsed.Minutes
    $s = $elapsed.Seconds
    return ("SERVICE MODE TIMER: {0:D2}h {1:D2}m {2:D2}s" -f $h, $m, $s)
}

function Write-ConsoleTimerValue([string]$Val) {
    # Val = "##h ##m ##s" — digits White, h/m/s letters Green
    try { [Console]::ForegroundColor = [ConsoleColor]"White" } catch {}; [Console]::Write($Val.Substring(0,2))
    try { [Console]::ForegroundColor = [ConsoleColor]"Green" } catch {}; [Console]::Write("h ")
    try { [Console]::ForegroundColor = [ConsoleColor]"White" } catch {}; [Console]::Write($Val.Substring(4,2))
    try { [Console]::ForegroundColor = [ConsoleColor]"Green" } catch {}; [Console]::Write("m ")
    try { [Console]::ForegroundColor = [ConsoleColor]"White" } catch {}; [Console]::Write($Val.Substring(8,2))
    try { [Console]::ForegroundColor = [ConsoleColor]"Green" } catch {}; [Console]::Write("s")
}
#endregion

#region === DRAWING ENGINE ===
function Write-Color {   # Write-Colored helper
    param([string]$Text,[string]$Fg="White",[string]$Bg="",[switch]$NL)
    $of = [Console]::ForegroundColor
    $ob = [Console]::BackgroundColor
    try { [Console]::ForegroundColor = [ConsoleColor]$Fg } catch {}
    if ($Bg -ne "") { try { [Console]::BackgroundColor = [ConsoleColor]$Bg } catch {} }
    if ($NL) { [Console]::Write($Text) } else { [Console]::WriteLine($Text) }
    [Console]::ForegroundColor = $of
    [Console]::BackgroundColor = $ob
}

function Get-InnerWidth { return [Console]::WindowWidth - 2 }

function CenterText([string]$raw, [string]$ansi, [int]$w) {
    $len = $raw.Length
    $lp  = [Math]::Floor(($w - $len) / 2)
    $rp  = $w - $len - $lp
    return (' ' * [Math]::Max(0,$lp)) + $ansi + (' ' * [Math]::Max(0,$rp))
}

function Write-InfoBar {
    $iw  = Get-InnerWidth
    $bc  = Get-StyleColor "BorderColor"
    $cv  = Get-StyleColor "ConfigValue"
    $sm  = if ($Script:ServiceMode) { "ON" } else { "OFF" }
    $smc = if ($Script:ServiceMode) { Get-LayoutVal "ServiceModeBorder" } else { "White" }

    Write-Color ("┌" + "─" * $iw + "┐") $bc
    [Console]::Write("│ ")
    Write-Color "VERSION: " $bc -NL; Write-Color $Script:Version $cv -NL
    Write-Color " │ USER-ACCESS: " $bc -NL; Write-Color $Script:UserAccess $cv -NL
    Write-Color " │ SERVICE-MODE: " $bc -NL; Write-Color $sm $smc -NL
    Write-Color " │ LOCALHOST: " $bc -NL; Write-Color $Script:Hostname $cv -NL
    $plain = " VERSION: $($Script:Version) | USER-ACCESS: $($Script:UserAccess) | SERVICE-MODE: $sm | LOCALHOST: $($Script:Hostname)"
    $pad = $iw - $plain.Length - 1
    if ($pad -gt 0) { [Console]::Write(' ' * $pad) }
    Write-Color " │" $bc
    Write-Color ("└" + "─" * $iw + "┘") $bc
}

function Write-MenuBorder([string]$Title,[int]$Cols,[array]$Opts,[string]$SelId="") {
    $iw    = Get-InnerWidth
    $bc    = Get-StyleColor "BorderColor"
    $tc    = Get-StyleColor "MenuTitleText"
    $oc    = Get-StyleColor "MenuOptionText"
    $kc    = Get-StyleColor "TriggerKeyColor"
    $bk    = Get-StyleColor "BracketHyphenColor"
    $tsc   = Get-StyleColor "LiveTimestamp"
    $hlBg  = Get-LayoutVal "HighlightBackground"
    $hlFg  = Get-LayoutVal "HighlightForeground"
    $vMarg = [int](Get-LayoutVal "BorderVerticalMargin")
    $pad   = [int](Get-LayoutVal "BorderPadding")

    # Top border + title
    Write-Color ("╔" + "═" * $iw + "╗") $bc
    $titleAnsi = fmtBI $Title
    $centered  = CenterText $Title $titleAnsi $iw
    [Console]::Write("║"); Write-Color $centered $tc -NL; Write-Color "║" $bc
    Write-Color ("╠══" + "─" * ($iw - 4) + "══╣") $bc

    # Vertical top margin
    for ($i=0;$i -lt $vMarg;$i++) { Write-Color ("║" + " " * $iw + "║") $bc }

    if ($Cols -eq 2) {
        $colW   = [Math]::Floor(($iw - 1) / 2)
        $leftO  = @($Opts | Where-Object { $_.column -eq "left"  })
        $rightO = @($Opts | Where-Object { $_.column -eq "right" })
        $rows   = [Math]::Max($leftO.Count, $rightO.Count)
        for ($r=0;$r -lt $rows;$r++) {
            [Console]::Write("║")
            if ($r -lt $leftO.Count) {
                $o=$leftO[$r]; $tl=1+$o.key.Length+4+$o.label.Length; $rem=$colW-$pad-$tl
                if ($SelId -ne "" -and "$($o.key)|$($o.column)" -eq $SelId) {
                    $of=[Console]::ForegroundColor; $ob=[Console]::BackgroundColor
                    try{[Console]::ForegroundColor=[ConsoleColor]$hlFg}catch{}
                    try{[Console]::BackgroundColor=[ConsoleColor]$hlBg}catch{}
                    [Console]::Write(" " * $pad + "[$($o.key)] - $($o.label)")
                    if($rem -gt 0){[Console]::Write(" " * $rem)}
                    [Console]::ForegroundColor=$of; [Console]::BackgroundColor=$ob
                } else {
                    [Console]::Write(" " * $pad)
                    Write-Color "[" $bk -NL; Write-Color $o.key $kc -NL; Write-Color "] - " $bk -NL; Write-Color $o.label $oc -NL
                    if($rem -gt 0){[Console]::Write(" " * $rem)}
                }
            } else { [Console]::Write(" " * $colW) }
            Write-Color "│" $bk -NL
            if ($r -lt $rightO.Count) {
                $o=$rightO[$r]; $tl=1+$o.key.Length+4+$o.label.Length; $rem=($iw-$colW-1)-$pad-$tl
                if ($SelId -ne "" -and "$($o.key)|$($o.column)" -eq $SelId) {
                    $of=[Console]::ForegroundColor; $ob=[Console]::BackgroundColor
                    try{[Console]::ForegroundColor=[ConsoleColor]$hlFg}catch{}
                    try{[Console]::BackgroundColor=[ConsoleColor]$hlBg}catch{}
                    [Console]::Write(" " * $pad + "[$($o.key)] - $($o.label)")
                    if($rem -gt 0){[Console]::Write(" " * $rem)}
                    [Console]::ForegroundColor=$of; [Console]::BackgroundColor=$ob
                } else {
                    [Console]::Write(" " * $pad)
                    Write-Color "[" $bk -NL; Write-Color $o.key $kc -NL; Write-Color "] - " $bk -NL; Write-Color $o.label $oc -NL
                    if($rem -gt 0){[Console]::Write(" " * $rem)}
                }
            } else { [Console]::Write(" " * ($iw-$colW-1)) }
            Write-Color "║" $bc
        }
    } else {
        foreach ($o in $Opts) {
            [Console]::Write("║")
            $tl=1+$o.key.Length+4+$o.label.Length; $rem=$iw-$pad-$tl
            if ($SelId -ne "" -and "$($o.key)|$($o.column)" -eq $SelId) {
                $of=[Console]::ForegroundColor; $ob=[Console]::BackgroundColor
                try{[Console]::ForegroundColor=[ConsoleColor]$hlFg}catch{}
                try{[Console]::BackgroundColor=[ConsoleColor]$hlBg}catch{}
                [Console]::Write(" " * $pad + "[$($o.key)] - $($o.label)")
                if($rem -gt 0){[Console]::Write(" " * $rem)}
                [Console]::ForegroundColor=$of; [Console]::BackgroundColor=$ob
            } else {
                [Console]::Write(" " * $pad)
                Write-Color "[" $bk -NL; Write-Color $o.key $kc -NL; Write-Color "] - " $bk -NL; Write-Color $o.label $oc -NL
                if($rem -gt 0){[Console]::Write(" " * $rem)}
            }
            Write-Color "║" $bc
        }
    }

    # Vertical bottom margin
    for ($i=0;$i -lt $vMarg;$i++) { Write-Color ("║" + " " * $iw + "║") $bc }

    # Timestamp row
    Write-Color ("╠══" + "─" * ($iw - 4) + "══╣") $bc
    $Script:TimestampRow = [Console]::CursorTop   # record for live updates
    [Console]::Write("║  ")
    $ts    = Get-Date -Format "dddd MMMM dd, yyyy  hh:mm:ss tt"
    $timer = Get-ServiceModeTimer
    Write-Color $ts $tsc -NL
    if ($timer -ne "") {
        $rem = $iw - 2 - $ts.Length - $timer.Length - 2; if ($rem -gt 0) { [Console]::Write(" " * $rem) }
        $of = [Console]::ForegroundColor
        try { [Console]::ForegroundColor = [ConsoleColor]"Green" } catch {}; [Console]::Write($timer.Substring(0,20))
        Write-ConsoleTimerValue $timer.Substring(20)
        [Console]::ForegroundColor = $of
        [Console]::Write("  ")
    } else {
        $rem = $iw - 2 - $ts.Length; if($rem -gt 0){[Console]::Write(" " * $rem)}
    }
    Write-Color "║" $bc
    Write-Color ("╚" + "═" * $iw + "╝") $bc
}

function Write-Footer([string]$Status) {
    if (-not $Status) { $Status = $Script:StatusMsg }
    $iw  = Get-InnerWidth
    $bc  = Get-StyleColor "BorderColor"
    $hc  = Get-StyleColor "FooterHint"
    $sc  = Get-StyleColor "StatusMessage"
    Write-Color ("╠══" + "─" * ($iw - 4) + "══╣") $bc
    [Console]::Write("║  "); Write-Color $Status $sc -NL
    $rem=$iw-2-$Status.Length; if($rem -gt 0){[Console]::Write(" " * $rem)}
    Write-Color "║" $bc
    $hint="Alt+M=Main  Alt+S=Settings  Alt+T=Exit  Esc=Back"
    [Console]::Write("║  "); Write-Color $hint $hc -NL
    $rem=$iw-2-$hint.Length; if($rem -gt 0){[Console]::Write(" " * $rem)}
    Write-Color "║" $bc
    Write-Color ("╚" + "═" * $iw + "╝") $bc
}

function Set-ConsoleSize {
    param($Menu)
    $pad   = [int](Get-LayoutVal "BorderPadding")
    $vMarg = [int](Get-LayoutVal "BorderVerticalMargin")
    $opts  = if ($Menu.options) { @($Menu.options) } else { @() }

    # Fixed-content minimum inner width
    $sm      = if ($Script:ServiceMode) { "ON" } else { "OFF" }
    $infoBar = " VERSION: $($Script:Version) | USER-ACCESS: $($Script:UserAccess) | SERVICE-MODE: $sm | LOCALHOST: $($Script:Hostname)"
    $hint    = "  Alt+M=Main  Alt+S=Settings  Alt+T=Exit  Esc=Back  "
    $tsRow   = if ($Script:ServiceMode) { "  " + (Get-Date -Format "dddd MMMM dd, yyyy  hh:mm:ss tt") + " SERVICE MODE TIMER: 00h 00m 00s  " } else { "" }
    $minIW   = [Math]::Max($infoBar.Length + 1, [Math]::Max($hint.Length, [Math]::Max($Menu.title.Length + 4, $tsRow.Length + 1)))

    # Option-content minimum inner width + row count
    if ([int]$Menu.columns -eq 2) {
        $leftO  = @($opts | Where-Object { $_.column -eq "left"  })
        $rightO = @($opts | Where-Object { $_.column -eq "right" })
        $maxL   = if ($leftO.Count)  { ($leftO  | ForEach-Object { $_.label.Length } | Measure-Object -Maximum).Maximum } else { 0 }
        $maxR   = if ($rightO.Count) { ($rightO | ForEach-Object { $_.label.Length } | Measure-Object -Maximum).Maximum } else { 0 }
        $optIW  = ($pad + 1 + 1 + 4 + $maxL + 2) + 1 + ($pad + 1 + 1 + 4 + $maxR + 2)
        $rows   = [Math]::Max($leftO.Count, $rightO.Count)
    } else {
        $maxL  = if ($opts.Count) { ($opts | ForEach-Object { $_.label.Length } | Measure-Object -Maximum).Maximum } else { 0 }
        $optIW = $pad + 1 + 1 + 4 + $maxL + $pad
        $rows  = $opts.Count
    }

    $innerW = [Math]::Max($minIW, [Math]::Max($optIW, 60))
    $width  = $innerW + 2

    # infobar(3) + blank(1) + menubox(top+title+sep+vMarg+rows+vMarg+sep+ts+bottom = 6+2*vMarg+rows) + footer(sep+status+hint+bottom = 4) + safety(2)
    $height = 3 + 1 + (6 + 2 * $vMarg + [Math]::Max($rows, 1)) + 4 + 2

    try {
        # Expand buffer if needed before resizing window, then shrink buffer to match (no scroll bars)
        $cur  = $Host.UI.RawUI.BufferSize
        $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size([Math]::Max($cur.Width,$width), [Math]::Max($cur.Height,$height))
        $Host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size($width, $height)
        $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size($width, $height)
    } catch {}
}

function Set-BuiltinConsoleSize {
    param([string]$Title, [int]$Rows, [int]$RequiredInnerW = 0)
    $sm      = if ($Script:ServiceMode) { "ON" } else { "OFF" }
    $infoBar = " VERSION: $($Script:Version) | USER-ACCESS: $($Script:UserAccess) | SERVICE-MODE: $sm | LOCALHOST: $($Script:Hostname)"
    $hint    = "  Alt+M=Main  Alt+S=Settings  Alt+T=Exit  Esc=Back  "
    $tsRow   = if ($Script:ServiceMode) { "  " + (Get-Date -Format "dddd MMMM dd, yyyy  hh:mm:ss tt") + " SERVICE MODE TIMER: 00h 00m 00s  " } else { "" }
    $minIW   = [Math]::Max($infoBar.Length + 1, [Math]::Max($hint.Length, [Math]::Max($Title.Length + 4, $tsRow.Length + 1)))
    $innerW  = [Math]::Max($minIW, [Math]::Max(60, $RequiredInnerW))
    $width   = $innerW + 2
    # InfoBar(3)+blank(1)+header(3)+blank(1)+rows+blank(1)+builtin-footer(3)+footer(4) = rows+16
    $height  = [Math]::Max($Rows + 16, 20)
    try {
        $cur = $Host.UI.RawUI.BufferSize
        $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size([Math]::Max($cur.Width,$width), [Math]::Max($cur.Height,$height))
        $Host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size($width, $height)
        $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size($width, $height)
    } catch {}
}

function Show-Menu([string]$MenuId) {
    if (-not $MenuId) { $MenuId = $Script:CurrentMenuId }
    $menu = Get-MenuById $MenuId
    if ($null -eq $menu) { $Script:StatusMsg = "Error: menu '$MenuId' not found."; $MenuId="main"; $menu=Get-MenuById "main" }
    if ($MenuId -ne $Script:LastShownMenuId) { $Script:SelectedIndex = 0; $Script:ArrowNavActive = $false }
    if (-not [string]::IsNullOrEmpty($menu.onLoad)) { try { Invoke-Expression $menu.onLoad } catch {} }
    Set-ConsoleSize $menu
    [Console]::Clear()
    Write-InfoBar
    Write-Host ""
    $opts   = if ($menu.options) { @($menu.options) } else { @() }
    $navOrd = Get-VisualNavOrder $menu
    $Script:SelectedIndex = [Math]::Min($Script:SelectedIndex, [Math]::Max($navOrd.Count - 1, 0))
    $selId  = if ($Script:ArrowNavActive -and $navOrd.Count -gt 0) { "$($navOrd[$Script:SelectedIndex].key)|$($navOrd[$Script:SelectedIndex].column)" } else { "" }
    Write-MenuBorder $menu.title ([int]$menu.columns) $opts $selId
    Write-Footer
    $Script:CurrentMenuId   = $MenuId
    $Script:LastShownMenuId = $MenuId
}

function Update-LiveTimestamp {
    if ($Script:TimestampRow -lt 0) { return }
    $iw     = Get-InnerWidth
    $bc     = Get-StyleColor "BorderColor"
    $tsc    = Get-StyleColor "LiveTimestamp"
    $ts     = Get-Date -Format "dddd MMMM dd, yyyy  hh:mm:ss tt"
    $timer  = Get-ServiceModeTimer
    $saveT  = [Console]::CursorTop
    $saveL  = [Console]::CursorLeft
    [Console]::SetCursorPosition(0, $Script:TimestampRow)
    $oldFg  = [Console]::ForegroundColor
    try { [Console]::ForegroundColor = [ConsoleColor]$bc }  catch {}; [Console]::Write("║  ")
    try { [Console]::ForegroundColor = [ConsoleColor]$tsc } catch {}; [Console]::Write($ts)
    [Console]::ForegroundColor = $oldFg
    if ($timer -ne "") {
        $rem = $iw - 2 - $ts.Length - $timer.Length - 2; if ($rem -gt 0) { [Console]::Write(" " * $rem) }
        try { [Console]::ForegroundColor = [ConsoleColor]"Green" } catch {}; [Console]::Write($timer.Substring(0,20))
        Write-ConsoleTimerValue $timer.Substring(20)
        [Console]::ForegroundColor = $oldFg
        [Console]::Write("  ")
    } else {
        $rem = $iw - 2 - $ts.Length; if ($rem -gt 0) { [Console]::Write(" " * $rem) }
    }
    try { [Console]::ForegroundColor = [ConsoleColor]$bc } catch {}; [Console]::Write("║")
    [Console]::ForegroundColor = $oldFg
    [Console]::SetCursorPosition($saveL, $saveT)
}
#endregion

#region === INPUT & NAVIGATION ===
function Push-MenuStack { $Script:MenuStack.Push($Script:CurrentMenuId) }

function Invoke-AltKey([ConsoleKeyInfo]$Key) {
    switch ($Key.Key) {
        'S' { Push-MenuStack; $Script:CurrentMenuId = "settings"; return "redraw" }
        'T' { return "exit" }
        'M' { $Script:MenuStack.Clear(); $Script:CurrentMenuId = "main"; return "redraw" }
        'B' {
            $Script:StatusMsg = "Use arrow keys to reposition the window.  Enter=Confirm  Esc=Cancel"
            Show-Menu $Script:CurrentMenuId
            try { [Win32Window]::MoveWindow() } catch {}
            $Script:StatusMsg = "Navigator ready."
            return "redraw"
        }
    }
    return "continue"
}

function Invoke-KeyPress([ConsoleKeyInfo]$Key) {
    $modAlt  = ($Key.Modifiers -band [ConsoleModifiers]::Alt) -ne 0
    $keyChar = $Key.KeyChar.ToString().ToUpper()

    if ($modAlt)                                  { return Invoke-AltKey $Key }
    if ($Key.Key -eq [ConsoleKey]::Escape) {
        if ($Script:MenuStack.Count -gt 0) { $Script:CurrentMenuId = $Script:MenuStack.Pop() }
        else                               { $Script:CurrentMenuId = "main" }
        return "redraw"
    }

    $menu = Get-MenuById $Script:CurrentMenuId
    if ($menu) {
        $navOrd = Get-VisualNavOrder $menu

        if ($Key.Key -eq [ConsoleKey]::DownArrow) {
            if ($navOrd.Count -gt 0) {
                if ($Script:ArrowNavActive) { $Script:SelectedIndex = ($Script:SelectedIndex + 1) % $navOrd.Count }
                $Script:ArrowNavActive = $true
                $Script:ArrowNavTime   = [DateTime]::Now
            }
            return "redraw"
        }
        if ($Key.Key -eq [ConsoleKey]::UpArrow) {
            if ($navOrd.Count -gt 0) {
                if ($Script:ArrowNavActive) { $Script:SelectedIndex = ($Script:SelectedIndex - 1 + $navOrd.Count) % $navOrd.Count }
                $Script:ArrowNavActive = $true
                $Script:ArrowNavTime   = [DateTime]::Now
            }
            return "redraw"
        }
        if ($Key.Key -eq [ConsoleKey]::RightArrow -and [int]$menu.columns -eq 2) {
            $lc = @($menu.options | Where-Object { $_.column -eq "left" }).Count
            if (-not $Script:ArrowNavActive) { $Script:ArrowNavActive = $true }
            elseif ($Script:SelectedIndex -lt $lc) {
                $t = $lc + $Script:SelectedIndex
                if ($t -lt $navOrd.Count) { $Script:SelectedIndex = $t }
            }
            $Script:ArrowNavTime = [DateTime]::Now
            return "redraw"
        }
        if ($Key.Key -eq [ConsoleKey]::LeftArrow -and [int]$menu.columns -eq 2) {
            $lc = @($menu.options | Where-Object { $_.column -eq "left" }).Count
            if (-not $Script:ArrowNavActive) { $Script:ArrowNavActive = $true }
            elseif ($Script:SelectedIndex -ge $lc) { $Script:SelectedIndex = $Script:SelectedIndex - $lc }
            $Script:ArrowNavTime = [DateTime]::Now
            return "redraw"
        }
        if ($Key.Key -eq [ConsoleKey]::Enter) {
            if ($navOrd.Count -gt 0 -and $Script:SelectedIndex -lt $navOrd.Count) {
                return Invoke-Option $navOrd[$Script:SelectedIndex]
            }
            return "continue"
        }

        $opt = @($menu.options) | Where-Object { $_.key.ToUpper() -eq $keyChar } | Select-Object -First 1
        if ($opt) { return (Invoke-Option $opt) }
    }
    return "continue"
}

function Invoke-Option($Opt) {
    switch ($Opt.type) {
        "menu" {
            if ($Opt.target -like "_*") { Invoke-BuiltinMenu $Opt.target; return "redraw" }
            Push-MenuStack; $Script:CurrentMenuId = $Opt.target; return "redraw"
        }
        "cmdlet" {
            if ($Opt.window -eq "new") {
                Start-Process powershell -ArgumentList "-NoExit","-Command",$Opt.target
            } else {
                try { Invoke-Expression $Opt.target | Out-Null; $Script:StatusMsg="Done: $($Opt.target)" } catch { $Script:StatusMsg="Error: $_" }
                Write-Host "`nPress Enter to continue..." -ForegroundColor DarkGray; Read-Host
            }
            return "redraw"
        }
        "function" { try { Invoke-Expression $Opt.target } catch { $Script:StatusMsg="Error: $_" }; return "redraw" }
        "script" {
            if ($Opt.window -eq "new") { Start-Process powershell -ArgumentList "-NoExit","-File",$Opt.target }
            else { try { & $Opt.target } catch { $Script:StatusMsg="Error: $_" }; Write-Host "`nPress Enter..." -ForegroundColor DarkGray; Read-Host }
            return "redraw"
        }
        "variable" {
            [Console]::Clear()
            Write-Host "Set Variable: $($Opt.target)" -ForegroundColor Cyan
            $val = Read-Host "Enter value"
            try { Set-Variable -Name ($Opt.target.TrimStart('$')) -Value $val -Scope Global; $Script:StatusMsg="$($Opt.target) = $val" } catch { $Script:StatusMsg="Error: $_" }
            return "redraw"
        }
    }
    return "continue"
}

function Invoke-BuiltinMenu([string]$Id) {
    switch ($Id) {
        "_manage-menus"          { Show-ManageMenus }
        "_manage-options"        { Show-ManageOptions }
        "_colors-styling"        { Show-ColorsAndStyling }
        "_termination-startup"   { Show-TerminationStartup }
        "_file-paths"            { Show-FilePathManagement }
        "_enable-service-mode"   { Show-EnableServiceModeSettings }
        "_disable-service-mode"  { Show-DisableServiceModeSettings }
    }
}
#endregion

#region === SHARED BUILTIN DRAW ===
function Write-BuiltinHeader([string]$Title) {
    $iw = Get-InnerWidth; $bc=Get-StyleColor "BorderColor"; $tc=Get-StyleColor "MenuTitleText"
    Write-InfoBar; Write-Host ""
    Write-Color ("╔" + "═" * $iw + "╗") $bc
    $centered = CenterText $Title (fmtBI $Title) $iw
    [Console]::Write("║"); Write-Color $centered $tc -NL; Write-Color "║" $bc
    Write-Color ("╠══" + "─" * ($iw - 4) + "══╣") $bc
}

function Write-BuiltinOptions([array]$Items) {
    $iw=Get-InnerWidth; $bc=Get-StyleColor "BorderColor"; $oc=Get-StyleColor "MenuOptionText"; $kc=Get-StyleColor "TriggerKeyColor"; $bk=Get-StyleColor "BracketHyphenColor"; $pad=4
    Write-Color ("║" + " " * $iw + "║") $bc
    foreach ($o in $Items) {
        [Console]::Write("║"); [Console]::Write(" " * $pad)
        Write-Color "[" $bk -NL; Write-Color $o.key $kc -NL; Write-Color "] - " $bk -NL; Write-Color $o.label $oc -NL
        $rem=$iw-$pad-1-$o.key.Length-4-$o.label.Length; if($rem -gt 0){[Console]::Write(" " * $rem)}
        Write-Color "║" $bc
    }
    Write-Color ("║" + " " * $iw + "║") $bc
}

function Write-BuiltinFooter([string]$Ts) {
    $iw=Get-InnerWidth; $bc=Get-StyleColor "BorderColor"; $tsc=Get-StyleColor "LiveTimestamp"
    Write-Color ("╠══" + "─" * ($iw - 4) + "══╣") $bc
    $Script:TimestampRow = [Console]::CursorTop   # record for live updates
    $timer = Get-ServiceModeTimer
    [Console]::Write("║  "); Write-Color $Ts $tsc -NL
    if ($timer -ne "") {
        $rem = $iw - 2 - $Ts.Length - $timer.Length - 2; if ($rem -gt 0) { [Console]::Write(" " * $rem) }
        $of = [Console]::ForegroundColor
        try { [Console]::ForegroundColor = [ConsoleColor]"Green" } catch {}; [Console]::Write($timer.Substring(0,20))
        Write-ConsoleTimerValue $timer.Substring(20)
        [Console]::ForegroundColor = $of
        [Console]::Write("  ")
    } else {
        $rem=$iw-2-$Ts.Length; if($rem -gt 0){[Console]::Write(" " * $rem)}
    }
    Write-Color "║" $bc
    Write-Color ("╚" + "═" * $iw + "╝") $bc
}

function Read-BuiltinKey {
    $tick = [DateTime]::Now
    while (-not [Console]::KeyAvailable) {
        Start-Sleep -Milliseconds 50
        if (([DateTime]::Now - $tick).TotalMilliseconds -ge 1000) {
            $tick = [DateTime]::Now
            Update-LiveTimestamp
        }
    }
    return [Console]::ReadKey($true)
}

function Invoke-BuiltinAlt([ConsoleKeyInfo]$Key) {
    switch ($Key.Key) {
        'T' { $Script:Running=$false; return $true }
        'M' { $Script:MenuStack.Clear(); $Script:CurrentMenuId="main"; return $true }
        'S' { $Script:CurrentMenuId="settings"; return $true }
    }
    return $false
}
#endregion

#region === MANAGE COLORS AND STYLING ===
function Show-ColorsAndStyling {
    $unsaved     = $false
    $styleBackup = $Script:Style | ConvertTo-Json -Depth 10 | ConvertFrom-Json
    while ($true) {
        $pad = 2
        $s   = $Script:Style
        $left=@(
            @{key="1";lbl="Border Color";          prop="BorderColor";        val=$s.colors.BorderColor;        sec="color"}
            @{key="2";lbl="Bracket/Hyphen Color";  prop="BracketHyphenColor"; val=$s.colors.BracketHyphenColor; sec="color"}
            @{key="3";lbl="Trigger Key Color";     prop="TriggerKeyColor";    val=$s.colors.TriggerKeyColor;    sec="color"}
            @{key="4";lbl="Menu Option Text Color";prop="MenuOptionText";     val=$s.colors.MenuOptionText;     sec="color"}
            @{key="5";lbl="Menu Title Text Color"; prop="MenuTitleText";      val=$s.colors.MenuTitleText;      sec="color"}
            @{key="6";lbl="Live Timestamp Color";  prop="LiveTimestamp";      val=$s.colors.LiveTimestamp;      sec="color"}
            @{key="7";lbl="Config Value Color";    prop="ConfigValue";        val=$s.colors.ConfigValue;        sec="color"}
            @{key="8";lbl="Footer Hint Color";     prop="FooterHint";         val=$s.colors.FooterHint;         sec="color"}
            @{key="9";lbl="Status Message Color"; prop="StatusMessage";      val=$s.colors.StatusMessage;      sec="color"}
        )
        $right=@(
            @{key="A";lbl="Border Vertical Margin";   prop="BorderVerticalMargin";   val=$s.layout.BorderVerticalMargin;   sec="int"}
            @{key="B";lbl="Border Horizontal Margin"; prop="BorderHorizontalMargin"; val=$s.layout.BorderHorizontalMargin; sec="int"}
            @{key="C";lbl="Border Padding";           prop="BorderPadding";          val=$s.layout.BorderPadding;          sec="int"}
            @{key="D";lbl="Highlight Background";     prop="HighlightBackground";    val=$s.layout.HighlightBackground;    sec="color"}
            @{key="E";lbl="Highlight Foreground";     prop="HighlightForeground";    val=$s.layout.HighlightForeground;    sec="color"}
            @{key="F";lbl="Service Mode Border";      prop="ServiceModeBorder";      val=$s.layout.ServiceModeBorder;      sec="color"}
            @{key="G";lbl="Set Color HEX Values";     prop="";                       val="Set";                            sec="hex"}
        )

        # Compute needed width from item content, then resize console
        $maxLP = ($left  | ForEach-Object { ("[$($_.key)] - $($_.lbl)" + ("."*[Math]::Max(1,24-$_.lbl.Length)) + ": $($_.val)").Length } | Measure-Object -Maximum).Maximum
        $maxRP = ($right | ForEach-Object { ("[$($_.key)] - $($_.lbl)" + ("."*[Math]::Max(1,24-$_.lbl.Length)) + ": $($_.val)").Length } | Measure-Object -Maximum).Maximum
        $reqIW = 2 * ([Math]::Max($maxLP,$maxRP) + $pad) + 1
        Set-BuiltinConsoleSize "Manage Colors and Styling" ([Math]::Max($left.Count,$right.Count)) $reqIW

        [Console]::Clear()
        Write-BuiltinHeader "Manage Colors and Styling"
        $iw   = Get-InnerWidth
        $bc   = Get-StyleColor "BorderColor"
        $oc   = Get-StyleColor "MenuOptionText"
        $kc   = Get-StyleColor "TriggerKeyColor"
        $bk   = Get-StyleColor "BracketHyphenColor"
        $cv   = Get-StyleColor "ConfigValue"
        $colW = [Math]::Floor(($iw-1)/2)

        Write-Color ("║" + " " * $iw + "║") $bc
        $rows = [Math]::Max($left.Count,$right.Count)
        for ($r=0;$r -lt $rows;$r++) {
            [Console]::Write("║"); [Console]::Write(" " * $pad)
            if ($r -lt $left.Count) {
                $it=$left[$r]; $dots="."*[Math]::Max(1,24-$it.lbl.Length)
                Write-Color "[" $bk -NL; Write-Color $it.key $kc -NL; Write-Color "] - $($it.lbl)" $oc -NL
                Write-Color $dots $bk -NL; Write-Color ": " $bk -NL; Write-Color $it.val $cv -NL
                $plain="[$($it.key)] - $($it.lbl)$($dots): $($it.val)"
                $rem=$colW-$pad-$plain.Length; if($rem -gt 0){[Console]::Write(" " * $rem)}
            } else { [Console]::Write(" " * ($colW-$pad)) }
            Write-Color "│" $bk -NL; [Console]::Write(" " * $pad)
            if ($r -lt $right.Count) {
                $it=$right[$r]; $dots="."*[Math]::Max(1,24-$it.lbl.Length)
                Write-Color "[" $bk -NL; Write-Color $it.key $kc -NL; Write-Color "] - $($it.lbl)" $oc -NL
                Write-Color $dots $bk -NL; Write-Color ": " $bk -NL; Write-Color $it.val $cv -NL
                $plain="[$($it.key)] - $($it.lbl)$($dots): $($it.val)"
                $rem=($iw-$colW-1)-$pad-$plain.Length; if($rem -gt 0){[Console]::Write(" " * $rem)}
            } else { $rem=($iw-$colW-1)-$pad; if($rem -gt 0){[Console]::Write(" " * $rem)} }
            Write-Color "║" $bc
        }
        Write-Color ("║" + " " * $iw + "║") $bc
        Write-BuiltinFooter (Get-Date -Format "dddd MMMM dd, yyyy  hh:mm:ss tt")
        $footerMsg = if ($unsaved) { "UNSAVED CHANGES  S=Save & Apply  Esc=Cancel Changes" } else { "Select a key to change its value.  Esc=Back" }
        Write-Footer $footerMsg

        $key = Read-BuiltinKey
        if (($key.Modifiers -band [ConsoleModifiers]::Alt) -ne 0) {
            if ($unsaved) { $Script:Style = $styleBackup }
            if (Invoke-BuiltinAlt $key) { return }
        }
        if ($key.Key -eq [ConsoleKey]::Escape) {
            if ($unsaved) { $Script:Style = $styleBackup }
            return
        }
        $ch = $key.KeyChar.ToString().ToUpper()

        if ($ch -eq 'S' -and $unsaved) {
            Save-StyleConfig
            $unsaved     = $false
            $styleBackup = $Script:Style | ConvertTo-Json -Depth 10 | ConvertFrom-Json
            $Script:StatusMsg = "Style settings saved and applied."
            continue
        }

        $lm = $left  | Where-Object { $_.key -eq $ch } | Select-Object -First 1
        if ($lm) { $nv=Read-ColorValue "Color for '$($lm.lbl)'" $lm.val; if($nv){$Script:Style.colors.$($lm.prop)=$nv; $unsaved=$true}; continue }

        $rm = $right | Where-Object { $_.key -eq $ch } | Select-Object -First 1
        if ($rm) {
            if ($rm.sec -eq "hex")       { if (Edit-HexColors) { $unsaved = $true } }
            elseif ($rm.sec -eq "color") { $nv=Read-ColorValue "Color for '$($rm.lbl)'" $rm.val; if($nv){$Script:Style.layout.$($rm.prop)=$nv; $unsaved=$true} }
            elseif ($rm.sec -eq "int")   { $nv=Read-IntValue   "Value for '$($rm.lbl)'" $rm.val; if($null -ne $nv){$Script:Style.layout.$($rm.prop)=$nv; $unsaved=$true} }
        }
    }
}

function Read-ColorValue([string]$Msg,[string]$Cur) {
    $valid=@("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")
    [Console]::Clear()
    Write-Host $Msg -ForegroundColor Cyan
    Write-Host "Current: $Cur" -ForegroundColor Yellow
    Write-Host "Valid:   $($valid -join ', ')" -ForegroundColor DarkGray
    $v=Read-Host "New value (blank=keep)"
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    if ($valid -contains $v) { return $v }
    Write-Host "Invalid color name." -ForegroundColor Red; Read-Host "Press Enter"; return $null
}

function Read-IntValue([string]$Msg,[object]$Cur) {
    [Console]::Clear(); Write-Host $Msg -ForegroundColor Cyan; Write-Host "Current: $Cur" -ForegroundColor Yellow
    $v=Read-Host "New value (blank=keep)"
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    if ($v -match '^\d+$') { return [int]$v }
    Write-Host "Not a valid integer." -ForegroundColor Red; Read-Host "Press Enter"; return $null
}

function Edit-HexColors {
    [Console]::Clear(); Write-Host "=== HEX Color Overrides ===" -ForegroundColor Cyan
    Write-Host "Maps a named console color to a HEX value (for true-color terminals)." -ForegroundColor DarkGray
    $name=Read-Host "Color name (e.g. Cyan, blank=cancel)"
    if ([string]::IsNullOrWhiteSpace($name)) { return $false }
    $hex=Read-Host "HEX value (e.g. #00FFFF, blank=cancel)"
    if ([string]::IsNullOrWhiteSpace($hex)) { return $false }
    $Script:Style.hexColors | Add-Member -NotePropertyName $name -NotePropertyValue $hex -Force
    $Script:StatusMsg = "HEX override: $name=$hex (unsaved)"
    return $true
}
#endregion

function Write-ServiceModeProgressBar([string]$Title, [string]$StepLabel, [int]$StepNum, [int]$TotalSteps) {
    [Console]::Clear()
    $iw    = [Math]::Max([Console]::WindowWidth - 4, 40)
    $bc    = Get-StyleColor "BorderColor"
    $tc    = Get-StyleColor "MenuTitleText"
    $cv    = Get-StyleColor "ConfigValue"
    $oc    = Get-StyleColor "MenuOptionText"
    $barFg = Get-StyleColor "TriggerKeyColor"
    Write-Color ("╔" + "═" * $iw + "╗") $bc
    [Console]::Write("║"); Write-Color (CenterText $Title (fmtBI $Title) $iw) $tc -NL; Write-Color "║" $bc
    Write-Color ("╠══" + "─" * ($iw - 4) + "══╣") $bc
    Write-Color ("║" + " " * $iw + "║") $bc
    $lbl = "  Applying: $StepLabel"
    if ($lbl.Length -gt $iw) { $lbl = $lbl.Substring(0, $iw) }
    [Console]::Write("║"); Write-Color $lbl $oc -NL; Write-Color (" " * ([Math]::Max(0, $iw - $lbl.Length)) + "║") $bc
    Write-Color ("║" + " " * $iw + "║") $bc
    $pct    = if ($TotalSteps -gt 0) { [Math]::Round(($StepNum / $TotalSteps) * 100) } else { 100 }
    $barW   = $iw - 8
    $filled = [Math]::Min($barW, [Math]::Round(($StepNum / [Math]::Max(1, $TotalSteps)) * $barW))
    $empty  = $barW - $filled
    [Console]::Write("║    ")
    $of = [Console]::ForegroundColor
    try { [Console]::ForegroundColor = [ConsoleColor]$barFg } catch {}
    [Console]::Write("█" * $filled)
    [Console]::ForegroundColor = $of
    [Console]::Write("░" * $empty)
    Write-Color "    ║" $bc
    $ctr = "  Step $StepNum of $TotalSteps  ($pct%)"
    if ($ctr.Length -gt $iw) { $ctr = $ctr.Substring(0, $iw) }
    [Console]::Write("║"); Write-Color $ctr $cv -NL; Write-Color (" " * ([Math]::Max(0, $iw - $ctr.Length)) + "║") $bc
    Write-Color ("║" + " " * $iw + "║") $bc
    Write-Color ("╚" + "═" * $iw + "╝") $bc
}

#region === SERVICE MODE SETTINGS ===
function Invoke-EnableServiceMode {
    if ($Script:ServiceMode) {
        # Toggle OFF — run disable actions then reset state
        Invoke-DisableServiceMode
        $Script:ServiceMode          = $false
        $Script:ServiceModeStartTime = $null
        $caMenu = @($Script:Menus.menus) | Where-Object { $_.id -eq "customer-account" } | Select-Object -First 1
        if ($caMenu -and $caMenu.options) {
            $opt = @($caMenu.options) | Where-Object { $_.key -eq "1" } | Select-Object -First 1
            if ($opt) { $opt.label = "Enable Technical Service Mode" }
        }
        Save-MenuConfig
        return
    }
    # Toggle ON — build ordered list of applicable steps then execute each with progress bar
    $ss     = if ($Script:ServiceSettings) { $Script:ServiceSettings.enable } else { $null }
    $gv     = { param([string]$k,[bool]$d) $v=if($ss){$ss.($k)}else{$null}; if($null -ne $v){[bool]$v}else{$d} }
    $advReg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    $expReg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer"
    $steps  = [System.Collections.ArrayList]::new()
    if (& $gv "1" $true)  { [void]$steps.Add(@{ L="Show hidden files, folders & drives";    A={ Set-ItemProperty -Path $advReg -Name "Hidden"                       -Value 1 -EA SilentlyContinue } }) }
    if (& $gv "2" $true)  { [void]$steps.Add(@{ L="Always show icons, never thumbnails";    A={ Set-ItemProperty -Path $advReg -Name "IconsOnly"                    -Value 1 -EA SilentlyContinue } }) }
    if (& $gv "3" $false) { [void]$steps.Add(@{ L="Show extensions for known file types";   A={ Set-ItemProperty -Path $advReg -Name "HideFileExt"                  -Value 0 -EA SilentlyContinue } }) }
    if (& $gv "4" $false) { [void]$steps.Add(@{ L="Show folder merge conflicts";            A={ Set-ItemProperty -Path $advReg -Name "HideMergeConflicts"           -Value 0 -EA SilentlyContinue } }) }
    if (& $gv "5" $false) { [void]$steps.Add(@{ L="Show protected system files";            A={ Set-ItemProperty -Path $advReg -Name "ShowSuperHidden"              -Value 1 -EA SilentlyContinue } }) }
    if (& $gv "6" $true)  { [void]$steps.Add(@{ L="Show sync provider notifications";       A={ Set-ItemProperty -Path $advReg -Name "ShowSyncProviderNotifications" -Value 1 -EA SilentlyContinue } }) }
    if (& $gv "7" $false) { [void]$steps.Add(@{ L="Hide recently used files";               A={ Set-ItemProperty -Path $expReg -Name "ShowRecent"                   -Value 0 -EA SilentlyContinue } }) }
    if (& $gv "8" $false) { [void]$steps.Add(@{ L="Hide frequently used folders";           A={ Set-ItemProperty -Path $expReg -Name "ShowFrequent"                 -Value 0 -EA SilentlyContinue } }) }
    if (& $gv "9" $false) { [void]$steps.Add(@{ L="Hide files from Office.com";             A={ Set-ItemProperty -Path $expReg -Name "ShowCloudFilesInQuickAccess"  -Value 0 -EA SilentlyContinue } }) }
    [void]$steps.Add(@{ L="Restarting Windows Explorer";                                     A={ try { Stop-Process -Name "explorer" -Force -EA SilentlyContinue; Start-Sleep -Milliseconds 1000; Start-Process "explorer.exe" } catch {} } })
    if (& $gv "A" $true)  { [void]$steps.Add(@{ L="Unhide C:\xTekFolder";                  A={ try { $f=Get-Item 'C:\xTekFolder' -Force -EA SilentlyContinue; if($f){$f.Attributes=$f.Attributes -band (-bnot [IO.FileAttributes]::Hidden)} } catch {} } }) }
    if (& $gv "B" $true)  { [void]$steps.Add(@{ L="Start RustDesk Service (Manual)";        A={ try { Set-Service -Name "RustDesk" -StartupType Manual -EA SilentlyContinue; Start-Service -Name "RustDesk" -EA SilentlyContinue } catch {} } }) }
    if (& $gv "C" $true)  { [void]$steps.Add(@{ L="Launch RustDesk Application as SYSTEM";  A={ try { Start-Process "\\live.sysinternals.com\tools\PSExec64.exe" -ArgumentList "-accepteula -s -i -d `"C:\Program Files\RustDesk\rustdesk.exe`"" -EA SilentlyContinue } catch {} } }) }
    if (& $gv "D" $true)  { [void]$steps.Add(@{ L="Enable Windows Auto Sign-in";            A={ try { Start-Process "\\live.sysinternals.com\tools\autologon64.exe" -EA SilentlyContinue } catch {} } }) }
    if (& $gv "E" $false) { [void]$steps.Add(@{ L="Escalate User Privileges";               A={ try { Add-LocalGroupMember -Group "Administrators" -Member $env:USERNAME -EA SilentlyContinue } catch {} } }) }
    $total = $steps.Count
    for ($i = 0; $i -lt $total; $i++) {
        Write-ServiceModeProgressBar "Enabling Technical Service Mode" $steps[$i].L ($i + 1) $total
        & $steps[$i].A
        Start-Sleep -Milliseconds 300
    }
    $Script:ServiceMode          = $true
    $Script:ServiceModeStartTime = [DateTime]::Now
    if (& $gv "F" $false) {
        $resumePath = Join-Path $Script:RootPath "data\servicemode-resume.json"
        try { @{ StartTime = $Script:ServiceModeStartTime.ToString("o") } | ConvertTo-Json | Set-Content $resumePath -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
        $scriptFile = Join-Path $Script:RootPath "navigator.ps1"
        $runCmd = "powershell.exe -NoExit -ExecutionPolicy Bypass -File `"$scriptFile`""
        try { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "NavigatorServiceMode" -Value $runCmd -ErrorAction SilentlyContinue } catch {}
    }
    $caMenu = @($Script:Menus.menus) | Where-Object { $_.id -eq "customer-account" } | Select-Object -First 1
    if ($caMenu -and $caMenu.options) {
        $opt = @($caMenu.options) | Where-Object { $_.key -eq "1" } | Select-Object -First 1
        if ($opt) { $opt.label = "Disable Technical Service Mode" }
    }
    Save-MenuConfig
    $Script:StatusMsg = "Service mode enabled."
}

function Invoke-DisableServiceMode {
    $ss     = if ($Script:ServiceSettings) { $Script:ServiceSettings.disable } else { $null }
    $gv     = { param([string]$k,[bool]$d) $v=if($ss){$ss.($k)}else{$null}; if($null -ne $v){[bool]$v}else{$d} }
    $advReg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    $expReg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer"
    $wlReg  = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    $steps  = [System.Collections.ArrayList]::new()
    if (& $gv "1" $false) { [void]$steps.Add(@{ L="Hide hidden files, folders & drives";    A={ Set-ItemProperty -Path $advReg -Name "Hidden"                       -Value 2 -EA SilentlyContinue } }) }
    if (& $gv "2" $false) { [void]$steps.Add(@{ L="Always show thumbnails, never icons";    A={ Set-ItemProperty -Path $advReg -Name "IconsOnly"                    -Value 0 -EA SilentlyContinue } }) }
    if (& $gv "3" $false) { [void]$steps.Add(@{ L="Hide extensions for known file types";   A={ Set-ItemProperty -Path $advReg -Name "HideFileExt"                  -Value 1 -EA SilentlyContinue } }) }
    if (& $gv "4" $true)  { [void]$steps.Add(@{ L="Hide folder merge conflicts";            A={ Set-ItemProperty -Path $advReg -Name "HideMergeConflicts"           -Value 1 -EA SilentlyContinue } }) }
    if (& $gv "5" $true)  { [void]$steps.Add(@{ L="Hide protected system files";            A={ Set-ItemProperty -Path $advReg -Name "ShowSuperHidden"              -Value 0 -EA SilentlyContinue } }) }
    if (& $gv "6" $false) { [void]$steps.Add(@{ L="Hide sync provider notifications";       A={ Set-ItemProperty -Path $advReg -Name "ShowSyncProviderNotifications" -Value 0 -EA SilentlyContinue } }) }
    if (& $gv "7" $true)  { [void]$steps.Add(@{ L="Show recently used files";               A={ Set-ItemProperty -Path $expReg -Name "ShowRecent"                   -Value 1 -EA SilentlyContinue } }) }
    if (& $gv "8" $true)  { [void]$steps.Add(@{ L="Show frequently used folders";           A={ Set-ItemProperty -Path $expReg -Name "ShowFrequent"                 -Value 1 -EA SilentlyContinue } }) }
    if (& $gv "9" $true)  { [void]$steps.Add(@{ L="Show files from Office.com";             A={ Set-ItemProperty -Path $expReg -Name "ShowCloudFilesInQuickAccess"  -Value 1 -EA SilentlyContinue } }) }
    [void]$steps.Add(@{ L="Restarting Windows Explorer";                                     A={ try { Stop-Process -Name "explorer" -Force -EA SilentlyContinue; Start-Sleep -Milliseconds 1000; Start-Process "explorer.exe" } catch {} } })
    if (& $gv "A" $true)  { [void]$steps.Add(@{ L="Hide C:\xTekFolder";                    A={ try { $f=Get-Item 'C:\xTekFolder' -Force -EA SilentlyContinue; if($f){$f.Attributes=$f.Attributes -bor [IO.FileAttributes]::Hidden} } catch {} } }) }
    if (& $gv "B" $true)  { [void]$steps.Add(@{ L="Disable RustDesk Service";              A={ try { Stop-Service -Name "RustDesk" -Force -EA SilentlyContinue; Set-Service -Name "RustDesk" -StartupType Disabled -EA SilentlyContinue } catch {} } }) }
    if (& $gv "C" $true)  { [void]$steps.Add(@{ L="Terminate RustDesk Application";        A={ try { Stop-Process -Name "rustdesk" -Force -EA SilentlyContinue } catch {} } }) }
    if (& $gv "D" $true)  { [void]$steps.Add(@{ L="Disable Windows Auto Sign-in";          A={ try { Set-ItemProperty -Path $wlReg -Name "AutoAdminLogon" -Value "0" -EA SilentlyContinue } catch {} } }) }
    [void]$steps.Add(@{ L="Remove launch after reboot";                                      A={
        try { Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "NavigatorServiceMode" -EA SilentlyContinue } catch {}
        $rp2 = Join-Path $Script:RootPath "data\servicemode-resume.json"
        try { if (Test-Path $rp2) { Remove-Item $rp2 -Force -EA SilentlyContinue } } catch {}
    } })
    $total = $steps.Count
    for ($i = 0; $i -lt $total; $i++) {
        Write-ServiceModeProgressBar "Disabling Technical Service Mode" $steps[$i].L ($i + 1) $total
        & $steps[$i].A
        Start-Sleep -Milliseconds 300
    }
    $Script:StatusMsg = "Service mode disabled."
}

function Show-ServiceModeWizard([string]$Title, [array]$Left, [array]$Right, [string]$SectionKey) {
    $pad = 2; $dotBase = 40
    $maxLP = ($Left  | ForEach-Object { ("[$($_.key)] - $($_.lbl)" + "." * [Math]::Max(1,$dotBase-$_.lbl.Length) + "[x]").Length } | Measure-Object -Maximum).Maximum
    $maxRP = ($Right | ForEach-Object { ("[$($_.key)] - $($_.lbl)" + "." * [Math]::Max(1,$dotBase-$_.lbl.Length) + "[x]").Length } | Measure-Object -Maximum).Maximum
    $reqIW = 2 * ([Math]::Max($maxLP,$maxRP) + $pad) + 1
    while ($true) {
        Set-BuiltinConsoleSize $Title ([Math]::Max($Left.Count,$Right.Count)) $reqIW
        [Console]::Clear()
        Write-BuiltinHeader $Title
        $iw   = Get-InnerWidth
        $bc   = Get-StyleColor "BorderColor"
        $oc   = Get-StyleColor "MenuOptionText"
        $kc   = Get-StyleColor "TriggerKeyColor"
        $bk   = Get-StyleColor "BracketHyphenColor"
        $cv   = Get-StyleColor "ConfigValue"
        $colW = [Math]::Floor(($iw-1)/2)
        Write-Color ("║" + " " * $iw + "║") $bc
        $nrows = [Math]::Max($Left.Count,$Right.Count)
        for ($r=0; $r -lt $nrows; $r++) {
            [Console]::Write("║"); [Console]::Write(" " * $pad)
            if ($r -lt $Left.Count) {
                $it = $Left[$r]; $dots = "." * [Math]::Max(1,$dotBase-$it.lbl.Length); $tv = if ($it.val) {"x"} else {"_"}
                Write-Color "[" $bk -NL; Write-Color $it.key $kc -NL; Write-Color "] - $($it.lbl)" $oc -NL
                Write-Color $dots $bk -NL; Write-Color "[" $bk -NL; Write-Color $tv $cv -NL; Write-Color "]" $bk -NL
                $plain = "[$($it.key)] - $($it.lbl)$($dots)[$tv]"
                $rem = $colW - $pad - $plain.Length; if ($rem -gt 0) { [Console]::Write(" " * $rem) }
            } else { [Console]::Write(" " * ($colW - $pad)) }
            Write-Color "│" $bk -NL; [Console]::Write(" " * $pad)
            if ($r -lt $Right.Count) {
                $it = $Right[$r]; $dots = "." * [Math]::Max(1,$dotBase-$it.lbl.Length); $tv = if ($it.val) {"x"} else {"_"}
                Write-Color "[" $bk -NL; Write-Color $it.key $kc -NL; Write-Color "] - $($it.lbl)" $oc -NL
                Write-Color $dots $bk -NL; Write-Color "[" $bk -NL; Write-Color $tv $cv -NL; Write-Color "]" $bk -NL
                $plain = "[$($it.key)] - $($it.lbl)$($dots)[$tv]"
                $rem = ($iw - $colW - 1) - $pad - $plain.Length; if ($rem -gt 0) { [Console]::Write(" " * $rem) }
            } else { $rem = ($iw - $colW - 1) - $pad; if ($rem -gt 0) { [Console]::Write(" " * $rem) } }
            Write-Color "║" $bc
        }
        Write-Color ("║" + " " * $iw + "║") $bc
        Write-BuiltinFooter (Get-Date -Format "dddd MMMM dd, yyyy  hh:mm:ss tt")
        Write-Footer "Toggle key to select.  S=Save Settings  Esc=Back"
        $key = Read-BuiltinKey
        if (($key.Modifiers -band [ConsoleModifiers]::Alt) -ne 0) { if (Invoke-BuiltinAlt $key) { return } }
        if ($key.Key -eq [ConsoleKey]::Escape) { return }
        $ch = $key.KeyChar.ToString().ToUpper()
        if ($ch -eq "S") {
            $stateHash = [ordered]@{}
            foreach ($it in ($Left + $Right)) { $stateHash[$it.key] = [bool]$it.val }
            if (-not $Script:ServiceSettings) { $Script:ServiceSettings = Get-DefaultServiceSettings }
            $Script:ServiceSettings.$SectionKey = [pscustomobject]$stateHash
            Save-ServiceSettings
            $Script:StatusMsg = "$Title settings saved."
            return
        }
        $lm = $Left  | Where-Object { $_.key -eq $ch } | Select-Object -First 1
        if ($lm) { $lm.val = -not [bool]$lm.val; continue }
        $rm = $Right | Where-Object { $_.key -eq $ch } | Select-Object -First 1
        if ($rm) { $rm.val = -not [bool]$rm.val }
    }
}

function Show-EnableServiceModeSettings {
    $left = @(
        @{key="1"; lbl="Show hidden files folders & drives";  val=$true}
        @{key="2"; lbl="Always show icons never thumbnails";  val=$true}
        @{key="3"; lbl="Show extensions 4 known file types";  val=$false}
        @{key="4"; lbl="Show folder merge conflicts";         val=$false}
        @{key="5"; lbl="Show protected system files";         val=$false}
        @{key="6"; lbl="Show sync provider notifications";    val=$true}
        @{key="7"; lbl="Hide recently used files";            val=$false}
        @{key="8"; lbl="Hide frequently used folders";        val=$false}
        @{key="9"; lbl="Hide files from Office.com";          val=$false}
    )
    $right = @(
        @{key="A"; lbl="Unhide C:\xTekFolder\>";              val=$true}
        @{key="B"; lbl="Start RustDesk Service";              val=$true}
        @{key="C"; lbl="Launch RustDesk Application";         val=$true}
        @{key="D"; lbl="Enable Windows Auto Sign-in";         val=$true}
        @{key="E"; lbl="Escalate User Privileges";            val=$false}
        @{key="F"; lbl="Launch Script After Reboot";          val=$false}
        @{key="G"; lbl="<<Undefined_Function>>";              val=$false}
        @{key="H"; lbl="<<Undefined_Function>>";              val=$false}
        @{key="I"; lbl="<<Undefined_Function>>";              val=$false}
    )
    $ss = if ($Script:ServiceSettings) { $Script:ServiceSettings.enable } else { $null }
    if ($ss) { foreach ($it in ($left + $right)) { $sv = $ss.($it.key); if ($null -ne $sv) { $it.val = [bool]$sv } } }
    Show-ServiceModeWizard "Enable Service Mode Settings" $left $right "enable"
}

function Show-DisableServiceModeSettings {
    $left = @(
        @{key="1"; lbl="Hide hidden files folders & drives";  val=$false}
        @{key="2"; lbl="Always show thumbnails never icons";  val=$false}
        @{key="3"; lbl="Hide extensions 4 known file types";  val=$false}
        @{key="4"; lbl="Hide folder merge conflicts";         val=$true}
        @{key="5"; lbl="Hide protected system files";         val=$true}
        @{key="6"; lbl="Hide sync provider notifications";    val=$false}
        @{key="7"; lbl="Show recently used files";            val=$true}
        @{key="8"; lbl="Show frequently used folders";        val=$true}
        @{key="9"; lbl="Show files from Office.com";          val=$true}
    )
    $right = @(
        @{key="A"; lbl="Hide C:\xTekFolder\>";                val=$true}
        @{key="B"; lbl="Disable RustDesk Service";             val=$true}
        @{key="C"; lbl="Terminate RustDesk Application";      val=$true}
        @{key="D"; lbl="Disable Windows Auto Sign-in";        val=$true}
        @{key="E"; lbl="Deescalate User Privileges";          val=$false}
        @{key="F"; lbl="Don't Launch Script After Reboot";    val=$false}
        @{key="G"; lbl="<<Undefined_Function>>";              val=$false}
        @{key="H"; lbl="<<Undefined_Function>>";              val=$false}
        @{key="I"; lbl="<<Undefined_Function>>";              val=$false}
    )
    $ss = if ($Script:ServiceSettings) { $Script:ServiceSettings.disable } else { $null }
    if ($ss) { foreach ($it in ($left + $right)) { $sv = $ss.($it.key); if ($null -ne $sv) { $it.val = [bool]$sv } } }
    Show-ServiceModeWizard "Disable Service Mode Settings" $left $right "disable"
}
#endregion

#region === DATABASE SETTINGS ===
function Invoke-TestDatabaseConnection {
    $cfg  = if ($Script:DbConfig) { $Script:DbConfig } else { Get-DefaultDbConfig }
    $h    = if ($cfg.Host)     { [string]$cfg.Host }     else { "localhost" }
    $p    = if ($cfg.Port)     { [int]$cfg.Port }         else { 3306 }
    $db   = if ($cfg.Database) { [string]$cfg.Database } else { "(not set)" }
    $usr  = if ($cfg.Username) { [string]$cfg.Username } else { "(not set)" }
    [Console]::Clear()
    $iw  = [Math]::Max([Console]::WindowWidth - 4, 40)
    $bc  = Get-StyleColor "BorderColor"; $tc = Get-StyleColor "MenuTitleText"
    $oc  = Get-StyleColor "MenuOptionText"; $cv = Get-StyleColor "ConfigValue"
    Write-Color ("╔" + "═" * $iw + "╗") $bc
    [Console]::Write("║"); Write-Color (CenterText "Test Database Connection" (fmtBI "Test Database Connection") $iw) $tc -NL; Write-Color "║" $bc
    Write-Color ("╠══" + "─" * ($iw - 4) + "══╣") $bc
    Write-Color ("║" + " " * $iw + "║") $bc
    $info = "  Host: ${h}:${p}  |  Database: $db  |  User: $usr"
    if ($info.Length -gt $iw) { $info = $info.Substring(0, $iw) }
    [Console]::Write("║"); Write-Color $info $oc -NL; Write-Color (" " * ([Math]::Max(0, $iw - $info.Length)) + "║") $bc
    Write-Color ("║" + " " * $iw + "║") $bc
    $testing = "  Testing TCP connection to ${h}:${p} ..."
    if ($testing.Length -gt $iw) { $testing = $testing.Substring(0, $iw) }
    [Console]::Write("║"); Write-Color $testing $cv -NL; Write-Color (" " * ([Math]::Max(0, $iw - $testing.Length)) + "║") $bc
    Write-Color ("║" + " " * $iw + "║") $bc
    $ok = $false; $errMsg = ""
    try {
        $tcp = [System.Net.Sockets.TcpClient]::new()
        $ar  = $tcp.BeginConnect($h, $p, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne(3000, $false)
        if (-not ($ok -and $tcp.Connected)) { $ok = $false; $errMsg = "Connection timed out or was refused on port $p." }
        try { $tcp.Close() } catch {}
    } catch { $ok = $false; $errMsg = $_.Exception.Message }
    $resFg  = if ($ok) { "Green" } else { "Red" }
    $resTxt = if ($ok) { "  SUCCESS  — TCP port ${p} on ${h} is reachable." } else { "  FAILED   — $errMsg" }
    if ($resTxt.Length -gt $iw) { $resTxt = $resTxt.Substring(0, $iw) }
    [Console]::Write("║")
    $of = [Console]::ForegroundColor; try { [Console]::ForegroundColor = [ConsoleColor]$resFg } catch {}
    [Console]::Write($resTxt); [Console]::ForegroundColor = $of
    Write-Color (" " * ([Math]::Max(0, $iw - $resTxt.Length)) + "║") $bc
    $note = "  Note: verifies TCP reachability only — MySQL credentials are not tested."
    if ($note.Length -gt $iw) { $note = $note.Substring(0, $iw) }
    [Console]::Write("║"); Write-Color $note "DarkGray" -NL; Write-Color (" " * ([Math]::Max(0, $iw - $note.Length)) + "║") $bc
    Write-Color ("║" + " " * $iw + "║") $bc
    $any = "  Press any key to return..."
    [Console]::Write("║"); Write-Color $any $oc -NL; Write-Color (" " * ([Math]::Max(0, $iw - $any.Length)) + "║") $bc
    Write-Color ("╚" + "═" * $iw + "╝") $bc
    [Console]::ReadKey($true) | Out-Null
}

function Show-DatabaseConfig {
    if (-not $Script:DbConfig) { $Script:DbConfig = Get-DefaultDbConfig }
    $unsaved = $false
    $dotBase = 28
    while ($true) {
        Set-BuiltinConsoleSize "Configure Database Connection" 8 0
        [Console]::Clear()
        Write-BuiltinHeader "Configure Database Connection"
        $iw  = Get-InnerWidth
        $bc  = Get-StyleColor "BorderColor"
        $oc  = Get-StyleColor "MenuOptionText"
        $kc  = Get-StyleColor "TriggerKeyColor"
        $bk  = Get-StyleColor "BracketHyphenColor"
        $cv  = Get-StyleColor "ConfigValue"
        $rows = @(
            @{key="H"; lbl="Host / Server IP";  prop="Host";     masked=$false}
            @{key="P"; lbl="Port";               prop="Port";     masked=$false}
            @{key="D"; lbl="Database Name";      prop="Database"; masked=$false}
            @{key="U"; lbl="Username";           prop="Username"; masked=$false}
            @{key="W"; lbl="Password";           prop="Password"; masked=$true}
        )
        Write-Color ("║" + " " * $iw + "║") $bc
        foreach ($row in $rows) {
            $rawVal  = [string]$Script:DbConfig.($row.prop)
            $dispVal = if ($row.masked) {
                if ([string]::IsNullOrEmpty($rawVal)) { "(not set)" } else { "*" * [Math]::Min(16, $rawVal.Length) }
            } else {
                if ([string]::IsNullOrEmpty($rawVal)) { "(not set)" } else { $rawVal }
            }
            $dots  = "." * [Math]::Max(1, $dotBase - $row.lbl.Length)
            $plain = "[$($row.key)] - $($row.lbl)$($dots): $dispVal"
            [Console]::Write("║  ")
            Write-Color "[" $bk -NL; Write-Color $row.key $kc -NL; Write-Color "] - $($row.lbl)" $oc -NL
            Write-Color $dots $bk -NL; Write-Color ": " $bk -NL; Write-Color $dispVal $cv -NL
            $rem = $iw - 2 - $plain.Length; if ($rem -gt 0) { [Console]::Write(" " * $rem) }
            Write-Color "║" $bc
        }
        Write-Color ("║" + " " * $iw + "║") $bc
        $tPlain = "[T] - Test Database Connection"
        [Console]::Write("║  ")
        Write-Color "[" $bk -NL; Write-Color "T" $kc -NL; Write-Color "] - Test Database Connection" $oc -NL
        $rem = $iw - 2 - $tPlain.Length; if ($rem -gt 0) { [Console]::Write(" " * $rem) }
        Write-Color "║" $bc
        if ($unsaved) {
            Write-Color ("║" + " " * $iw + "║") $bc
            $us = "  * Unsaved changes — press S to save"
            [Console]::Write("║"); Write-Color $us "Yellow" -NL
            $rem = $iw - $us.Length; if ($rem -gt 0) { [Console]::Write(" " * $rem) }
            Write-Color "║" $bc
        }
        Write-BuiltinFooter (Get-Date -Format "dddd MMMM dd, yyyy  hh:mm:ss tt")
        Write-Footer "Press key to edit field.  T=Test Connection  S=Save  Esc=Back"
        $key = Read-BuiltinKey
        if (($key.Modifiers -band [ConsoleModifiers]::Alt) -ne 0) { if (Invoke-BuiltinAlt $key) { return }; continue }
        if ($key.Key -eq [ConsoleKey]::Escape) { return }
        $ch = $key.KeyChar.ToString().ToUpper()
        if ($ch -eq "S") { Save-DbConfig; $unsaved = $false; $Script:StatusMsg = "Database settings saved."; continue }
        if ($ch -eq "T") { Invoke-TestDatabaseConnection; continue }
        $hit = $rows | Where-Object { $_.key -eq $ch } | Select-Object -First 1
        if ($hit) {
            [Console]::Clear()
            Write-BuiltinHeader "Configure Database Connection"
            $iw2 = Get-InnerWidth; $bc2 = Get-StyleColor "BorderColor"; $oc2 = Get-StyleColor "MenuOptionText"
            Write-Color ("║" + " " * $iw2 + "║") $bc2
            $hint = if ($hit.masked) { "  Enter new $($hit.lbl) (hidden, blank=keep):" } else { "  Enter new $($hit.lbl) (blank=keep):" }
            [Console]::Write("║"); Write-Color $hint $oc2 -NL
            $rem2 = $iw2 - $hint.Length; if ($rem2 -gt 0) { [Console]::Write(" " * $rem2) }
            Write-Color "║" $bc2
            Write-Color ("║" + " " * $iw2 + "║") $bc2
            Write-Color ("╚" + "═" * $iw2 + "╝") $bc2
            [Console]::Write("  ")
            $inp = ""
            if ($hit.masked) {
                while ($true) {
                    $k2 = [Console]::ReadKey($true)
                    if ($k2.Key -eq [ConsoleKey]::Enter) { break }
                    if ($k2.Key -eq [ConsoleKey]::Backspace) {
                        if ($inp.Length -gt 0) { $inp = $inp.Substring(0, $inp.Length - 1); [Console]::Write("`b `b") }
                    } elseif ($k2.KeyChar -ne [char]0) { $inp += $k2.KeyChar; [Console]::Write("*") }
                }
            } else {
                $inp = Read-Host
            }
            if (-not [string]::IsNullOrEmpty($inp)) {
                if ($hit.prop -eq "Port") {
                    if ($inp -match '^\d+$') { $Script:DbConfig.Port = [int]$inp; $unsaved = $true }
                } else {
                    $Script:DbConfig.($hit.prop) = $inp; $unsaved = $true
                }
            }
        }
    }
}
#endregion

#region === TERMINATION & STARTUP SETTINGS ===
function Show-DurationProgressBar([string]$Label,[int]$Ms) {
    [Console]::Clear()
    $iw   = [Math]::Max([Console]::WindowWidth - 4, 40)
    $bc   = Get-StyleColor "BorderColor"
    $tc   = Get-StyleColor "MenuTitleText"
    $cv   = Get-StyleColor "ConfigValue"
    Write-Color ("╔" + "═" * $iw + "╗") $bc
    $centered = CenterText $Label (fmtBI $Label) $iw
    [Console]::Write("║"); Write-Color $centered $tc -NL; Write-Color "║" $bc
    Write-Color ("╠══" + "─" * ($iw - 4) + "══╣") $bc
    Write-Color ("║" + " " * $iw + "║") $bc

    # Bar inner width (leave 4 chars each side for padding + borders)
    $barW  = $iw - 8
    $barFg = Get-StyleColor "TriggerKeyColor"

    # Animate fill over Ms milliseconds in ~50ms steps
    $steps    = [Math]::Max(1, [Math]::Round($Ms / 50))
    $stepMs   = $Ms / $steps
    for ($i = 1; $i -le $steps; $i++) {
        $filled = [Math]::Round(($i / $steps) * $barW)
        $empty  = $barW - $filled
        $pct    = [Math]::Round(($i / $steps) * 100)
        [Console]::SetCursorPosition(0, [Console]::CursorTop - 1 + 0)
        # Rewrite the bar row (same row each iteration via cursor save)
        $barRow = [Console]::CursorTop
        [Console]::Write("║    ")
        try { [Console]::ForegroundColor = [ConsoleColor]$barFg } catch {}
        [Console]::Write("█" * $filled)
        [Console]::ResetColor()
        [Console]::Write("░" * $empty)
        [Console]::Write("    ║")
        # percentage line below
        [Console]::SetCursorPosition(0, $barRow + 1)
        $pctStr = "  $pct%  ($($Ms)ms)"
        try { [Console]::ForegroundColor = [ConsoleColor]$cv } catch {}
        [Console]::Write("║" + $pctStr.PadRight($iw) + "║")
        [Console]::ResetColor()
        [Console]::SetCursorPosition(0, $barRow)
        Start-Sleep -Milliseconds $stepMs
    }
    [Console]::SetCursorPosition(0, [Console]::CursorTop + 2)
    Write-Color ("╚" + "═" * $iw + "╝") $bc
    Write-Host ""
}

function Show-TerminationStartup {
    $unsaved    = $false
    $propsBackup= $Script:Props | ConvertTo-Json -Depth 5 | ConvertFrom-Json
    while ($true) {
        $p   = $Script:Props
        $yesNo = { param($b) if ($b) { "Yes" } else { "No" } }

        # Build display rows — null key = blank separator row
        $rows = @(
            @{key="1"; lbl="Navigator Startup Menu";               val=$p.StartupMenu}
            @{key="2"; lbl="Navigator Startup Delay Duration (ms)";val="$($p.StartupDelayMs)ms"; type="ms"; prop="StartupDelayMs"}
            $null
            @{key="3"; lbl="Show Startup Splash-Screen";           val=(&$yesNo $p.ShowStartupSplash);     type="toggle"; prop="ShowStartupSplash"}
            @{key="4"; lbl="Startup Splash Screen Duration (ms)";  val="$($p.StartupSplashMs)ms";          type="ms";     prop="StartupSplashMs"}
            @{key="5"; lbl="Navigator Script Startup Delay";       val=$p.StartupScriptDelay}
            $null
            @{key="6"; lbl="Show Termination Splash-Screen";       val=(&$yesNo $p.ShowTerminationSplash); type="toggle"; prop="ShowTerminationSplash"}
            @{key="7"; lbl="Shutdown Splash Screen Duration (ms)"; val="$($p.TerminationSplashMs)ms";      type="ms";     prop="TerminationSplashMs"}
            @{key="8"; lbl="Terminate Active Running Process";     val=(&$yesNo $p.TerminateActiveProcess);type="toggle"; prop="TerminateActiveProcess"}
        )

        # Measure longest line for width
        $dotBase = 48
        $maxLine = ($rows | Where-Object {$_ -ne $null} | ForEach-Object {
            ("[$($_.key)] - $($_.lbl)" + ("." * [Math]::Max(1,$dotBase - $_.lbl.Length)) + ": $($_.val)").Length
        } | Measure-Object -Maximum).Maximum
        $reqIW   = [Math]::Max($maxLine + 8, 60)
        Set-BuiltinConsoleSize "Termination & Startup Settings" ($rows.Count) $reqIW

        [Console]::Clear()
        Write-BuiltinHeader "Termination & Startup Settings"
        $iw  = Get-InnerWidth
        $bc  = Get-StyleColor "BorderColor"
        $oc  = Get-StyleColor "MenuOptionText"
        $kc  = Get-StyleColor "TriggerKeyColor"
        $bk  = Get-StyleColor "BracketHyphenColor"
        $cv  = Get-StyleColor "ConfigValue"

        Write-Color ("║" + " " * $iw + "║") $bc
        foreach ($row in $rows) {
            if ($null -eq $row) {
                Write-Color ("║" + " " * $iw + "║") $bc
                continue
            }
            $dots   = "." * [Math]::Max(1, $dotBase - $row.lbl.Length)
            $plain  = "[$($row.key)] - $($row.lbl)$($dots): $($row.val)"
            $rem    = $iw - 4 - $plain.Length
            [Console]::Write("║"); [Console]::Write("    ")
            Write-Color "[" $bk -NL; Write-Color $row.key $kc -NL
            Write-Color "] - $($row.lbl)" $oc -NL
            Write-Color $dots $bk -NL
            Write-Color ": " $bk -NL
            Write-Color $row.val $cv -NL
            if ($rem -gt 0) { [Console]::Write(" " * $rem) }
            Write-Color "║" $bc
        }
        Write-Color ("║" + " " * $iw + "║") $bc
        Write-BuiltinFooter (Get-Date -Format "dddd MMMM dd, yyyy  hh:mm:ss tt")
        $footerMsg = if ($unsaved) { "UNSAVED CHANGES  S=Save & Apply  Esc=Cancel Changes" } else { "Select a key to adjust.  Esc=Back" }
        Write-Footer $footerMsg

        $key = Read-BuiltinKey
        if (($key.Modifiers -band [ConsoleModifiers]::Alt) -ne 0) {
            if ($unsaved) { $Script:Props = $propsBackup }
            if (Invoke-BuiltinAlt $key) { return }
        }
        if ($key.Key -eq [ConsoleKey]::Escape) {
            if ($unsaved) { $Script:Props = $propsBackup }
            return
        }
        $ch = $key.KeyChar.ToString()

        if ($ch -eq 'S' -and $unsaved) {
            Save-PropsConfig
            $unsaved     = $false
            $propsBackup = $Script:Props | ConvertTo-Json -Depth 5 | ConvertFrom-Json
            $Script:StatusMsg = "Properties saved."
            continue
        }

        $hit = $rows | Where-Object { $_ -ne $null -and $_.key -eq $ch } | Select-Object -First 1
        if ($hit) {
            if ($hit.type -eq "toggle") {
                $Script:Props.$($hit.prop) = -not [bool]$Script:Props.$($hit.prop)
                $unsaved = $true
            } elseif ($hit.type -eq "ms") {
                $curMs = [int]$Script:Props.$($hit.prop)
                Show-DurationProgressBar $hit.lbl $curMs
                Write-Host "Enter new value in milliseconds (blank=keep): " -ForegroundColor Cyan -NoNewline
                $inp = Read-Host
                if ($inp -match '^\d+$') { $Script:Props.$($hit.prop) = [int]$inp; $unsaved = $true }
            }
        }
    }
}
#endregion

#region === FILE & FOLDER PATH MANAGEMENT ===
function Show-FilePathManagement {
    $unsaved     = $false
    $pathsBackup = $Script:PathsData | ConvertTo-Json -Depth 5 | ConvertFrom-Json
    while ($true) {
        $pd   = $Script:PathsData
        $rows = @(
            @{key="1"; lbl="Navigator Root Path";      val=$Script:RootPath}
            @{key="2"; lbl="Menu Config File Path";    val=$pd.MenuConfigPath;  prop="MenuConfigPath"}
            @{key="3"; lbl="Style Config File Path";   val=$pd.StyleConfigPath; prop="StyleConfigPath"}
            @{key="4"; lbl="Properties Config Path";   val=$pd.PropsConfigPath; prop="PropsConfigPath"}
            @{key="5"; lbl="Exit Splash Script Path";  val=$pd.ExitSplashPath;  prop="ExitSplashPath"}
            @{key="6"; lbl="Data Folder Path";         val=$pd.DataFolderPath;  prop="DataFolderPath"}
        )
        $dotBase = 28
        $maxLine = ($rows | ForEach-Object {
            ("[$($_.key)] - $($_.lbl)" + ("." * [Math]::Max(1,$dotBase - $_.lbl.Length)) + ": $($_.val)").Length
        } | Measure-Object -Maximum).Maximum
        $reqIW   = [Math]::Max($maxLine + 8, 60)
        Set-BuiltinConsoleSize "File & Folder Path Management" ($rows.Count) $reqIW

        [Console]::Clear()
        Write-BuiltinHeader "File & Folder Path Management"
        $iw  = Get-InnerWidth
        $bc  = Get-StyleColor "BorderColor"
        $oc  = Get-StyleColor "MenuOptionText"
        $kc  = Get-StyleColor "TriggerKeyColor"
        $bk  = Get-StyleColor "BracketHyphenColor"
        $cv  = Get-StyleColor "ConfigValue"

        Write-Color ("║" + " " * $iw + "║") $bc
        foreach ($row in $rows) {
            $dots  = "." * [Math]::Max(1, $dotBase - $row.lbl.Length)
            $plain = "[$($row.key)] - $($row.lbl)$($dots): $($row.val)"
            $rem   = $iw - 4 - $plain.Length
            [Console]::Write("║"); [Console]::Write("    ")
            Write-Color "[" $bk -NL; Write-Color $row.key $kc -NL
            Write-Color "] - $($row.lbl)" $oc -NL
            Write-Color $dots $bk -NL
            Write-Color ": " $bk -NL
            Write-Color $row.val $cv -NL
            if ($rem -gt 0) { [Console]::Write(" " * $rem) }
            Write-Color "║" $bc
        }
        Write-Color ("║" + " " * $iw + "║") $bc
        Write-BuiltinFooter (Get-Date -Format "dddd MMMM dd, yyyy  hh:mm:ss tt")
        $footerMsg = if ($unsaved) { "UNSAVED CHANGES  S=Save & Apply  Esc=Cancel Changes" } else { "Select a key to edit its path.  Esc=Back" }
        Write-Footer $footerMsg

        $key = Read-BuiltinKey
        if (($key.Modifiers -band [ConsoleModifiers]::Alt) -ne 0) {
            if ($unsaved) { $Script:PathsData = $pathsBackup }
            if (Invoke-BuiltinAlt $key) { return }
        }
        if ($key.Key -eq [ConsoleKey]::Escape) {
            if ($unsaved) { $Script:PathsData = $pathsBackup }
            return
        }
        $ch = $key.KeyChar.ToString()

        if ($ch -eq 'S' -and $unsaved) {
            Save-PathsConfig
            $unsaved     = $false
            $pathsBackup = $Script:PathsData | ConvertTo-Json -Depth 5 | ConvertFrom-Json
            $Script:StatusMsg = "Path settings saved."
            continue
        }

        $hit = $rows | Where-Object { $_.ContainsKey("prop") -and $_.key -eq $ch } | Select-Object -First 1
        if ($hit) {
            [Console]::Clear()
            Write-BuiltinHeader "File & Folder Path Management"
            Write-Host ""
            Write-Host "  Editing: $($hit.lbl)" -ForegroundColor Cyan
            Write-Host "  Current: $($hit.val)" -ForegroundColor DarkYellow
            Write-Host ""
            Write-Host "  Enter new path (blank=keep): " -ForegroundColor Cyan -NoNewline
            $inp = Read-Host
            if ($inp.Trim() -ne "") { $Script:PathsData.$($hit.prop) = $inp.Trim(); $unsaved = $true }
        }
    }
}
#endregion

#region === MANAGE NAVIGATION MENUS ===
function Show-ManageMenus {
    $opts=@(
        [pscustomobject]@{key="1";label="Create New Menu"}
        [pscustomobject]@{key="2";label="Modify Existing Menu"}
        [pscustomobject]@{key="3";label="Remove a Menu"}
    )
    while ($true) {
        Set-BuiltinConsoleSize "Manage Navigation Menus" $opts.Count
        [Console]::Clear()
        Write-BuiltinHeader "Manage Navigation Menus"
        Write-BuiltinOptions $opts
        Write-BuiltinFooter (Get-Date -Format "dddd MMMM dd, yyyy  hh:mm:ss tt")
        Write-Footer "Manage menus.  Esc=Back"

        $key=Read-BuiltinKey
        if (($key.Modifiers -band [ConsoleModifiers]::Alt) -ne 0) { if (Invoke-BuiltinAlt $key) { return } }
        if ($key.Key -eq [ConsoleKey]::Escape) { return }
        switch ($key.KeyChar) {
            '1' { New-NavMenu }
            '2' { Update-NavMenu }
            '3' { Remove-Menu }
        }
    }
}

function New-NavMenu {
    [Console]::Clear(); Write-Host "=== Create New Menu ===" -ForegroundColor Cyan
    $id=Read-Host "Menu ID (unique, no spaces)"
    if ([string]::IsNullOrWhiteSpace($id)) { return }
    if (Get-MenuById $id) { Write-Host "Error: ID '$id' already exists!" -ForegroundColor Red; Read-Host "Press Enter"; return }
    $title  = Read-Host "Title"
    $cols   = Read-Host "Columns (1 or 2)"
    $onLoad = Read-Host "On-load cmdlet (optional, e.g. `$Var = FunctionName)"
    $colNum = if ($cols -eq "2") { 2 } else { 1 }
    $newM   = [pscustomobject]@{id=$id;title=$title;columns=$colNum;onLoad=$onLoad;options=@()}
    $list   = [System.Collections.Generic.List[object]]::new()
    foreach ($m in $Script:Menus.menus) { $list.Add($m) }
    $list.Add($newM); $Script:Menus.menus=$list.ToArray(); Save-MenuConfig
    $Script:StatusMsg="Menu '$id' created."; Write-Host "Created!" -ForegroundColor Green; Read-Host "Press Enter"
}

function Update-NavMenu {
    [Console]::Clear(); Write-Host "=== Modify Existing Menu ===" -ForegroundColor Cyan
    foreach ($m in $Script:Menus.menus) { Write-Host "  $($m.id): $($m.title)" -ForegroundColor DarkYellow }
    $id=Read-Host "Menu ID to modify"; if ([string]::IsNullOrWhiteSpace($id)) { return }
    $menu=Get-MenuById $id
    if ($null -eq $menu) { Write-Host "Not found!" -ForegroundColor Red; Read-Host "Press Enter"; return }
    Write-Host "Current title: $($menu.title)" -ForegroundColor DarkGray
    $t=Read-Host "New title (blank=keep)"; if (-not [string]::IsNullOrWhiteSpace($t)) { $menu.title=$t }
    Write-Host "Current columns: $($menu.columns)" -ForegroundColor DarkGray
    $c=Read-Host "New columns 1/2 (blank=keep)"; if ($c -match '^[12]$') { $menu.columns=[int]$c }
    Write-Host "Current onLoad: $($menu.onLoad)" -ForegroundColor DarkGray
    $ol=Read-Host "New onLoad (blank=keep)"; if (-not [string]::IsNullOrWhiteSpace($ol)) { $menu.onLoad=$ol }
    Save-MenuConfig; $Script:StatusMsg="Menu '$id' updated."; Write-Host "Updated!" -ForegroundColor Green; Read-Host "Press Enter"
}

function Remove-Menu {
    $protected=@("main","settings")
    [Console]::Clear(); Write-Host "=== Remove a Menu ===" -ForegroundColor Cyan
    Write-Host "Protected (cannot remove): $($protected -join ', ')" -ForegroundColor DarkGray
    foreach ($m in $Script:Menus.menus | Where-Object { $_.id -notin $protected }) {
        Write-Host "  $($m.id): $($m.title)" -ForegroundColor DarkYellow }
    $id=Read-Host "Menu ID to remove"; if ([string]::IsNullOrWhiteSpace($id)) { return }
    if ($id -in $protected) { Write-Host "Cannot remove protected menu!" -ForegroundColor Red; Read-Host "Press Enter"; return }
    if (-not (Get-MenuById $id)) { Write-Host "Not found!" -ForegroundColor Red; Read-Host "Press Enter"; return }
    $c=Read-Host "Remove '$id'? (yes/no)"; if ($c -ne "yes") { return }
    $Script:Menus.menus=@($Script:Menus.menus | Where-Object { $_.id -ne $id }); Save-MenuConfig
    $Script:StatusMsg="Menu '$id' removed."; Write-Host "Removed!" -ForegroundColor Green; Read-Host "Press Enter"
}
#endregion

#region === MANAGE MENU OPTIONS ===
function Show-ManageOptions {
    $opts=@(
        [pscustomobject]@{key="1";label="Create New Menu Option"}
        [pscustomobject]@{key="2";label="Modify Existing Menu Option"}
        [pscustomobject]@{key="3";label="Remove a Menu Option"}
    )
    while ($true) {
        Set-BuiltinConsoleSize "Manage Menu Options" $opts.Count
        [Console]::Clear()
        Write-BuiltinHeader "Manage Menu Options"
        Write-BuiltinOptions $opts
        Write-BuiltinFooter (Get-Date -Format "dddd MMMM dd, yyyy  hh:mm:ss tt")
        Write-Footer "Manage menu options.  Esc=Back"

        $key=Read-BuiltinKey
        if (($key.Modifiers -band [ConsoleModifiers]::Alt) -ne 0) { if (Invoke-BuiltinAlt $key) { return } }
        if ($key.Key -eq [ConsoleKey]::Escape) { return }
        switch ($key.KeyChar) {
            '1' { New-MenuOption }
            '2' { Update-MenuOption }
            '3' { Remove-Option }
        }
    }
}

function Select-MenuPrompt([string]$Prompt) {
    $all=@($Script:Menus.menus)
    for ($i=0;$i -lt $all.Count;$i++) { Write-Host "  [$($i+1)] $($all[$i].id) - $($all[$i].title)" -ForegroundColor DarkYellow }
    $v=Read-Host $Prompt
    if ($v -match '^\d+$') { $idx=[int]$v-1; if($idx -ge 0 -and $idx -lt $all.Count){return $all[$idx].id} }
    return $v
}

function New-MenuOption {
    [Console]::Clear(); Write-Host "=== Create New Menu Option ===" -ForegroundColor Cyan
    $menuId=Select-MenuPrompt "Select menu (number or ID)"
    $menu=Get-MenuById $menuId
    if ($null -eq $menu) { Write-Host "Menu not found!" -ForegroundColor Red; Read-Host "Press Enter"; return }
    Write-Host "Menu: $($menu.title)" -ForegroundColor Yellow
    $colChoice=if($menu.columns -eq 2){"left or right"}else{"left"}
    $col=(Read-Host "Column ($colChoice)").ToLower()
    if ($col -notin @("left","right")) { $col="left" }
    $existKeys=@(@($menu.options) | Where-Object {$_.column -eq $col} | ForEach-Object {$_.key.ToUpper()})
    $key=(Read-Host "Trigger key (1 char)").ToUpper()
    if ([string]::IsNullOrWhiteSpace($key) -or $key.Length -ne 1) { Write-Host "Invalid key!" -ForegroundColor Red; Read-Host "Press Enter"; return }
    if ($existKeys -contains $key) { Write-Host "Error: Key '$key' already assigned in $col column!" -ForegroundColor Red; Read-Host "Press Enter"; return }
    $label=Read-Host "Label/description"
    if ([string]::IsNullOrWhiteSpace($label)) { return }
    Write-Host "Types: 1=Open Menu  2=Run Cmdlet  3=Run Function  4=Run Script  5=Set Variable" -ForegroundColor Yellow
    $tc=Read-Host "Type (1-5)"
    $optType=switch($tc){"1"{"menu"}"2"{"cmdlet"}"3"{"function"}"4"{"script"}"5"{"variable"}default{"menu"}}
    $target=""
    if ($optType -eq "variable") { $target=Read-Host "Variable name (e.g. `$SampleVar)" }
    elseif ($optType -eq "menu") { Write-Host ""; $target=Select-MenuPrompt "Target menu" }
    else { $target=Read-Host "Cmdlet/Function/Script path" }
    $window="current"
    if ($optType -in @("cmdlet","script")) { $w=Read-Host "Run in (c)urrent or (n)ew window [c/n]"; if($w -eq "n"){$window="new"} }
    $newOpt=[pscustomobject]@{key=$key;label=$label;column=$col;type=$optType;target=$target;window=$window}
    $list=[System.Collections.Generic.List[object]]::new()
    if ($menu.options) { foreach ($o in $menu.options) { $list.Add($o) } }
    $list.Add($newOpt); $menu.options=$list.ToArray(); Save-MenuConfig
    $Script:StatusMsg="Option '$key' added to '$menuId'."; Write-Host "Created!" -ForegroundColor Green; Read-Host "Press Enter"
}

function Update-MenuOption {
    [Console]::Clear(); Write-Host "=== Modify Menu Option ===" -ForegroundColor Cyan
    $menuId=Select-MenuPrompt "Select menu"; $menu=Get-MenuById $menuId
    if ($null -eq $menu) { Write-Host "Not found!" -ForegroundColor Red; Read-Host "Press Enter"; return }
    if (-not $menu.options -or @($menu.options).Count -eq 0) { Write-Host "No options." -ForegroundColor Yellow; Read-Host "Press Enter"; return }
    foreach ($o in $menu.options) { Write-Host "  [$($o.key)] ($($o.column)) - $($o.label)" -ForegroundColor DarkYellow }
    $key=(Read-Host "Trigger key to modify").ToUpper()
    $opt=@($menu.options) | Where-Object {$_.key.ToUpper() -eq $key} | Select-Object -First 1
    if ($null -eq $opt) { Write-Host "Not found!" -ForegroundColor Red; Read-Host "Press Enter"; return }
    Write-Host "Label: $($opt.label)" -ForegroundColor DarkGray
    $nl=Read-Host "New label (blank=keep)"; if(-not [string]::IsNullOrWhiteSpace($nl)){$opt.label=$nl}
    Write-Host "Column: $($opt.column)" -ForegroundColor DarkGray
    $nc=(Read-Host "New column left/right (blank=keep)").ToLower()
    if ($nc -in @("left","right")) {
        $clash=@($menu.options) | Where-Object {$_.column -eq $nc -and $_.key.ToUpper() -eq $key -and $_ -ne $opt} | Select-Object -First 1
        if ($clash) { Write-Host "Error: key '$key' already exists in '$nc' column!" -ForegroundColor Red; Read-Host "Press Enter" }
        else { $opt.column=$nc }
    }
    Write-Host "Target: $($opt.target)" -ForegroundColor DarkGray
    $nt=Read-Host "New target (blank=keep)"; if(-not [string]::IsNullOrWhiteSpace($nt)){$opt.target=$nt}
    Write-Host "Window: $($opt.window)" -ForegroundColor DarkGray
    $nw=(Read-Host "New window current/new (blank=keep)").ToLower()
    if ($nw -in @("current","new")) { $opt.window=$nw }
    Save-MenuConfig; $Script:StatusMsg="Option '$key' in '$menuId' updated."
    Write-Host "Updated!" -ForegroundColor Green; Read-Host "Press Enter"
}

function Remove-Option {
    [Console]::Clear(); Write-Host "=== Remove Menu Option ===" -ForegroundColor Cyan
    $menuId=Select-MenuPrompt "Select menu"; $menu=Get-MenuById $menuId
    if ($null -eq $menu) { Write-Host "Not found!" -ForegroundColor Red; Read-Host "Press Enter"; return }
    if (-not $menu.options -or @($menu.options).Count -eq 0) { Write-Host "No options." -ForegroundColor Yellow; Read-Host "Press Enter"; return }
    foreach ($o in $menu.options) { Write-Host "  [$($o.key)] ($($o.column)) - $($o.label)" -ForegroundColor DarkYellow }
    $key=(Read-Host "Trigger key to remove").ToUpper()
    $opt=@($menu.options) | Where-Object {$_.key.ToUpper() -eq $key} | Select-Object -First 1
    if ($null -eq $opt) { Write-Host "Not found!" -ForegroundColor Red; Read-Host "Press Enter"; return }
    $c=Read-Host "Remove [$key] '$($opt.label)'? (yes/no)"; if($c -ne "yes"){return}
    $menu.options=@($menu.options | Where-Object {$_.key.ToUpper() -ne $key}); Save-MenuConfig
    $Script:StatusMsg="Option '$key' removed from '$menuId'."; Write-Host "Removed!" -ForegroundColor Green; Read-Host "Press Enter"
}
#endregion

#region === MAIN LOOP ===
function Start-Navigator {
    $Host.UI.RawUI.WindowTitle = "CybtekSTK SysOp Navigator v$($Script:Version)"
    [Console]::TreatControlCAsInput = $true
    [Console]::CursorVisible = $false
    Initialize-Config
    if ($Script:Props -and $Script:Props.ShowStartupSplash -and (Test-Path $Script:StartupSplashPath)) {
        $splashMs = [int]$Script:Props.StartupSplashMs
        try { & powershell -WindowStyle Normal -File $Script:StartupSplashPath -DurationMs $splashMs } catch {}
    }
    Show-Menu "main"
    while ($Script:Running) {
        $tick = [DateTime]::Now
        while (-not [Console]::KeyAvailable -and $Script:Running) {
            Start-Sleep -Milliseconds 50
            if (([DateTime]::Now - $tick).TotalMilliseconds -ge 1000) {
                $tick = [DateTime]::Now
                Update-LiveTimestamp
            }
            if ($Script:ArrowNavActive -and ([DateTime]::Now - $Script:ArrowNavTime).TotalSeconds -ge 7) {
                $Script:ArrowNavActive = $false
                Show-Menu $Script:CurrentMenuId
            }
        }
        if ($Script:Running -and [Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            $r   = Invoke-KeyPress $key
            if ($r -eq "exit")   { $Script:Running = $false }
            elseif ($Script:Running) { Show-Menu $Script:CurrentMenuId }
        }
    }
    [Console]::Clear()
    [Console]::CursorVisible = $true
    [Console]::TreatControlCAsInput = $false
    if ($Script:Props -and $Script:Props.ShowTerminationSplash -and (Test-Path $Script:ExitSplashPath)) {
        $elapsed  = [Math]::Floor(([DateTime]::Now - $Script:StartTime).TotalMinutes)
        $splashMs = [int]$Script:Props.TerminationSplashMs
        try { & powershell -WindowStyle Normal -File $Script:ExitSplashPath -ElapsedMinutes $elapsed -DurationMs $splashMs } catch {}
    }
    Write-Host "CybtekSTK SysOp Navigator terminated." -ForegroundColor Cyan
    if ($Script:Props -and $Script:Props.TerminateActiveProcess) {
        Stop-Process -Id $PID -Force
    }
}

Start-Navigator
#endregion
