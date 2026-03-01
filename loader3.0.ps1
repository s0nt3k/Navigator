# Get the full path of this script
$currentScript = $MyInvocation.MyCommand.Path

# Get the directory where this script is located
$scriptDirectory = Split-Path -Path $currentScript -Parent

# Get all .ps1 files in the directory excluding this script
$latestScript = Get-ChildItem -Path $scriptDirectory -Filter *.ps1 |
    Where-Object { $_.FullName -ne $currentScript } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

# If no other script is found, exit with message
if (-not $latestScript) {
    Write-Host "No other PowerShell scripts found in directory." -ForegroundColor Yellow
    exit 1
}

# Build command that hides the title bar then runs the target script
$scriptPath = $latestScript.FullName
$command = @"
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
[StructLayout(LayoutKind.Sequential)]
public struct RECT { public int Left, Top, Right, Bottom; }
public class Win32 {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr h, int n);
    [DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr h, int n, int d);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr i, int x, int y, int w, int ht, uint f);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern int GetSystemMetrics(int n);
}
'@
`$h = [Win32]::GetConsoleWindow()
`$s = [Win32]::GetWindowLong(`$h, -16) -band (-bnot 0xC00000)
[Win32]::SetWindowLong(`$h, -16, `$s)
`$screenW = [Win32]::GetSystemMetrics(0)
`$rect = New-Object RECT
[Win32]::GetWindowRect(`$h, [ref]`$rect)
`$winW = `$rect.Right - `$rect.Left
`$x = `$screenW - `$winW
[Win32]::SetWindowPos(`$h, [IntPtr]::Zero, `$x, 0, 0, 0, 37)
& '$scriptPath'
"@

$encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))

# Launch the latest modified script as Administrator with hidden title bar
Start-Process pwsh.exe `
    -ArgumentList "-NoExit", "-ExecutionPolicy Bypass", "-EncodedCommand", $encodedCommand `
    -Verb RunAs