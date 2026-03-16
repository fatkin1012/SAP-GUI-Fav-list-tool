param(
    [string]$WindowTitleKeyword = "WAP"
)

Add-Type -AssemblyName System.Windows.Forms

# -----------------------------
# WinAPI for window detection / activation
# -----------------------------
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class Win32 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}
"@

# -----------------------------
# Config
# -----------------------------
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$tcodeFile = Join-Path $scriptDir "tcodes.txt"

# Delay settings (milliseconds)
$delayBetweenKeys   = 1500
$delayAfterOpenBox  = 1000
$delayAfterEnter    = 1000
$delayAfterEditMode = 1000
$delayAfterHome     = 1000
$delayBetweenItems  = 1000

# -----------------------------
# Helpers
# -----------------------------
function Escape-SendKeysText {
    param([string]$Text)

    $escaped = $Text `
        -replace '\{', '{{}' `
        -replace '\}', '{}}' `
        -replace '\+', '{+}' `
        -replace '\^', '{^}' `
        -replace '%', '{%}' `
        -replace '~', '{~}' `
        -replace '\(', '{(}' `
        -replace '\)', '{)}' `
        -replace '\[', '{[}' `
        -replace '\]', '{]}'

    return $escaped
}

function Send-KeyText {
    param(
        [string]$Text,
        [int]$Delay = 100
    )
    [System.Windows.Forms.SendKeys]::SendWait($Text)
    Start-Sleep -Milliseconds $Delay
}

function Get-VisibleWindows {
    $windows = New-Object System.Collections.ArrayList

    $callback = [Win32+EnumWindowsProc]{
        param($hWnd, $lParam)

        if ([Win32]::IsWindowVisible($hWnd)) {
            $sb = New-Object System.Text.StringBuilder 1024
            [void][Win32]::GetWindowText($hWnd, $sb, $sb.Capacity)
            $title = $sb.ToString().Trim()

            if ($title.Length -gt 0) {
                [void]$windows.Add([PSCustomObject]@{
                    Handle = $hWnd
                    Title  = $title
                })
            }
        }
        return $true
    }

    [void][Win32]::EnumWindows($callback, [IntPtr]::Zero)
    return $windows
}

function Activate-WindowByPartialTitle {
    param([string]$Keyword)

    $allWindows = Get-VisibleWindows
    $match = $allWindows | Where-Object { $_.Title -like "*$Keyword*" } | Select-Object -First 1

    if (-not $match) {
        Write-Host ""
        Write-Host "Visible windows found:" -ForegroundColor Yellow
        $allWindows | ForEach-Object { Write-Host $_.Title }
        throw "Could not find any visible window with title containing: $Keyword"
    }

    Write-Host "Matched window:" -ForegroundColor Cyan
    Write-Host $match.Title -ForegroundColor Green

    [void][Win32]::ShowWindowAsync($match.Handle, 5)
    Start-Sleep -Milliseconds 300
    [void][Win32]::SetForegroundWindow($match.Handle)
    Start-Sleep -Milliseconds 700
}

# -----------------------------
# Validate input file
# -----------------------------
if (-not (Test-Path $tcodeFile)) {
    Write-Host "File not found: $tcodeFile" -ForegroundColor Red
    Write-Host "Create tcodes.txt in the same folder as this script." -ForegroundColor Yellow
    exit 1
}

$tcodes = Get-Content $tcodeFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

if ($tcodes.Count -eq 0) {
    Write-Host "No T-codes found in tcodes.txt" -ForegroundColor Yellow
    exit 1
}

Write-Host "Found $($tcodes.Count) T-codes." -ForegroundColor Green
Write-Host "Window title keyword: $WindowTitleKeyword" -ForegroundColor Cyan
Write-Host ""
Write-Host "IMPORTANT BEFORE STARTING:" -ForegroundColor Cyan
Write-Host "1. Open SAP and go to SAP Easy Access"
Write-Host "2. Expand Favorites"
Write-Host "3. Click/select the target folder (Favorites or PW)"
Write-Host "4. Keep the SAP window visible"
Write-Host "5. Do NOT touch keyboard/mouse while script is running"
Write-Host ""

Start-Sleep -Seconds 5

# -----------------------------
# Main loop
# -----------------------------
try {
    Activate-WindowByPartialTitle -Keyword $WindowTitleKeyword

    foreach ($tcode in $tcodes) {
        $renameText = "$tcode - "
        $safeTcode  = Escape-SendKeysText $tcode
        $safeRename = Escape-SendKeysText $renameText

        Write-Host "Processing: $tcode" -ForegroundColor White

        # Ctrl+Shift+F4
        Send-KeyText "^+{F4}" $delayBetweenKeys

        # Type T-code
        Send-KeyText $safeTcode $delayAfterOpenBox

        # Enter
        Send-KeyText "{ENTER}" $delayAfterEnter

        # Ctrl+Shift+F3
        Send-KeyText "^+{F3}" $delayAfterEditMode

        # Home
        Send-KeyText "{HOME}" $delayAfterHome

        # Type rename text
        Send-KeyText $safeRename $delayBetweenKeys

        # Enter
        Send-KeyText "{ENTER}" $delayBetweenItems
    }

    Write-Host ""
    Write-Host "Done! All T-codes processed." -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}