# Author: Mariusz Szuster
# email: gmardoss@gmail.com
# Created date: 2025-05-28
# Description: A PowerShell script that toggles the Scroll Lock key at a specified interval.
# Usage: Run the script with an optional interval parameter (in seconds).

[CmdletBinding()]
param(
    [int]$interval = 5 # Default interval is 5 seconds
)

# Validate interval
if ($interval -le 0) {
    Write-Error "Interval must be greater than 0 seconds."
    exit
}

# Add required .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Global state variables
$global:run = $true
$global:interval = $interval

# Constants
$ICON_TEXT_BASE = "Auto Scroll Lock"
$ICON_DLL_PATH = [System.IO.Path]::Combine($env:windir, "System32", "imageres.dll")
$ICON_INDEX = 81

# Add ScrollLockToggler class
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class ScrollLockToggler {
    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, IntPtr dwExtraInfo);

    public static void ToggleScrollLock() {
        const byte VK_SCROLL = 0x91;
        const uint KEYEVENTF_KEYUP = 0x0002;

        // Press Scroll Lock
        keybd_event(VK_SCROLL, 0, 0, IntPtr.Zero);
        // Release Scroll Lock
        keybd_event(VK_SCROLL, 0, KEYEVENTF_KEYUP, IntPtr.Zero);
    }
}
"@

# Add IconExtractor class
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;

public class IconExtractor {
    [DllImport("Shell32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);
}
"@

# Function to create tray icon
function New-TrayIcon {
    $hIcon = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $ICON_DLL_PATH, $ICON_INDEX)
    $iconObj = [System.Drawing.Icon]::FromHandle($hIcon)

    $icon = New-Object System.Windows.Forms.NotifyIcon
    $icon.Icon = $iconObj
    $icon.Visible = $true
    $icon.Text = $ICON_TEXT_BASE
    return $icon
}

# Function to create context menu
function New-ContextMenu {
    $menu = New-Object System.Windows.Forms.ContextMenuStrip

    # Pause item
    $pauseItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $pauseItem.Text = "Pause"
    $pauseItem.Add_Click({ $global:run = $false })

    # Resume item
    $resumeItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $resumeItem.Text = "Resume"
    $resumeItem.Add_Click({ $global:run = $true })

    # Set Time Interval item
    $setIntervalItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $setIntervalItem.Text = "Set Time Interval"
    $setIntervalItem.Add_Click({
        $userInput = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter time interval in seconds:",
            "Set Time Interval",
            "$global:interval"
        )
        if ($userInput -match '^\d+$' -and [int]$userInput -gt 0) {
            $global:interval = [int]$userInput
            $timer.Interval = $global:interval * 1000
        } else {
            [System.Windows.Forms.MessageBox]::Show("Invalid input. Please enter a number greater than 0.")
        }
    })

    # Exit item
    $exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitItem.Text = "Exit"
    $exitItem.Add_Click({
        $timer.Stop()
        $icon.Visible = $false
        [System.Windows.Forms.Application]::Exit()
    })

    # Add items to menu
    $menu.Items.Add($pauseItem)        | Out-Null
    $menu.Items.Add($resumeItem)       | Out-Null
    $menu.Items.Add($setIntervalItem)  | Out-Null
    $menu.Items.Add($exitItem)         | Out-Null

    return $menu
}

# Function to update tray icon text
function Update-TrayIconText {
    param (
        [System.Windows.Forms.NotifyIcon]$icon,
        [string]$status
    )
    $icon.Text = "$ICON_TEXT_BASE $status"
}

# Create tray icon and context menu
$icon = New-TrayIcon
$menu = New-ContextMenu
$icon.ContextMenuStrip = $menu

# Timer to control Scroll Lock keypress
$script:statusScroll = 0
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = $global:interval * 1000
$timer.Add_Tick({
    if ($global:run -eq $true) {
        [ScrollLockToggler]::ToggleScrollLock()

        # Read actual Scroll Lock state
        $state = [System.Windows.Forms.Control]::IsKeyLocked("Scroll")
        Write-Debug "Scroll: $state"

        if ($script:statusScroll -eq 0) {
            $script:statusScroll = 1
            Update-TrayIconText -icon $icon -status "ON"
            Write-Debug "Scroll Lock ON"
        } else {
            $script:statusScroll = 0
            Update-TrayIconText -icon $icon -status "OFF"
            Write-Debug "Scroll Lock OFF"
        }
    }
})

# Start timer and message loop
$timer.Start()
[System.Windows.Forms.Application]::Run()
