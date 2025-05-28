# Author: Mariusz Szuster
# email: gmardoss@gmail.com
# Created date: 2025-05-28
# Description: A PowerShell script that toggles the Scroll Lock key at a specified interval.
# Usage: Run the script with an optional interval parameter (in seconds).

[CmdletBinding()]
param(
    [int]$interval = 5 # Default interval is 5 seconds
)

# $interval must be greater than 0
if ($interval -le 0) {
    Write-Error "Interval must be greater than 0 seconds."
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Global state variables
$global:run = $true
$global:interval = $interval

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

# Import WinAPI to loading icon from DLL
Add-Type @"

using System;
using System.Runtime.InteropServices;
using System.Drawing;

public class IconExtractor {
    [DllImport("Shell32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);
}
"@

# Load system keyboard icon from imageres.dll (icon index 205)
$dllPath = "C:\Windows\System32\imageres.dll"
$iconIndex = 081
$hIcon = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $dllPath, $iconIndex)
$iconObj = [System.Drawing.Icon]::FromHandle($hIcon)

# Create NotifyIcon (tray icon)
$icon = New-Object System.Windows.Forms.NotifyIcon
$icon.Icon = $iconObj
#$icon.Icon = [System.Drawing.SystemIcons]::Asterisk  # System icon – no external file needed
$icon.Visible = $true
$icon.Text = "Auto Scroll Lock"

# Create context menu
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
    # must be greater than 0
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

# Add menu items to context menu
$menu.Items.Add($pauseItem)        | Out-Null
$menu.Items.Add($resumeItem)       | Out-Null
$menu.Items.Add($setIntervalItem)  | Out-Null
$menu.Items.Add($exitItem)         | Out-Null

# Assign menu to tray icon
$icon.ContextMenuStrip = $menu

# Timer to control Scroll Lock keypress
$script:statusScroll = 0
$baseTextIcon = "Auto Scroll Lock"
$icon.Text = $baseTextIcon
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = $global:interval * 1000  # Timer checks every second
$timer.Add_Tick({
    if ($global:run) {
        # Toggle Scroll Lock and update the LED
        [ScrollLockToggler]::ToggleScrollLock()

        # Read actual Scroll Lock state:
        $state = [System.Windows.Forms.Control]::IsKeyLocked("Scroll")
        Write-Debug "Scroll: $state"

        if ($script:statusScroll -eq 0) {
            $script:statusScroll = 1
            $icon.Text = $baseTextIcon + " ON"
            Write-Debug "Scroll Lock ON"
        }
        else {
            $script:statusScroll = 0
            $icon.Text = $baseTextIcon + " OFF"
            Write-Debug "Scroll Lock OFF"
        }
    }
})

$script:counter = 0
$timer.Start()

# Start message loop
[System.Windows.Forms.Application]::Run()
