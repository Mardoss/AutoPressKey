# AutoPressKey

A PowerShell script that runs in the Windows system tray and automatically toggles the Scroll Lock key at a user-defined interval.
The script provides a tray icon with a context menu, allowing you to pause/resume the automation, change the interval, or exit the application.
It uses Windows Forms for the tray icon and menu, and native Windows API calls to simulate key presses.
The script is easily extendable to support other keys or key combinations in the future.

# Main features:

Automatically toggles the Scroll Lock key at a specified interval
Tray icon with context menu for control (Pause, Resume, Set Interval, Exit)
Interval can be changed at runtime
Runs silently in the background
Written entirely in PowerShell, no external dependencies
Usage:
Run the script with an optional interval parameter (in seconds):

<code>.\AutoPressKey.ps1 -interval 10</code>
or just run <code>Start.vbs</code> in silent mode

Default interval is 5 seconds if not specified.
