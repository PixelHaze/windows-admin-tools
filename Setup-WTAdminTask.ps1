# Self-elevate if not already running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Dynamically resolve current user and wt.exe path
$currentUser = $env:USERNAME
$wtPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"

# Create the scheduled task (runs wt.exe with highest privileges, no UAC)
$action = New-ScheduledTaskAction -Execute $wtPath
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0
$principal = New-ScheduledTaskPrincipal -UserId $currentUser -RunLevel Highest -LogonType Interactive
Register-ScheduledTask -TaskName 'Windows Terminal Admin' -Action $action -Settings $settings -Principal $principal -Force

Write-Host "Scheduled task created successfully." -ForegroundColor Green

# Create desktop shortcut
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut("$env:USERPROFILE\Desktop\Windows Terminal Admin.lnk")
$shortcut.TargetPath = 'C:\Windows\System32\schtasks.exe'
$shortcut.Arguments = '/run /tn "Windows Terminal Admin"'
$shortcut.IconLocation = $wtPath
$shortcut.Description = 'Open Windows Terminal as Admin (no UAC)'
$shortcut.Save()

Write-Host "Desktop shortcut created." -ForegroundColor Green
Write-Host ""
Write-Host "Done! You can delete this script now." -ForegroundColor Cyan
Write-Host "Double-click 'Windows Terminal Admin' on your desktop to launch." -ForegroundColor Cyan
pause
