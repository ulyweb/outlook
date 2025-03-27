# Check if running as Administrator
$CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$Principal = New-Object Security.Principal.WindowsPrincipal $CurrentUser
$IsAdmin = $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "This script must be run as an Administrator. Restarting with elevated privileges..."
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Define Outlook RoamCache Path
$RoamCachePath = "$env:LOCALAPPDATA\Microsoft\Outlook\RoamCache"

# Step 1: Ensure Outlook is closed before making any changes
$OutlookProcess = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
if ($OutlookProcess) {
    Write-Host "Closing Outlook..."
    Stop-Process -Name OUTLOOK -Force
    Start-Sleep -Seconds 5
} else {
    Write-Host "Outlook is already closed."
}

# Step 2: Clear Auto-Complete Cache for Legacy Mode
Write-Host "Clearing Auto-Complete Cache..."
if (Test-Path $RoamCachePath) {
    Get-ChildItem -Path $RoamCachePath -Filter "Stream_Autocomplete_*" | Remove-Item -Force
    Write-Host "Auto-Complete Cache Cleared."
} else {
    Write-Host "RoamCache folder not found. Skipping Auto-Complete Cache Reset."
}

# Step 3: Ensure Outlook Global Address List (GAL) is Updated
Write-Host "Checking GAL Update Status..."
$GALUpdatePath = "$env:APPDATA\Microsoft\Outlook\Offline Address Books"

if (Test-Path $GALUpdatePath) {
    Write-Host "Removing existing GAL Cache..."
    Remove-Item -Path "$GALUpdatePath\*" -Recurse -Force
    Write-Host "GAL Cache Removed. Outlook will re-download it on next startup."
} else {
    Write-Host "GAL Cache not found. Assuming Online GAL is used."
}

# Step 4: Restart Outlook
Write-Host "Restarting Outlook..."
Start-Process "outlook.exe"

Write-Host "Outlook Cache Refresh Completed!"
