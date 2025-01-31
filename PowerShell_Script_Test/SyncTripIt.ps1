# Start Outlook as a process to bypass COM issues
Start-Process -FilePath "OUTLOOK.EXE"
Start-Sleep -Seconds 5  # Give Outlook time to start

# Connect to Outlook using COM
$outlook = [Activator]::CreateInstance([Type]::GetTypeFromProgID("Outlook.Application"))
$namespace = $outlook.GetNamespace("MAPI")
$namespace.Logon()

# Get the default calendar
$defaultCalendar = $namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar

# Check if the "TripIt Calendar" exists, and create it if it doesn't
$calendarRoot = $defaultCalendar.Parent
$tripItCalendar = $calendarRoot.Folders | Where-Object { $_.Name -eq "TripIt Calendar" }
if (-not $tripItCalendar) {
    $tripItCalendar = $calendarRoot.Folders.Add("TripIt Calendar", 9)
    Write-Host "TripIt Calendar created!"
}

# Search the default calendar for items containing "TripIt" in the subject or body
$items = $defaultCalendar.Items | Where-Object {
    ($_.Subject -like "*TripIt*") -or ($_.Body -like "*TripIt*")
}

# Move matching items to the TripIt Calendar
foreach ($item in $items) {
    $item.Move($tripItCalendar)
    Write-Host "Moved item: $($item.Subject)"
}

Write-Host "TripIt Calendar synced successfully!"
