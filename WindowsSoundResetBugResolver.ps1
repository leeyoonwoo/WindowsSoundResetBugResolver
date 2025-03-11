[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Define global variables for task and file paths
$taskName = "SetVolumeOnWake"
$saveVolumeTaskName = "SaveVolumeOnSleep"
$scriptFilePath = "C:\SetVolumeOnWake.ps1"
$jsonFilePath = "$env:APPDATA\VolumeSettings.json"
$taskFilePath = "$env:TEMP\SetVolumeOnWakeScheduledTask.xml"
$saveVolumeTaskFilePath = "$env:TEMP\SaveVolumeTask.xml"

# Function to clean up existing tasks and files
function Remove-ExistingTasksAndFiles {
    # Remove existing scheduled tasks if they exist
    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        Write-Host "Removing existing SetVolumeOnWake task..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    }
    if (Get-ScheduledTask -TaskName $saveVolumeTaskName -ErrorAction SilentlyContinue) {
        Write-Host "Removing existing SaveVolumeOnSleep task..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $saveVolumeTaskName -Confirm:$false
    }
    # Remove associated files if they exist
    foreach ($path in @($taskFilePath, $jsonFilePath, $scriptFilePath, $saveVolumeTaskFilePath)) {
        if (Test-Path -Path $path) { Remove-Item -Path $path -Force }
    }
}

# Check if the script is running with administrator privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires administrator privileges. Restarting as administrator..." -ForegroundColor Yellow
    Start-Process -FilePath "powershell" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Prompt user for action with a clear menu
do {
    Write-Host "`nPlease select an option:" -ForegroundColor Cyan
    Write-Host "  Y - Install volume management tasks"
    Write-Host "  D - Delete existing tasks and files"
    Write-Host "  N - Exit without changes"
    $choice = Read-Host "Enter your choice (Y/D/N)"
    switch ($choice.ToUpper()) {
        "Y" {
            Write-Host "Proceeding with installation..." -ForegroundColor Green
            break
        }
        "D" {
            Write-Host "Deleting existing configurations..." -ForegroundColor Yellow
            Remove-ExistingTasksAndFiles
            Write-Host "Configurations deleted successfully." -ForegroundColor Green
            exit
        }
        "N" {
            Write-Host "Exiting without making changes." -ForegroundColor Red
            exit
        }
        default {
            Write-Host "Invalid choice. Please enter Y, D, or N." -ForegroundColor Red
        }
    }
} while ($choice.ToUpper() -notin @("Y", "D", "N"))

# Ensure AudioDeviceCmdlets module is installed and imported
if (-not (Get-Module -ListAvailable -Name "AudioDeviceCmdlets")) {
    Write-Host "Installing AudioDeviceCmdlets module..." -ForegroundColor Cyan
    try {
        Install-Module -Name "AudioDeviceCmdlets" -Force -Scope CurrentUser -ErrorAction Stop
        Write-Host "Module installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to install AudioDeviceCmdlets module: $_" -ForegroundColor Red
        exit
    }
}
Import-Module "AudioDeviceCmdlets" -Force

# Select playback device
Write-Host "`n=== Playback Device Selection ===" -ForegroundColor Cyan
$audioDevices = Get-AudioDevice -List | Where-Object { $_.Type -eq "Playback" }
if (-not $audioDevices) {
    Write-Host "No playback devices found. Exiting..." -ForegroundColor Red
    exit
}
Write-Host "Available playback devices:" -ForegroundColor Cyan
$audioDevices | ForEach-Object { Write-Host "  Index: $($_.Index) | Name: $($_.Name)" -ForegroundColor White }
do {
    $selectedPlaybackIndex = Read-Host "`nEnter the Index of the playback device"
    $selectedPlaybackDevice = $audioDevices | Where-Object { $_.Index -eq $selectedPlaybackIndex }
    if ($selectedPlaybackDevice) {
        $selectedPlaybackID = $selectedPlaybackDevice.ID
        Write-Host "Selected: $($selectedPlaybackDevice.Name)" -ForegroundColor Green
        $isValid = $true
    } else {
        Write-Host "Invalid Index. Please try again." -ForegroundColor Red
        $isValid = $false
    }
} while (-not $isValid)

# Set playback volume
do {
    $setPlaybackVolume = Read-Host "`nEnter default playback volume (0-100)"
    try {
        $setPlaybackVolume = [int]$setPlaybackVolume
        if ($setPlaybackVolume -ge 0 -and $setPlaybackVolume -le 100) {
            Write-Host "Playback volume set to: $setPlaybackVolume" -ForegroundColor Green
            $isValid = $true
        } else {
            throw "Out of range"
        }
    } catch {
        Write-Host "Invalid input. Please enter a number between 0 and 100." -ForegroundColor Red
        $isValid = $false
    }
} while (-not $isValid)

# Optionally select recording device
Write-Host "`n=== Microphone Device Selection ===" -ForegroundColor Cyan
$recordingDevices = Get-AudioDevice -List | Where-Object { $_.Type -eq "Recording" }
$configureRecoding = $false
if (-not $recordingDevices) {
    Write-Host "No microphone devices found. Skipping..." -ForegroundColor Yellow
} else {
    Write-Host "Available microphone devices:" -ForegroundColor Cyan
    $recordingDevices | ForEach-Object { Write-Host "  Index: $($_.Index) | Name: $($_.Name)" -ForegroundColor White }
    do {
        $selectedRecordingIndex = Read-Host "`nEnter the Index of the microphone device (or press Enter to skip)"
        if ([string]::IsNullOrEmpty($selectedRecordingIndex)) {
            Write-Host "Microphone configuration skipped." -ForegroundColor Yellow
            $isValid = $true
            break
        }
        $selectedRecordingDevice = $recordingDevices | Where-Object { $_.Index -eq $selectedRecordingIndex }
        if ($selectedRecordingDevice) {
            $selectedRecordingID = $selectedRecordingDevice.ID
            Write-Host "Selected: $($selectedRecordingDevice.Name)" -ForegroundColor Green
            $configureRecoding = $true
            $isValid = $true
        } else {
            Write-Host "Invalid Index. Please try again." -ForegroundColor Red
            $isValid = $false
        }
    } while (-not $isValid)

    if ($configureRecoding) {
        do {
            $setRecordingVolume = Read-Host "`nEnter default microphone volume (0-100)"
            try {
                $setRecordingVolume = [int]$setRecordingVolume
                if ($setRecordingVolume -ge 0 -and $setRecordingVolume -le 100) {
                    Write-Host "Microphone volume set to: $setRecordingVolume" -ForegroundColor Green
                    $isValid = $true
                } else {
                    throw "Out of range"
                }
            } catch {
                Write-Host "Invalid input. Please enter a number between 0 and 100." -ForegroundColor Red
                $isValid = $false
            }
        } while (-not $isValid)
    }
}

# Ask if volumes should be saved on sleep/shutdown
do {
    $choice = Read-Host "`nSave volume settings automatically on shutdown/sleep? (Y/N)"
    switch ($choice.ToUpper()) {
        "Y" {
            $configureSaveVolume = $true
            Write-Host "Volume saving enabled." -ForegroundColor Green
        }
        "N" {
            $configureSaveVolume = $false
            Write-Host "Volume saving disabled." -ForegroundColor Yellow
        }
        default {
            Write-Host "Invalid choice. Please enter Y or N." -ForegroundColor Red
        }
    }
} while ($choice.ToUpper() -notin @("Y", "N"))

# Generate the script content for scheduled tasks
$configureRecodingBoolText = if ($configureRecoding) { "`$true" } else { "`$false" }
$configureSaveVolumeBoolText = if ($configureSaveVolume) { "`$true" } else { "`$false" }

$ScriptFile = @"
param (
    [string]`$Action
)

# Import required module
Import-Module AudioDeviceCmdlets

# Configuration variables
`$checkInterval = 50
`$maxAttempts = 1000  # Approximately 50 seconds
`$setRecordingVolume = $setRecordingVolume
`$setPlaybackVolume = $setPlaybackVolume
`$configureRecoding = $configureRecodingBoolText
`$configureSaveVolume = $configureSaveVolumeBoolText
`$playbackDeviceID = "$selectedPlaybackID"
`$recordingDeviceID = "$selectedRecordingID"
`$jsonFilePath = "$jsonFilePath"

# Load saved volumes if available
if (`$configureSaveVolume -and (Test-Path "`$jsonFilePath")) {
    try {
        `$settings = Get-Content -Path "`$jsonFilePath" | ConvertFrom-Json
        if (`$settings.PlaybackVolume -ne `$null) { `$setPlaybackVolume = `$settings.PlaybackVolume }
        if (`$settings.MicVolume -ne `$null) { `$setRecordingVolume = `$settings.MicVolume }
    } catch {
        Write-Output "Failed to load JSON settings: `$_ Using predefined volumes."
    }
}

# Handle Save action
if (`$Action -eq "Save") {
    `$PlaybackDevice = Get-AudioDevice -List | Where-Object { `$_.Id -eq `$playbackDeviceID }
    if (`$PlaybackDevice) {
        `$currentPlaybackVolume = [int][math]::Floor([double]((Get-AudioDevice -PlaybackVolume) -replace '[^0-9.]', ''))
        if (`$currentPlaybackVolume) { `$setPlaybackVolume = `$currentPlaybackVolume }
    }
    if (`$configureRecoding) {
        `$RecordingDevice = Get-AudioDevice -List | Where-Object { `$_.Id -eq `$recordingDeviceID -and `$_.Type -eq 'Recording' }
        if (`$RecordingDevice) {
            `$currentMicVolume = [int][math]::Floor([double]((Get-AudioDevice -RecordingVolume) -replace '[^0-9.]', ''))
            if (`$currentMicVolume) { `$setRecordingVolume = `$currentMicVolume }
        }
    }
    `$settings = @{ PlaybackVolume = `$setPlaybackVolume; MicVolume = `$setRecordingVolume }
    `$settings | ConvertTo-Json | Set-Content -Path "`$jsonFilePath"
    Write-Output "Volumes saved to `$jsonFilePath"
}

# Handle Set action
elseif (`$Action -eq "Set") {
    `$isCompletedPlaybackDevice = `$false
    `$isCompletedRecordingDevice = if (`$configureRecoding) { `$false } else { `$true }
    `$attempt = 0
    while (`$attempt -lt `$maxAttempts) {
        if (-not `$isCompletedPlaybackDevice) {
            `$PlaybackDevice = Get-AudioDevice -List | Where-Object { `$_.Id -eq `$playbackDeviceID }
            if (`$PlaybackDevice) {
                try {
                    Set-AudioDevice `$playbackDeviceID
                    Set-AudioDevice -PlaybackVolume `$setPlaybackVolume
                    Write-Output "Set playback volume to `$setPlaybackVolume for `$(`$PlaybackDevice.Name)`n"
                    `$isCompletedPlaybackDevice = `$true
                } catch {
                    Write-Error "Failed to set playback volume: `$_`n"
                }
            }
        }
        if (-not `$isCompletedRecordingDevice) {
            `$RecordingDevice = Get-AudioDevice -List | Where-Object { `$_.Id -eq `$recordingDeviceID -and `$_.Type -eq 'Recording' }
            if (`$RecordingDevice) {
                try {
                    Set-AudioDevice `$recordingDeviceID
                    Set-AudioDevice -RecordingVolume `$setRecordingVolume
                    Write-Output "Set microphone volume to `$setRecordingVolume for `$(`$RecordingDevice.Name)`n"
                    `$isCompletedRecordingDevice = `$true
                } catch {
                    Write-Error "Failed to set microphone volume: `$_`n"
                }
            }
        }
        if (`$isCompletedPlaybackDevice -and `$isCompletedRecordingDevice) { break }
        Start-Sleep -Milliseconds `$checkInterval
        `$attempt++
    }
    if (`$attempt -ge `$maxAttempts) {
        Write-Error "Timeout: Devices not found after 10 seconds."
    }
}
"@

# Define Task Scheduler XML for setting volume on wake
$TaskXML = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>$(Get-Date -Format "yyyy-MM-ddTHH:mm:ss")</Date>
    <Author>$env:UserName</Author>
    <Description>Sets volume to predefined values after waking from sleep or on boot.</Description>
  </RegistrationInfo>
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="System"&gt;&lt;Select Path="System"&gt;*[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and EventID=1]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
    <BootTrigger>
      <Enabled>true</Enabled>
    </BootTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$env:UserName</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>StopExisting</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <ExecutionTimeLimit>PT1M</ExecutionTimeLimit>
    <Enabled>true</Enabled>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell</Command>
      <Arguments>-ExecutionPolicy Bypass -File "$scriptFilePath" -Action Set</Arguments>
    </Exec>
  </Actions>
</Task>
"@

# Define Task Scheduler XML for saving volume on sleep, if enabled
$saveVolumeTaskXML = ""
if ($configureSaveVolume) {
    $saveVolumeTaskXML = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>$(Get-Date -Format "yyyy-MM-ddTHH:mm:ss")</Date>
    <Author>$env:UserName</Author>
    <Description>Saves volume settings before entering sleep mode.</Description>
  </RegistrationInfo>
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="System"&gt;&lt;Select Path="System"&gt;*[System[Provider[@Name='Microsoft-Windows-Kernel-Power'] and EventID=42]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="System"&gt;&lt;Select Path="System"&gt;*[System[Provider[@Name='User32'] and EventID=1074]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$env:UserName</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>StopExisting</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <ExecutionTimeLimit>PT1M</ExecutionTimeLimit>
    <Enabled>true</Enabled>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell</Command>
      <Arguments>-ExecutionPolicy Bypass -File "$scriptFilePath" -Action Save</Arguments>
    </Exec>
  </Actions>
</Task>
"@
}

# Clean up before setting up new tasks
Remove-ExistingTasksAndFiles

# Write script and task files
Write-Host "Creating script and task files..." -ForegroundColor Cyan
Set-Content -Path $scriptFilePath -Value $ScriptFile -Force
Set-Content -Path $taskFilePath -Value $TaskXML -Force
if ($configureSaveVolume) {
    Set-Content -Path $saveVolumeTaskFilePath -Value $saveVolumeTaskXML -Force
}

# Register scheduled tasks
Write-Host "Registering scheduled tasks..." -ForegroundColor Cyan
schtasks /Create /NP /TN $taskName /XML $taskFilePath /F
if ($LASTEXITCODE -eq 0) {
    Write-Host "Task '$taskName' registered successfully." -ForegroundColor Green
} else {
    Write-Host "Failed to register task '$taskName'." -ForegroundColor Red
}
if ($configureSaveVolume) {
    schtasks /Create /NP /TN $saveVolumeTaskName /XML $saveVolumeTaskFilePath /F
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Task '$saveVolumeTaskName' registered successfully." -ForegroundColor Green
    } else {
        Write-Host "Failed to register task '$saveVolumeTaskName'." -ForegroundColor Red
    }
}

# Execute the script for testing
Write-Host "Executing the script for testing..." -ForegroundColor Cyan
if($configureSaveVolume){
    powershell -ExecutionPolicy Bypass -File "$scriptFilePath" -Action Set
}
powershell -ExecutionPolicy Bypass -File "$ScriptFilePath" -Action Save

# Display configuration summary
Write-Host "`n=== Configuration Summary ===" -ForegroundColor Cyan
Write-Host "Playback Device: $($selectedPlaybackDevice.Name)" -ForegroundColor White
Write-Host "Playback Volume: $setPlaybackVolume" -ForegroundColor White
if ($configureRecoding) {
    Write-Host "Recording Device: $($selectedRecordingDevice.Name)" -ForegroundColor White
    Write-Host "Recording Volume: $setRecordingVolume" -ForegroundColor White
}
Write-Host "Save Volume on Sleep: $configureSaveVolume" -ForegroundColor White
Write-Host "Installation completed successfully!" -ForegroundColor Green

# Wait for user confirmation to exit
Read-Host "`nPress Enter to exit"
