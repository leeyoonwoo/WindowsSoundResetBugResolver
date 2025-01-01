[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Variables
$TaskName = "SetVolumeOnWake"
$ScriptFilePath = "C:\SetVolumeOnWake.ps1"
$TaskFilePath = "$env:TEMP\SetVolumeOnWakeScheduledTask.xml"

function Cleanup {
    # Unregister existing scheduled task
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Write-Host "Removing existing scheduled task..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }

    # Remove existing temporary file
    if (Test-Path -Path $TaskFilePath) {
        Write-Host "Removing existing temporary XML file..." -ForegroundColor Yellow
        Remove-Item $TaskFilePath
    }

    # Remove existing script file
    if (Test-Path -Path $ScriptFilePath) {
        Write-Host "Removing existing script file..." -ForegroundColor Yellow
        Remove-Item $ScriptFilePath
    }
}

# Ensure script is running with administrator privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires administrator privileges. Restarting as administrator..." -ForegroundColor Yellow
    Start-Process -FilePath "powershell" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Select Action
do {
    $choice = Read-Host "Choose an action: Y to Install, D to Delete, or N to Exit (Y/D/N)"

    switch ($choice.ToUpper()) {
        "Y" {
            Write-Host "You chose to install." -ForegroundColor Green
        }
        "D" {
            Write-Host "You chose to delete." -ForegroundColor Yellow
            Cleanup
            exit
        }
        "N" {
            Write-Host "You chose to exit." -ForegroundColor Red
            exit
        }
        default {
            Write-Host "Invalid choice. Please enter Y, D, or N." -ForegroundColor Red
        }
    }
} while ($choice.ToUpper() -notin @("Y", "D", "N"))

# Check if AudioDeviceCmdlets module is installed
if (-not (Get-Module -ListAvailable -Name "AudioDeviceCmdlets")) {
    Write-Host "The 'AudioDeviceCmdlets' module is not installed. Attempting to install..." -ForegroundColor Cyan
    try {
        Install-Module -Name "AudioDeviceCmdlets" -Force -Scope CurrentUser -ErrorAction Stop
        Write-Host "The 'AudioDeviceCmdlets' module was installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to install 'AudioDeviceCmdlets'. Please install it manually using the following command:" -ForegroundColor Red
        Write-Host "Install-Module -Name AudioDeviceCmdlets -Force" -ForegroundColor Yellow
        Read-Host "Press Enter to exit."
        exit
    }
}

# Load the AudioDeviceCmdlets module
Write-Host "Loading the 'AudioDeviceCmdlets' module..." -ForegroundColor Cyan
Import-Module "AudioDeviceCmdlets" -Force

# Fetch playback audio devices
Write-Host "Fetching the list of playback audio devices..." -ForegroundColor Cyan
$audioDevices = Get-AudioDevice -List | Where-Object { $_.Type -eq "Playback" }

# Validate playback devices
if (-not $audioDevices) {
    Write-Host "No playback devices were found. Exiting..." -ForegroundColor Red
    Read-Host "Press Enter to exit."
    exit
}

# Display available playback devices
Write-Host "Available playback devices:" -ForegroundColor Cyan
foreach ($device in $audioDevices) {
    Write-Host "Index: $($device.Index) | Name: $($device.Name)" -ForegroundColor White
}

# Prompt the user to select a device
do {
    $selectedIndex = Read-Host "Enter the Index of the playback device you want to select"

    # Find the selected device
    $selectedDevice = $audioDevices | Where-Object { $_.Index -eq $selectedIndex }

    if ($selectedDevice) {
        $selectedID = $selectedDevice.ID
        Write-Host "Selected playback device ID: $selectedID" -ForegroundColor Green
        $isValid = $true
    } else {
        Write-Host "Invalid Index. Try again." -ForegroundColor Red
        $isValid = $false
    }
} while (-not $isValid)

# Prompt the user to set volume
do {
    $SetVolume = Read-Host "Enter default volume [0-100]"

    try {
        $SetVolume = [int]$SetVolume

        if ($SetVolume -ge 0 -and $SetVolume -le 100) {
            Write-Host "Default Volume: $SetVolume" -ForegroundColor Green
            $isValid = $true
        } else {
            Write-Host "Invalid input. Please enter a number between 0 and 100." -ForegroundColor Red
            $isValid = $false
        }
    } catch {
        Write-Host "Invalid input. Please enter a valid number." -ForegroundColor Red
        $isValid = $false
    }
} while (-not $isValid)


# Create the PowerShell script to be executed by the Task Scheduler
$ScriptFile = @"
Import-Module AudioDeviceCmdlets

`$AudioDeviceID = "$selectedID"
`$SetVolume = $SetVolume
`$checkInterval = 50

while (`$true) {
    `$Device = Get-AudioDevice -List | Where-Object { `$_.Id -eq `$AudioDeviceID }

    if (`$Device) {
        try {
            Set-AudioDevice `$AudioDeviceID
            Set-AudioDevice -PlaybackVolume `$SetVolume
            Write-Output "Volume set to `$SetVolume for device `$(`$device.Name)(`$AudioDeviceID)."
            break
        } catch {
            Write-Error "Failed to set volume for device `$AudioDeviceID. Error: `$_"
        }
    } else {
        Write-Host "Device not found. Retrying..." -ForegroundColor Red
    }

    Start-Sleep -Milliseconds `$checkInterval
}
"@

# Create XML configuration for the Task Scheduler
$TaskXML = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>$(Get-Date -Format "yyyy-MM-ddTHH:mm:ss")</Date>
    <Author>$env:UserName</Author>
    <Description>Sets volume to a predefined value after waking from sleep.</Description>
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
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell</Command>
      <Arguments>-ExecutionPolicy Bypass -File "$ScriptFilePath"</Arguments>
    </Exec>
  </Actions>
</Task>
"@

Cleanup

# Create the script and XML files
Write-Host "Creating task script and temporary XML file..." -ForegroundColor Cyan
Set-Content -Path $ScriptFilePath -Value $ScriptFile
Set-Content -Path $TaskFilePath -Value $TaskXML

# Register the scheduled task
Write-Host "Registering the scheduled task..." -ForegroundColor Cyan
schtasks /Create /NP /TN $TaskName /XML $TaskFilePath /F

# Register the scheduled task
Write-Host "Execute the script for testing..." -ForegroundColor Cyan
powershell -ExecutionPolicy Bypass -File "$ScriptFilePath"

# Exit prompt
Write-Host "Configuration Complete."
Read-Host "Press Enter to exit."
