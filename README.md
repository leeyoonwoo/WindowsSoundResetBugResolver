# WindowsSoundResetBugResolver

`WindowsSoundResetBugResolver` is a PowerShell script designed to fix the issue where Windows resets the sound volume to 100 after waking from sleep or rebooting. This script automatically restores the volume of your selected playback device to a predefined level.

---

## Features

- Automatically restores the volume level for a selected playback device after waking from sleep.
- Configurable volume levels between 0 and 100.
- Easy-to-use prompts for setup.

---

## Download

Click the link below to download the latest version:

[Download WindowsSoundResetBugResolver.ps1](https://github.com/leeyoonwoo/WindowsSoundResetBugResolver/releases/download/v1.0.0/WindowsSoundResetBugResolver.ps1)

---

## Installation and Setup

1. **Download the Script**  
   Download the PowerShell Script File `WindowsSoundResetBugResolver.ps1`.

3. **Run the Script**  
   Press `Win + R` to open the **Run** dialog and enter the following command:
  
   ```plaintext
   powershell -ExecutionPolicy Bypass -File "%UserProfile%\Downloads\WindowsSoundResetBugResolver.ps1"
   ```
   or
   ```plaintext
   powershell -ExecutionPolicy Bypass -File "PATH_TO_DOWNLOADED_SCRIPT"
   ```

4. **Follow Prompts**  
   The script will
   - Display a list of playback audio devices.
   - Allow you to select a default device.
   - Ask for a default volume level to set after waking from sleep.
     
5. **Verify Scheduled Task**
   The script will create a Task Scheduler entry named `SetVolumeOnWake` to execute the script automatically after waking from sleep.

### Usage

#### Running the Script Manually

To run the script manually, follow these steps:

1. Open the **Run** dialog by pressing `Win + R`.
2. Enter the following command and press Enter:

   powershell -ExecutionPolicy Bypass -File "C:\Path\To\WindowsSoundResetBugResolver.ps1"

   Replace `C:\Path\To` with the actual directory path where the script is saved.

---

#### Testing the Scheduled Task

1. Put your system to sleep.
2. Wake it up and check if the sound volume has been restored to the predefined level.

---

### Troubleshooting

1. **Scheduled Task Issues**  
   - Open the **Task Scheduler** and ensure a task named `WindowsSoundResetBugResolver` exists.
   - Verify that the task is set to run under your user account.

2. **Playback Device Not Detected**  
   Ensure your playback device is properly connected and functioning before running the script.

---
