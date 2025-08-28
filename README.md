# Microphone Volume Monitor

A PowerShell-based background application that automatically monitors and maintains your microphone volume at 100% to prevent applications like Google Meet from reducing it.

## Problem Solved

This program addresses the issue where certain applications (like Google Meet in the Arc browser) automatically reduce microphone input volume to around 80%, requiring manual adjustment back to 100%.

## Why PowerShell Solution?

This PowerShell implementation eliminates the need for Visual Studio Build Tools installation, making it easier to deploy and use. It leverages Windows built-in capabilities and external utilities for audio control.

## Features

- **No Build Tools Required**: Pure PowerShell solution — no compilation needed
- **Automatic Volume Monitoring**: Continuously monitors the specified microphone device
- **Instant Volume Correction**: Automatically adjusts volume back to 100% when detected below a threshold
- **Windows Startup Integration**: Automatically adds itself to the Windows startup registry
- **External Tool Integration**: Downloads and uses SoundVolumeView for reliable volume control
- **Comprehensive Logging**: Detailed logging to the temp directory for monitoring and debugging
- **Device-Specific Targeting**: Monitors only the specified microphone device
- **Error Resilience**: Handles failures gracefully with retry logic

## Target Device

By default, the program automatically detects and monitors your default Windows microphone. You can also specify a particular microphone device by providing its device ID through the configuration file or command-line parameters.

## Requirements

- **Windows 10 or later**
- **PowerShell 5.1 or later** (included with Windows)
- **Internet connection** (for downloading SoundVolumeView utility)
- **Administrator privileges** (recommended for registry access)

## Installation

### Quick Start

1. **Download the project files**:
   - Download and extract all files to a permanent location (e.g., `C:\Tools\MicrophoneMonitor\`)

2. **First run** (as Administrator):
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File "MicrophoneVolumeMonitor.ps1" -Verbose
   ```

3. **The script will automatically**:
   - Download SoundVolumeView utility
   - Add itself to Windows startup
   - Begin monitoring microphone volume

### Manual Installation Steps

1. **Create directory**:
   ```powershell
   New-Item -ItemType Directory -Path "C:\Tools\MicrophoneMonitor" -Force
   cd "C:\Tools\MicrophoneMonitor"
   ```

2. **Copy the files**:
   - Place all project files in the directory

3. **Set execution policy** (if needed):
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

4. **Run the monitor**:
   ```powershell
   .\MicrophoneVolumeMonitor.ps1 -Verbose
   ```

## Usage

### Command Line Parameters

```powershell
.\MicrophoneVolumeMonitor.ps1 [parameters]
```

**Parameters**:
- `-TargetDeviceId <string>`: Device ID to monitor (default: empty - automatically detects default Windows microphone)
- `-TargetVolume <int>`: Target volume percentage (default: 100)
- `-PollingIntervalSeconds <int>`: Check interval in seconds (default: 1)
- `-Verbose`: Show detailed output and logging

**Examples**:
```powershell
# Standard operation with verbose logging
.\MicrophoneVolumeMonitor.ps1 -Verbose

# Custom target volume (90%)
.\MicrophoneVolumeMonitor.ps1 -TargetVolume 90

# Different polling interval (5 seconds)
.\MicrophoneVolumeMonitor.ps1 -PollingIntervalSeconds 5

# Different device ID
.\MicrophoneVolumeMonitor.ps1 -TargetDeviceId "YOUR_DEVICE_ID"
```

### Background Operation

For a completely hidden background operation, use the provided starter script:
```powershell
.\run-background.ps1
```

This will:
- Start the monitor completely hidden (no visible terminal)
- Allow you to close the terminal window safely
- Continue running independently in the background

See [HIDDEN-BACKGROUND-USAGE.md](docs/HIDDEN-BACKGROUND-USAGE.md) for detailed information about hidden execution.

### Logs

Monitor the log file to see activity:
```powershell
Get-Content "$env:TEMP\MicrophoneVolumeMonitor.log" -Tail 20 -Wait
```

## Configuration

### Configuration File

The microphone volume monitor uses a `config.json` file to store default settings. This file is automatically loaded when the script starts, and its values are used when no command-line parameters are provided.

#### Default Configuration

The default `config.json` file contains:
```json
{
  "TargetDeviceId": "",
  "TargetVolume": 100,
  "PollingIntervalSeconds": 1,
  "Verbose": false
}
```

#### Configuration Options

- **TargetDeviceId**: Device ID to monitor. When empty (""), automatically detects the default Windows microphone
- **TargetVolume**: Target volume percentage (1-100)
- **PollingIntervalSeconds**: How often to check volume in seconds
- **Verbose**: Enable detailed logging output

#### Customizing Settings

1. **Edit the config.json file** to change default behavior:
   ```json
   {
     "TargetDeviceId": "YOUR_SPECIFIC_DEVICE_ID",
     "TargetVolume": 90,
     "PollingIntervalSeconds": 2,
     "Verbose": true
   }
   ```

2. **Command-line parameters override config file values**:
   ```powershell
   # This will use TargetVolume=90 instead of the config file value
   .\MicrophoneVolumeMonitor.ps1 -TargetVolume 90
   ```

3. **Mix config file and parameters**:
   ```powershell
   # Uses config file for most settings, but overrides TargetDeviceId
   .\MicrophoneVolumeMonitor.ps1 -TargetDeviceId "SPECIFIC_DEVICE"
   ```

### Changing Target Device

To monitor a different microphone:

1. **Find your device ID**:
   ```powershell
   Get-WmiObject -Class Win32_SoundDevice | Select-Object Name, DeviceID
   ```

2. **Update the script parameter**:
   ```powershell
   .\MicrophoneVolumeMonitor.ps1 -TargetDeviceId "YOUR_DEVICE_ID"
   ```

### Startup Configuration

The script automatically adds itself to Windows startup. To manage this manually:

**View the current startup entry**:
```powershell
Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "MicrophoneVolumeMonitor"
```

**Remove from startup**:
```powershell
Remove-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "MicrophoneVolumeMonitor"
```

**Add to startup manually**:
```powershell
$scriptPath = "C:\Tools\MicrophoneMonitor\MicrophoneVolumeMonitor.ps1"
$command = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
Set-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "MicrophoneVolumeMonitor" -Value $command
```

## How It Works

### Detection Methods

1. **WMI Device Enumeration**: Uses `Get-WmiObject Win32_SoundDevice` to find audio devices
2. **Registry Search**: Searches Windows audio device registry keys
3. **Device ID Matching**: Matches against the specified device identifier

### Volume Control Methods

1. **SoundVolumeView**: Downloads and uses NirSoft's SoundVolumeView utility (primary method)
2. **Fallback Methods**: Registry manipulation and Windows APIs (limited effectiveness)

### Process Flow

1. **Startup**: Initialize logging, download tools if needed, add to startup
2. **Device Detection**: Search for a target microphone using multiple methods
3. **Volume Monitoring**: Continuously check current volume level
4. **Volume Adjustment**: Set volume to the target level when below a threshold
5. **Error Handling**: Retry on failures, extend intervals on consecutive errors

## External Dependencies

### SoundVolumeView

- **Source**: [NirSoft SoundVolumeView](https://www.nirsoft.net/utils/sound_volume_view.html)
- **Purpose**: Reliable volume control for specific audio devices
- **Installation**: Automatically downloaded on first run
- **Size**: ~50KB
- **License**: Freeware for personal use

The script automatically downloads SoundVolumeView from:
```
https://www.nirsoft.net/utils/soundvolumeview.zip
```

## Management Scripts

### Stop Monitor
Use the included stop script to safely terminate all monitor processes:
```powershell
.\stop-monitor.ps1
```

See [STOP-MONITOR-GUIDE.md](docs/STOP-MONITOR-GUIDE.md) for detailed usage information.

### Finding in Task Manager
To locate the running monitor process in Task Manager, see [TASK_MANAGER_GUIDE.md](docs/TASK_MANAGER_GUIDE.md) for detailed instructions.

## Troubleshooting

### Common Issues

#### Script Won't Start
```powershell
# Check execution policy
Get-ExecutionPolicy
# Set if needed
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### Device Not Found
```powershell
# List all audio devices to find correct ID
Get-WmiObject -Class Win32_SoundDevice | Format-Table Name, DeviceID -AutoSize
```

#### Volume Control Not Working
1. **Check SoundVolumeView**:
   ```powershell
   Test-Path ".\SoundVolumeView.exe"
   ```

2. **Test SoundVolumeView manually**:
   ```powershell
   .\SoundVolumeView.exe /stext audio_devices.txt
   Get-Content audio_devices.txt
   ```

3. **Check internet connection** (for downloading SoundVolumeView):
   ```powershell
   Test-NetConnection www.nirsoft.net -Port 80
   ```

### Log Analysis

**View recent activity**:
```powershell
Get-Content "$env:TEMP\MicrophoneVolumeMonitor.log" | Select-Object -Last 50
```

**Filter for errors**:
```powershell
Get-Content "$env:TEMP\MicrophoneVolumeMonitor.log" | Where-Object { $_ -like "*ERROR*" }
```

**Monitor live**:
```powershell
Get-Content "$env:TEMP\MicrophoneVolumeMonitor.log" -Wait -Tail 10
```

## Performance

- **CPU Usage**: Minimal (~0.1% on modern systems)
- **Memory Usage**: ~10–20 MB
- **Disk Usage**: ~100 KB (script + SoundVolumeView)
- **Network Usage**: One-time download of SoundVolumeView (~50 KB)

## Security Considerations

- **Registry Access**: Writes to current user startup registry only
- **External Downloads**: Downloads SoundVolumeView from the official NirSoft website
- **Local Operation**: No data transmission or external communication after initial setup
- **Permissions**: Runs under current user context (Administrator recommended)

## Uninstalling

1. **Stop the script**:
   ```powershell
   .\stop-monitor.ps1 -RemoveFromStartup -CleanupFiles
   ```

2. **Delete project directory**:
   ```powershell
   Remove-Item "C:\Tools\MicrophoneMonitor" -Recurse -Force
   ```

## Advanced Usage

### Running as Windows Service

For a more robust operation, you can run the script as a Windows service using tools like NSSM:

1. **Download NSSM** (Non-Sucking Service Manager)
2. **Install service**:
   ```cmd
   nssm install MicrophoneVolumeMonitor powershell.exe
   nssm set MicrophoneVolumeMonitor Arguments "-WindowStyle Hidden -ExecutionPolicy Bypass -File C:\Tools\MicrophoneMonitor\MicrophoneVolumeMonitor.ps1"
   nssm start MicrophoneVolumeMonitor
   ```

### Custom Volume Thresholds

Modify the script to use different volume thresholds or multiple target volumes based on the time of day or application usage.

## Project Files

- `MicrophoneVolumeMonitor.ps1` - Main monitoring script
- `run-background.ps1` - Hidden background starter
- `stop-monitor.ps1` - Process management and cleanup script
- `SoundVolumeView.exe` - Volume control utility (auto-downloaded)
- Documentation files with detailed guides

## License

This project is provided as-is for personal use in solving microphone volume issues on Windows systems. SoundVolumeView is used under NirSoft's freeware license terms.