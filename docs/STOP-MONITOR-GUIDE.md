# Stop Monitor Script - Complete Usage Guide

## Overview

The `stop-monitor.ps1` script provides a comprehensive solution for stopping all microphone volume monitor processes and cleaning up related resources. This script is designed to safely terminate monitor processes, remove startup entries, and clean temporary files.

## Quick Start

### Basic Usage

```powershell
# Interactive mode - asks for confirmation
.\stop-monitor.ps1

# Force stop all processes without prompts
.\stop-monitor.ps1 -Force

# Stop processes and remove from startup
.\stop-monitor.ps1 -RemoveFromStartup

# Full cleanup - stop processes, remove startup, clean files
.\stop-monitor.ps1 -RemoveFromStartup -CleanupFiles -Verbose
```

## Parameters

### Available Options

| Parameter | Type | Description |
|-----------|------|-------------|
| `-RemoveFromStartup` | Switch | Automatically removes the monitor from Windows startup without asking |
| `-CleanupFiles` | Switch | Automatically cleans up temporary files and logs without asking |
| `-Force` | Switch | Stops all processes without confirmation prompts |
| `-Verbose` | Switch | Shows detailed information during execution |

### Parameter Examples

```powershell
# Stop with detailed output
.\stop-monitor.ps1 -Verbose

# Stop and automatically remove from startup
.\stop-monitor.ps1 -RemoveFromStartup

# Complete cleanup without prompts
.\stop-monitor.ps1 -Force -RemoveFromStartup -CleanupFiles

# Interactive cleanup with detailed logging
.\stop-monitor.ps1 -Verbose
```

## What the Script Does

### 1. Current Status Check
- Lists all running PowerShell processes
- Shows SoundVolumeView processes
- Checks Windows startup registry entry
- Reports existing log files

### 2. Process Detection Methods

#### Method 1: WMI Command Line Detection
- Uses `Get-WmiObject Win32_Process` to find processes with "MicrophoneVolumeMonitor" in command line
- Most accurate method for identifying monitor processes
- Works on most Windows versions

#### Method 2: High Memory PowerShell Detection
- Identifies PowerShell processes using >50MB memory
- Presents an interactive list for user confirmation
- Useful when the WMI method fails

#### Method 3: Manual PID Entry
- Allows manual entry of specific Process IDs
- Displays current PowerShell processes with memory usage
- Fallback for edge cases

### 3. Process Termination
- Gracefully stops processes using `Stop-Process`
- Force termination available with `-Force` parameter
- Handles multiple processes simultaneously

### 4. SoundVolumeView Cleanup
- Automatically detects and stops SoundVolumeView processes
- Prevents file access conflicts
- Cleans up audio utility processes

### 5. Startup Management
- Removes Windows startup registry entry
- Shows the current startup configuration
- Optional interactive confirmation

### 6. File Cleanup
- Removes log files: `MicrophoneVolumeMonitor*.log`
- Cleans temporary audio files: `audio_vol_*.txt`, `audio_settings*.txt`
- Removes download files: `soundvolumeview.zip`
- Cleans SoundVolumeView temp directories

## Usage Scenarios

### Scenario 1: Quick Process Stop
**Goal**: Stop running monitor processes
```powershell
.\stop-monitor.ps1 -Force
```
**Result**: All monitor processes stopped, startup and files left intact

### Scenario 2: Complete Removal
**Goal**: Remove monitor entirely from a system
```powershell
.\stop-monitor.ps1 -RemoveFromStartup -CleanupFiles -Verbose
```
**Result**: Processes stopped, startup removed, files cleaned, detailed output

### Scenario 3: Troubleshooting
**Goal**: Interactive cleanup with full control
```powershell
.\stop-monitor.ps1 -Verbose
```
**Result**: Shows detailed status, asks for each action, provides full visibility

### Scenario 4: Temporary Stop
**Goal**: Stop processes but keep autostart for later
```powershell
.\stop-monitor.ps1
# Answer "N" to startup removal prompt
```
**Result**: Processes stopped, will restart on the next boot

## Expected Output

### Successful Execution Example
```
[INFO] === Microphone Volume Monitor - Process Cancellation ===

[INFO] === Current System Status ===
[INFO] PowerShell processes currently running: 13
[WARN] SoundVolumeView processes running: 1
[WARN] Startup entry exists

[INFO] Searching for Microphone Volume Monitor processes...
[INFO] Found 4 monitor processes via WMI
[SUCCESS] Stopped monitor process (PID: 36016)
[SUCCESS] Stopped monitor process (PID: 46640)
[SUCCESS] Stopped monitor process (PID: 33644)
[SUCCESS] Stopped monitor process (PID: 43140)

[INFO] Checking for SoundVolumeView processes...
[SUCCESS] Stopped all SoundVolumeView processes

[SUCCESS] === Cancellation Summary ===
[INFO] PowerShell processes stopped: 4
[INFO] SoundVolumeView processes stopped: 1
[INFO] Startup entry removed: Yes
[INFO] Files cleaned up: 3

✓ Monitor processes have been stopped successfully!
✓ Monitor will not start automatically on next boot

[INFO] Cancellation script completed.
```

## Error Handling

### Common Issues and Solutions

#### No Processes Found
```
[WARN] No monitor processes found automatically.
```
**Solution**: Script will show a manual PID entry option or use `-Verbose` to see all processes

#### Access Denied
```
[ERROR] Failed to stop process 1234: Access is denied
```
**Solution**: Run PowerShell as Administrator or use `-Force` parameter

#### Registry Access Issues
```
[ERROR] Failed to remove startup entry: Access to the registry key is denied
```
**Solution**: Run as Administrator for registry write permissions

#### File Access Conflicts
```
[WARN] Failed to delete MicrophoneVolumeMonitor.log: file is in use
```
**Solution**: Files will be cleaned up after processes are stopped

## Safety Features

### Confirmation Prompts
- Interactive confirmation for each high-memory PowerShell process
- Startup removal confirmation (unless `-RemoveFromStartup` specified)
- File cleanup confirmation (unless `-CleanupFiles` specified)

### Non-Destructive by Default
- Won't remove startup entries unless explicitly requested
- Won't delete files unless explicitly requested
- Shows detailed information before taking action

### Fallback Methods
- Multiple process detection methods
- Graceful handling of permission issues
- Manual PID entry as last resort

## Integration with Monitor System

### Relationship to Other Scripts
- Counterpart to `run-background.ps1` (starts monitor)
- Works with `MicrophoneVolumeMonitor.ps1` (the main monitor)
- Cleans up files created by the monitor system

### Recommended Workflow
1. **Stop Monitor**: `.\stop-monitor.ps1`
2. **Make Changes**: Modify configuration or troubleshoot
3. **Restart Monitor**: `.\run-background.ps1`

## Advanced Usage

### Scripted Automation
```powershell
# Stop monitor in automated script
.\stop-monitor.ps1 -Force -RemoveFromStartup -CleanupFiles

# Verify all stopped
$remaining = Get-Process powershell -ErrorAction SilentlyContinue | Where-Object { $_.WorkingSet -gt 50MB }
if ($remaining) { Write-Host "Warning: Some processes may still be running" }
```

### Scheduled Maintenance
```powershell
# Weekly cleanup script
.\stop-monitor.ps1 -Force -CleanupFiles
Start-Sleep 5
.\run-background.ps1
```

## Verification Commands

### Check if Monitor Stopped
```powershell
# Look for monitor processes
Get-Process powershell | Where-Object { $_.WorkingSet -gt 50MB }

# Check startup entry
Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "MicrophoneVolumeMonitor" -ErrorAction SilentlyContinue
```

### Check Cleanup Success
```powershell
# Look for remaining files
Get-ChildItem "$env:TEMP\MicrophoneVolumeMonitor*.log"
Get-ChildItem "$env:TEMP\audio_vol_*.txt"
```

## Troubleshooting

### Script Won't Run
- Check execution policy: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`
- Run as Administrator for full functionality
- Verify the script path is correct

### Processes Keep Coming Back
- Check if the startup entry was removed
- Look for scheduled tasks or services
- Verify no other scripts are starting the monitor

### Cleanup Incomplete
- Run with `-Verbose` to see detailed cleanup information
- Check file permissions in the temp directory
- Re-run with `-Force` parameter

## Summary

The stop-monitor.ps1 script provides a reliable, user-friendly way to completely stop and clean up the microphone volume monitor system. It handles multiple detection methods, provides safety confirmations, and offers flexible cleanup options to suit different use cases.