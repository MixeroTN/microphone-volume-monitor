# Finding Microphone Volume Monitor in Task Manager

## 🎯 Quick Answer

Your microphone volume monitor appears in Task Manager as a **PowerShell process** with the name:
- **Process Name**: `powershell.exe` 
- **Command Line**: Contains `MicrophoneVolumeMonitor.ps1`

## 📋 Step-by-Step Guide

### Method 1: Task Manager Details Tab (Recommended)

1. **Open Task Manager**:
   - Press `Ctrl + Shift + Esc`
   - Or right-click the taskbar → "Task Manager"

2. **Switch to Details Tab**:
   - Click the "Details" tab at the top

3. **Look for PowerShell processes**:
   - Find all entries named `powershell.exe`
   - There may be multiple PowerShell processes

4. **Enable Command Line column**:
   - Right-click on any column header
   - Select "Command line" from the context menu
   - This shows the full command that started each process

5. **Identify your monitor**:
   - Look for the PowerShell process with the command line containing:
   ```
   -WindowStyle Hidden -ExecutionPolicy Bypass -File "...\MicrophoneVolumeMonitor.ps1"
   ```

### Method 2: Task Manager Processes Tab

1. **Open Task Manager** (`Ctrl + Shift + Esc`)

2. **Go to Processes Tab**:
   - Click the "Processes" tab

3. **Expand Windows PowerShell**:
   - Look for "Windows PowerShell" in the Apps section
   - Click the arrow to expand it
   - You'll see individual PowerShell instances

4. **Identify by resource usage**:
   - The microphone monitor typically uses minimal CPU (~0%)
   - Memory usage around 10-20 MB

### Method 3: Using Task Manager Search

1. **Open Task Manager**

2. **Use the search function** (if available in your Windows version):
   - Look for a search box in Task Manager
   - Type "MicrophoneVolumeMonitor" or "powershell"

## 🔍 What to Look For

### Process Characteristics:
- **Name**: `powershell.exe`
- **Description**: Windows PowerShell
- **CPU Usage**: Very low (0–1%)
- **Memory**: 10–30 MB typically
- **Status**: Running
- **Window**: Hidden (no visible window)

### Command Line Indicators:
The command line will contain these key elements:
```
powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Tools\MicrophoneMonitor\MicrophoneVolumeMonitor.ps1"
```

## 📊 Visual Identification Guide

### In Details Tab:
```
Name             PID    CPU    Memory    Command Line
powershell.exe   53536  0%     15.2 MB   powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Tools\MicrophoneMonitor\MicrophoneVolumeMonitor.ps1"
powershell.exe   12345  0%     8.1 MB    powershell.exe -NoProfile -Command "Get-Date"
powershell.exe   67890  2%     45.3 MB   powershell.exe -File "SomeOtherScript.ps1"
```
**The first one is your microphone monitor!**

### In Processes Tab:
```
📁 Windows PowerShell
   └── Windows PowerShell (15.2 MB)  ← Your monitor
   └── Windows PowerShell (8.1 MB)   ← Different script
```

## 🛠️ Alternative Detection Methods

### Using PowerShell Command Line:
```powershell
# Find the specific monitor process
Get-Process powershell | Where-Object { $_.CommandLine -like "*MicrophoneVolumeMonitor*" }

# See all PowerShell processes with details
Get-Process powershell | Select-Object Id, ProcessName, @{Name="CommandLine";Expression={$_.CommandLine}}
```

### Using Process Explorer (Advanced):
1. Download Process Explorer from Microsoft Sysinternals
2. Run Process Explorer as Administrator
3. Press `Ctrl + F` to search
4. Search for "MicrophoneVolumeMonitor"
5. Double-click the result to see full details

### Using Resource Monitor:
1. Open Resource Monitor (`resmon.exe`)
2. Go to the CPU tab
3. Look for `powershell.exe` processes
4. Check the Command Line column

## ⚡ Managing the Process from Task Manager

### To View Process Details:
1. Right-click on the PowerShell process
2. Select "Properties" or "Go to details"
3. View memory usage, CPU time, etc.

### To End the Process:
1. Right-click on the PowerShell process
2. Select "End task" or "End process"
3. Confirm the action
4. **Note**: The monitor will restart on the next boot (autostart enabled)

### To Restart After Termination:
1. Open PowerShell as an Administrator
2. Navigate to your project folder:
   ```powershell
   cd "C:\Tools\MicrophoneMonitor"
   ```
3. Run the background starter:
   ```powershell
   .\run-background.ps1
   ```

## 🚨 Troubleshooting

### If You Don't See the Process:

1. **Check if it's actually running**:
   ```powershell
   Get-Process powershell | Where-Object { $_.CommandLine -like "*MicrophoneVolumeMonitor*" }
   ```

2. **Check the log file**:
   ```powershell
   Get-Content "$env:TEMP\MicrophoneVolumeMonitor.log" | Select-Object -Last 10
   ```

3. **Restart the monitor**:
   ```powershell
   cd "C:\Tools\MicrophoneMonitor"
   .\run-background.ps1
   ```

### If Multiple PowerShell Processes Exist:

- Use the Command Line column to distinguish them
- Your monitor will have the specific path to `MicrophoneVolumeMonitor.ps1`
- Check the memory usage - the monitor typically uses 10–30 MB

### If the Task Manager Shows "Access Denied":

- Run Task Manager as Administrator:
  - Right-click the Task Manager icon
  - Select "Run as administrator"

## 📱 Quick Reference Card

| What to Look For | Where to Find It |
|------------------|------------------|
| **Process Name** | `powershell.exe` |
| **Tab** | Details or Processes |
| **Memory Usage** | 10-30 MB typically |
| **CPU Usage** | Very low (0-1%) |
| **Command Line** | Contains `MicrophoneVolumeMonitor.ps1` |
| **Window Style** | Hidden (no visible window) |

## 🔄 Expected Process States

### Normal Operation:
- **Status**: Running
- **CPU**: 0–1% (spikes briefly every second during volume check)
- **Memory**: Stable around 15–25 MB
- **Threads**: 1–3 active threads

### During Volume Adjustment:
- **CPU**: May spike to 2–5% briefly
- **Disk Activity**: Minimal (log writing)
- **Network**: None (except initial SoundVolumeView download)

Your microphone volume monitor is now easily identifiable in Task Manager using any of these methods!