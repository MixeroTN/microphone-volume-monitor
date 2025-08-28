# Hidden Background Monitor Starter
# This script starts the microphone monitor completely hidden without any visible terminal

param(
    [switch]$Silent,
    [switch]$ShowStatus
)

$monitorPath = ".\MicrophoneVolumeMonitor.ps1"

# Function to write status (only if not silent)
function Write-Status {
    param([string]$Message, [string]$Color = "White")
    if (-not $Silent) {
        Write-Host $Message -ForegroundColor $Color
    }
}

# If this script is being run in a visible window and Silent not specified, restart hidden
if (-not $Silent -and -not $ShowStatus -and $Host.UI.RawUI.WindowTitle -ne "Administrator: Windows PowerShell") {
    Write-Status "Restarting in completely hidden mode..." "Yellow"
    
    # Restart this script with -Silent in hidden mode
    $scriptPath = $MyInvocation.MyCommand.Path
    Start-Process powershell.exe -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Silent" -WindowStyle Hidden
    
    Write-Status "Monitor startup initiated in background. Check Task Manager for powershell.exe processes." "Green"
    Write-Status "Log file: $env:TEMP\MicrophoneVolumeMonitor.log" "Cyan"
    Write-Status "`nTo check status later, run:" "White"
    Write-Status "powershell -ExecutionPolicy Bypass -File `"$scriptPath`" -ShowStatus" "Gray"
    return
}

Write-Status "Starting monitor in background..." "Green"

# Kill existing processes
Get-Process powershell | Where-Object { $_.CommandLine -like "*MicrophoneVolumeMonitor*" } | Stop-Process -Force -ErrorAction SilentlyContinue

# Start new background process with full detachment
$processArgs = @(
    "-WindowStyle", "Hidden",
    "-ExecutionPolicy", "Bypass", 
    "-File", $monitorPath
)

$processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
$processStartInfo.FileName = "powershell.exe"
$processStartInfo.Arguments = $processArgs -join " "
$processStartInfo.UseShellExecute = $false
$processStartInfo.CreateNoWindow = $true
$processStartInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

$process = [System.Diagnostics.Process]::Start($processStartInfo)

Write-Status "Monitor started with PID: $($process.Id)" "Green"

# If showing status, do verification
if ($ShowStatus -or -not $Silent) {
    Write-Status "Log file: $env:TEMP\MicrophoneVolumeMonitor.log" "Cyan"
    
    # Quick verification
    Start-Sleep 3
    $running = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
    if ($running) {
        Write-Status "SUCCESS: Monitor is running in background!" "Green"
        
        # Show log entries
        $logFile = "$env:TEMP\MicrophoneVolumeMonitor.log"
        if (Test-Path $logFile) {
            Write-Status "`nRecent log entries:" "Yellow"
            Get-Content $logFile | Select-Object -Last 5
        }
    } else {
        Write-Status "ERROR: Monitor failed to start" "Red"
    }
    
    Write-Status "`nTo stop: Stop-Process -Id $($process.Id)" "White"
}