# Microphone Volume Monitor - Process Cancellation Script
# Stops all monitor processes and cleans up resources

param(
    [switch]$RemoveFromStartup,
    [switch]$CleanupFiles,
    [switch]$Force,
    [switch]$Verbose
)

function Write-Status {
    param([string]$Message, [string]$Level = "INFO")
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        "INFO"  { "Green" }
        "SUCCESS" { "Cyan" }
        default { "White" }
    }
    Write-Host "[$Level] $Message" -ForegroundColor $color
}

function Stop-MicrophoneMonitorProcesses {
    Write-Status "Searching for Microphone Volume Monitor processes..." "INFO"
    
    $stoppedProcesses = 0
    $totalProcesses = 0
    
    # Method 1: Try to identify by process arguments (works on some PowerShell versions)
    try {
        $monitorProcesses = Get-WmiObject Win32_Process | Where-Object { 
            $_.Name -eq "powershell.exe" -and $_.CommandLine -like "*MicrophoneVolumeMonitor*" 
        }
        
        if ($monitorProcesses) {
            Write-Status "Found $($monitorProcesses.Count) monitor processes via WMI" "INFO"
            foreach ($proc in $monitorProcesses) {
                try {
                    if ($Force) {
                        Stop-Process -Id $proc.ProcessId -Force -ErrorAction Stop
                        Write-Status "Force-stopped monitor process (PID: $($proc.ProcessId))" "SUCCESS"
                    } else {
                        Stop-Process -Id $proc.ProcessId -ErrorAction Stop
                        Write-Status "Stopped monitor process (PID: $($proc.ProcessId))" "SUCCESS"
                    }
                    $stoppedProcesses++
                } catch {
                    Write-Status "Failed to stop process $($proc.ProcessId): $($_.Exception.Message)" "ERROR"
                }
                $totalProcesses++
            }
        }
    } catch {
        if ($Verbose) {
            Write-Status "WMI method failed: $($_.Exception.Message)" "WARN"
        }
    }
    
    # Method 2: Check for processes by examining log files and stopping high-memory PowerShell processes
    if ($stoppedProcesses -eq 0) {
        Write-Status "Using alternative detection method..." "INFO"
        
        # Get all PowerShell processes with high memory usage (likely monitors)
        $powerShellProcesses = Get-Process powershell -ErrorAction SilentlyContinue | Where-Object {
            $_.WorkingSet -gt 50MB  # Monitor typically uses 50MB+
        }
        
        if ($powerShellProcesses) {
            Write-Status "Found $($powerShellProcesses.Count) potential monitor processes (high memory PowerShell)" "INFO"
            
            foreach ($proc in $powerShellProcesses) {
                if ($Force) {
                    # Don't ask, just stop if Force is specified
                    try {
                        Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                        Write-Status "Force-stopped PowerShell process (PID: $($proc.Id), Memory: $([math]::Round($proc.WorkingSet/1MB, 1))MB)" "SUCCESS"
                        $stoppedProcesses++
                    } catch {
                        Write-Status "Failed to stop process $($proc.Id): $($_.Exception.Message)" "ERROR"
                    }
                } else {
                    # Ask for confirmation for each process
                    $response = Read-Host "Stop PowerShell process PID $($proc.Id) (Memory: $([math]::Round($proc.WorkingSet/1MB, 1))MB)? (y/N/a=all)"
                    if ($response -eq 'y' -or $response -eq 'Y' -or $response -eq 'a' -or $response -eq 'A') {
                        try {
                            Stop-Process -Id $proc.Id -ErrorAction Stop
                            Write-Status "Stopped PowerShell process (PID: $($proc.Id))" "SUCCESS"
                            $stoppedProcesses++
                        } catch {
                            Write-Status "Failed to stop process $($proc.Id): $($_.Exception.Message)" "ERROR"
                        }
                        
                        if ($response -eq 'a' -or $response -eq 'A') {
                            # Stop all remaining without asking
                            $Force = $true
                        }
                    }
                }
                $totalProcesses++
            }
        }
    }
    
    # Method 3: Manual process identification by PID (if user knows specific PIDs)
    if ($stoppedProcesses -eq 0) {
        Write-Status "No monitor processes found automatically." "WARN"
        Write-Status "Current PowerShell processes:" "INFO"
        Get-Process powershell -ErrorAction SilentlyContinue | Select-Object Id, @{Name="Memory(MB)";Expression={[math]::Round($_.WorkingSet/1MB,1)}}, CPU, StartTime | Format-Table -AutoSize
        
        if (-not $Force) {
            $manualPids = Read-Host "Enter specific Process IDs to stop (comma-separated, or press Enter to skip)"
            if ($manualPids -and $manualPids.Trim() -ne "") {
                $pidList = $manualPids.Split(',') | ForEach-Object { $_.Trim() }
                foreach ($pid in $pidList) {
                    if ($pid -match '^\d+$') {
                        try {
                            Stop-Process -Id $pid -Force -ErrorAction Stop
                            Write-Status "Stopped process PID $pid" "SUCCESS"
                            $stoppedProcesses++
                        } catch {
                            Write-Status "Failed to stop process $pid`: $($_.Exception.Message)" "ERROR"
                        }
                        $totalProcesses++
                    }
                }
            }
        }
    }
    
    return @{ Stopped = $stoppedProcesses; Total = $totalProcesses }
}

function Stop-SoundVolumeViewProcesses {
    Write-Status "Checking for SoundVolumeView processes..." "INFO"
    
    $svvProcesses = Get-Process SoundVolumeView -ErrorAction SilentlyContinue
    if ($svvProcesses) {
        Write-Status "Found $($svvProcesses.Count) SoundVolumeView processes" "INFO"
        
        try {
            if ($Force) {
                $svvProcesses | Stop-Process -Force -ErrorAction Stop
                Write-Status "Force-stopped all SoundVolumeView processes" "SUCCESS"
            } else {
                $svvProcesses | Stop-Process -ErrorAction Stop
                Write-Status "Stopped all SoundVolumeView processes" "SUCCESS"
            }
            return $svvProcesses.Count
        } catch {
            Write-Status "Error stopping SoundVolumeView processes: $($_.Exception.Message)" "ERROR"
            return 0
        }
    } else {
        Write-Status "No SoundVolumeView processes found" "INFO"
        return 0
    }
}

function Remove-StartupEntry {
    Write-Status "Removing from Windows startup..." "INFO"
    
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
        $entryName = "MicrophoneVolumeMonitor"
        
        $currentEntry = Get-ItemProperty -Path $regPath -Name $entryName -ErrorAction SilentlyContinue
        if ($currentEntry) {
            Remove-ItemProperty -Path $regPath -Name $entryName -ErrorAction Stop
            Write-Status "Successfully removed startup entry" "SUCCESS"
            if ($Verbose) {
                Write-Status "Removed: $($currentEntry.$entryName)" "INFO"
            }
            return $true
        } else {
            Write-Status "No startup entry found to remove" "INFO"
            return $false
        }
    } catch {
        Write-Status "Failed to remove startup entry: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Cleanup-TempFiles {
    Write-Status "Cleaning up temporary files..." "INFO"
    
    $cleanedFiles = 0
    
    # Cleanup patterns
    $cleanupPatterns = @(
        "$env:TEMP\MicrophoneVolumeMonitor*.log",
        "$env:TEMP\audio_vol_*.txt",
        "$env:TEMP\audio_settings*.txt",
        "$env:TEMP\soundvolumeview.zip"
    )
    
    foreach ($pattern in $cleanupPatterns) {
        try {
            $files = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue
            if ($files) {
                foreach ($file in $files) {
                    try {
                        Remove-Item $file.FullName -Force -ErrorAction Stop
                        Write-Status "Deleted: $($file.Name)" "SUCCESS"
                        $cleanedFiles++
                    } catch {
                        Write-Status "Failed to delete $($file.Name): $($_.Exception.Message)" "WARN"
                    }
                }
            }
        } catch {
            if ($Verbose) {
                Write-Status "Pattern $pattern cleanup failed: $($_.Exception.Message)" "WARN"
            }
        }
    }
    
    # Clean up any SoundVolumeView downloads directory
    $svvTempDir = "$env:TEMP\SoundVolumeView"
    if (Test-Path $svvTempDir) {
        try {
            Remove-Item $svvTempDir -Recurse -Force -ErrorAction Stop
            Write-Status "Deleted SoundVolumeView temp directory" "SUCCESS"
            $cleanedFiles++
        } catch {
            Write-Status "Failed to delete SoundVolumeView temp directory: $($_.Exception.Message)" "WARN"
        }
    }
    
    if ($cleanedFiles -eq 0) {
        Write-Status "No temporary files found to clean up" "INFO"
    } else {
        Write-Status "Cleaned up $cleanedFiles temporary files" "SUCCESS"
    }
    
    return $cleanedFiles
}

function Show-CurrentStatus {
    Write-Status "=== Current System Status ===" "INFO"
    
    # Check PowerShell processes
    $psProcesses = Get-Process powershell -ErrorAction SilentlyContinue
    if ($psProcesses) {
        Write-Status "PowerShell processes currently running: $($psProcesses.Count)" "INFO"
        if ($Verbose) {
            $psProcesses | Select-Object Id, @{Name="Memory(MB)";Expression={[math]::Round($_.WorkingSet/1MB,1)}}, CPU | Format-Table -AutoSize
        }
    } else {
        Write-Status "No PowerShell processes running" "INFO"
    }
    
    # Check SoundVolumeView processes
    $svvProcesses = Get-Process SoundVolumeView -ErrorAction SilentlyContinue
    if ($svvProcesses) {
        Write-Status "SoundVolumeView processes running: $($svvProcesses.Count)" "WARN"
    } else {
        Write-Status "No SoundVolumeView processes running" "INFO"
    }
    
    # Check startup entry
    $startupEntry = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "MicrophoneVolumeMonitor" -ErrorAction SilentlyContinue
    if ($startupEntry) {
        Write-Status "Startup entry exists" "WARN"
        if ($Verbose) {
            Write-Status "Entry: $($startupEntry.MicrophoneVolumeMonitor)" "INFO"
        }
    } else {
        Write-Status "No startup entry found" "INFO"
    }
    
    # Check for log files
    $logFiles = Get-ChildItem "$env:TEMP\MicrophoneVolumeMonitor*.log" -ErrorAction SilentlyContinue
    if ($logFiles) {
        Write-Status "Log files found: $($logFiles.Count)" "INFO"
        if ($Verbose) {
            $logFiles | Select-Object Name, @{Name="Size(KB)";Expression={[math]::Round($_.Length/1KB,1)}}, LastWriteTime | Format-Table -AutoSize
        }
    } else {
        Write-Status "No log files found" "INFO"
    }
}

# Main execution
Write-Host ""
Write-Status "=== Microphone Volume Monitor - Process Cancellation ===" "INFO"
Write-Host ""

# Show current status first
Show-CurrentStatus
Write-Host ""

# Stop processes
$processResult = Stop-MicrophoneMonitorProcesses
$svvStopped = Stop-SoundVolumeViewProcesses

Write-Host ""

# Handle startup removal
$startupRemoved = $false
if ($RemoveFromStartup) {
    $startupRemoved = Remove-StartupEntry
} elseif ($processResult.Stopped -gt 0 -and -not $Force) {
    $removeStartup = Read-Host "Remove from Windows startup? (y/N)"
    if ($removeStartup -eq 'y' -or $removeStartup -eq 'Y') {
        $startupRemoved = Remove-StartupEntry
    }
}

Write-Host ""

# Handle file cleanup
$filesCleanedUp = 0
if ($CleanupFiles) {
    $filesCleanedUp = Cleanup-TempFiles
} elseif (($processResult.Stopped -gt 0 -or $svvStopped -gt 0) -and -not $Force) {
    $cleanFiles = Read-Host "Clean up temporary files and logs? (y/N)"
    if ($cleanFiles -eq 'y' -or $cleanFiles -eq 'Y') {
        $filesCleanedUp = Cleanup-TempFiles
    }
}

# Final summary
Write-Host ""
Write-Status "=== Cancellation Summary ===" "SUCCESS"
Write-Status "PowerShell processes stopped: $($processResult.Stopped)" "INFO"
Write-Status "SoundVolumeView processes stopped: $svvStopped" "INFO"
Write-Status "Startup entry removed: $(if ($startupRemoved) { 'Yes' } else { 'No' })" "INFO"
Write-Status "Files cleaned up: $filesCleanedUp" "INFO"
Write-Host ""

if ($processResult.Stopped -gt 0 -or $svvStopped -gt 0) {
    Write-Status "✓ Monitor processes have been stopped successfully!" "SUCCESS"
} else {
    Write-Status "⚠ No monitor processes were found or stopped" "WARN"
}

if ($startupRemoved) {
    Write-Status "✓ Monitor will not start automatically on next boot" "SUCCESS"
}

Write-Host ""
Write-Status "Cancellation script completed." "INFO"