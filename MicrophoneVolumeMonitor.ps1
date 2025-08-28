# Microphone Volume Monitor - PowerShell Implementation
# Monitors and maintains microphone volume at 100%

param(
    [string]$TargetDeviceId,
    [int]$TargetVolume,
    [int]$PollingIntervalSeconds,
    [switch]$Verbose
)

# Load configuration from config.json
function Load-Configuration {
    $configPath = Join-Path $PSScriptRoot "config.json"
    $defaultConfig = @{
        TargetDeviceId = ""
        TargetVolume = 100
        PollingIntervalSeconds = 1
        Verbose = $false
    }
    
    if (Test-Path $configPath) {
        try {
            $configContent = Get-Content $configPath -Raw | ConvertFrom-Json
            return @{
                TargetDeviceId = $configContent.TargetDeviceId
                TargetVolume = $configContent.TargetVolume
                PollingIntervalSeconds = $configContent.PollingIntervalSeconds
                Verbose = $configContent.Verbose
            }
        } catch {
            Write-Warning "Failed to load configuration from config.json: $($_.Exception.Message). Using defaults."
            return $defaultConfig
        }
    } else {
        Write-Warning "Configuration file config.json not found. Using defaults."
        return $defaultConfig
    }
}

# Load configuration
$config = Load-Configuration

# Use parameters if provided, otherwise use config values
$script:TargetDeviceId = if ($PSBoundParameters.ContainsKey('TargetDeviceId')) { $TargetDeviceId } else { $config.TargetDeviceId }
$script:TargetVolume = if ($PSBoundParameters.ContainsKey('TargetVolume')) { $TargetVolume } else { $config.TargetVolume }
$script:PollingInterval = if ($PSBoundParameters.ContainsKey('PollingIntervalSeconds')) { $PollingIntervalSeconds } else { $config.PollingIntervalSeconds }
if (-not $PSBoundParameters.ContainsKey('Verbose') -and $config.Verbose) { $Verbose = $true }

# Configuration
$script:LogPath = "$env:TEMP\MicrophoneVolumeMonitor.log"
$script:MutexName = "Global\MicrophoneVolumeMonitor_Mutex"

function Test-SingleInstance {
    try {
        # Check if another instance is already running
        $existingProcesses = Get-Process powershell -ErrorAction SilentlyContinue | Where-Object { 
            $_.Id -ne $PID -and $_.CommandLine -like "*MicrophoneVolumeMonitor.ps1*" 
        }
        
        if ($existingProcesses) {
            Write-Log "Another microphone monitor instance is already running (PID: $($existingProcesses.Id -join ', '))" "WARN"
            Write-Log "Stopping duplicate instances to prevent conflicts..." "INFO"
            
            # Stop other instances
            $existingProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            
            Write-Log "Duplicate instances stopped. Continuing with this instance." "INFO"
        }
        
        return $true
    } catch {
        Write-Log "Error checking for existing instances: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"
    
    if ($Verbose -or $Level -eq "ERROR" -or $Level -eq "WARN") {
        Write-Host $logMessage -ForegroundColor $(
            switch ($Level) {
                "ERROR" { "Red" }
                "WARN"  { "Yellow" }
                "INFO"  { "Green" }
                default { "White" }
            }
        )
    }
    
    # Safe file writing with retry logic to prevent access conflicts
    $maxRetries = 3
    $retryCount = 0
    $writeSuccess = $false
    
    while (-not $writeSuccess -and $retryCount -lt $maxRetries) {
        try {
            # Use unique temporary file per process to avoid conflicts
            $processId = [System.Diagnostics.Process]::GetCurrentProcess().Id
            $tempLogFile = "$env:TEMP\MicrophoneVolumeMonitor_${processId}.log"
            
            # Try to write to process-specific log file first
            Add-Content -Path $tempLogFile -Value $logMessage -ErrorAction Stop
            
            # Then try to append to main log file
            Add-Content -Path $script:LogPath -Value $logMessage -ErrorAction Stop
            $writeSuccess = $true
            
        } catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Start-Sleep -Milliseconds (100 * $retryCount)  # Progressive delay
            } else {
                # Fallback: write only to process-specific log
                try {
                    $processId = [System.Diagnostics.Process]::GetCurrentProcess().Id
                    $tempLogFile = "$env:TEMP\MicrophoneVolumeMonitor_${processId}.log"
                    Add-Content -Path $tempLogFile -Value $logMessage -ErrorAction Stop
                } catch {
                    # Last resort: skip file logging but continue execution
                    Write-Warning "Unable to write to log file: $($_.Exception.Message)"
                }
            }
        }
    }
}

function Find-TargetMicrophone {
    Write-Log "Searching for target microphone device..." "INFO"
    
    try {
        # If TargetDeviceId is empty, find the default Windows microphone
        if ([string]::IsNullOrEmpty($script:TargetDeviceId)) {
            Write-Log "No specific device ID configured. Searching for default Windows microphone..." "INFO"
            
            # Method 1: Use SoundVolumeView to find default capture microphone
            $svv = "${env:ProgramFiles}\SoundVolumeView\SoundVolumeView.exe"
            $svvPortable = ".\SoundVolumeView.exe"
            
            $svcPath = $null
            if (Test-Path $svv) { $svcPath = $svv }
            elseif (Test-Path $svvPortable) { $svcPath = $svvPortable }
            
            if ($svcPath) {
                try {
                    # Export current audio devices
                    $tempFile = "$env:TEMP\audio_devices_detection.txt"
                    & $svcPath /stext $tempFile
                    
                    if (Test-Path $tempFile) {
                        Start-Sleep -Milliseconds 100  # Wait for file to be written
                        $content = Get-Content $tempFile -ErrorAction SilentlyContinue
                        
                        if ($content) {
                            # Find device block with Default: Capture
                            $deviceBlock = @()
                            $inDeviceBlock = $false
                            $isDefaultCapture = $false
                            $isCaptureDevice = $false
                            
                            for ($i = 0; $i -lt $content.Count; $i++) {
                                $line = $content[$i]
                                
                                # Check if we're starting a new device block
                                if ($line -eq "==================================================") {
                                    # If we found a default capture device in the previous block, we're done
                                    if ($inDeviceBlock -and $isCaptureDevice -and $isDefaultCapture) {
                                        break
                                    }
                                    # Reset for new device block
                                    $deviceBlock = @()
                                    $inDeviceBlock = $false
                                    $isDefaultCapture = $false
                                    $isCaptureDevice = $false
                                } elseif ($line -like "*Name*" -and $line -like "*:*" -and -not $inDeviceBlock) {
                                    # Starting a new device block
                                    $inDeviceBlock = $true
                                    $deviceBlock += $line
                                } elseif ($inDeviceBlock) {
                                    # We're in a device block, collect lines
                                    $deviceBlock += $line
                                    
                                    # Check if this is a capture device
                                    if ($line -like "*Direction*" -and $line -like "*Capture*") {
                                        $isCaptureDevice = $true
                                    }
                                    
                                    # Check if this is the default capture device
                                    if ($line -like "*Default*" -and $line -like "*Capture*" -and $line -notlike "*Default Communications*") {
                                        $isDefaultCapture = $true
                                    }
                                }
                            }
                            
                            if ($deviceBlock.Count -gt 0 -and $isCaptureDevice -and $isDefaultCapture) {
                                # Extract device information from the block
                                $deviceName = ($deviceBlock | Where-Object { $_ -like "*Device Name*" } | ForEach-Object { $_.Split(':')[1].Trim() }) | Select-Object -First 1
                                $itemId = ($deviceBlock | Where-Object { $_ -like "*Item ID*" } | ForEach-Object { $_.Split(':')[1].Trim() }) | Select-Object -First 1
                                
                                if ($deviceName -and $itemId) {
                                    # Extract the GUID from Item ID (e.g., {ebbbb3b1-5416-4af0-9848-3bb4398b600f})
                                    $guidMatch = $itemId | Select-String "\{([^}]+)\}$"
                                    if ($guidMatch) {
                                        $extractedDeviceId = $guidMatch.Matches[0].Groups[1].Value
                                        Write-Log "Found default capture microphone via SoundVolumeView: $deviceName (ID: $extractedDeviceId)" "INFO"
                                        
                                        # Set the extracted device ID for further processing
                                        $script:TargetDeviceId = $extractedDeviceId
                                        
                                        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
                                        return @{
                                            Found = $true
                                            Method = "SoundVolumeView_Default"
                                            DeviceName = $deviceName
                                            DeviceID = $extractedDeviceId
                                            ItemID = $itemId
                                        }
                                    }
                                }
                            }
                        }
                        
                        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
                    }
                } catch {
                    Write-Log "Error using SoundVolumeView for default microphone detection: $($_.Exception.Message)" "WARN"
                }
            }
            
            # Method 2: Try to find default microphone using WMI (fallback)
            Write-Log "SoundVolumeView method failed, trying WMI..." "INFO"
            $audioDevices = Get-WmiObject -Class Win32_SoundDevice -ErrorAction SilentlyContinue | Where-Object { 
                $_.Name -like "*microphone*" -or $_.Name -like "*mic*" 
            }
            
            if ($audioDevices) {
                $defaultMic = $audioDevices | Select-Object -First 1
                Write-Log "Found microphone via WMI: $($defaultMic.Name)" "INFO"
                return @{
                    Found = $true
                    Method = "WMI_Default"
                    Device = $defaultMic
                    Name = $defaultMic.Name
                    DeviceID = $defaultMic.DeviceID
                }
            }
            
            # Method 3: Try to find any capture device
            Write-Log "No microphone found by name. Searching for any capture device..." "INFO"
            $captureDevices = Get-WmiObject -Class Win32_SoundDevice -ErrorAction SilentlyContinue | Where-Object {
                $_.DeviceID -like "*capture*" -or $_.Description -like "*capture*" -or $_.Name -like "*input*"
            }
            
            if ($captureDevices) {
                $defaultCapture = $captureDevices | Select-Object -First 1
                Write-Log "Found default capture device: $($defaultCapture.Name)" "INFO"
                return @{
                    Found = $true
                    Method = "WMI_Capture"
                    Device = $defaultCapture
                    Name = $defaultCapture.Name
                    DeviceID = $defaultCapture.DeviceID
                }
            }
            
            # Last resort: Use any audio device (will rely on SoundVolumeView to find microphones)
            $allAudioDevices = Get-WmiObject -Class Win32_SoundDevice -ErrorAction SilentlyContinue
            if ($allAudioDevices) {
                $firstDevice = $allAudioDevices | Select-Object -First 1
                Write-Log "Using first available audio device as fallback: $($firstDevice.Name)" "WARN"
                return @{
                    Found = $true
                    Method = "WMI_Fallback"
                    Device = $firstDevice
                    Name = $firstDevice.Name
                    DeviceID = $firstDevice.DeviceID
                }
            }
        } else {
            Write-Log "Searching for specific device ID: $script:TargetDeviceId" "INFO"
        }
        
        # Method 1: Try WMI Win32_SoundDevice (for specific device ID)
        if (-not [string]::IsNullOrEmpty($script:TargetDeviceId)) {
            $soundDevices = Get-WmiObject -Class Win32_SoundDevice -ErrorAction SilentlyContinue
            foreach ($device in $soundDevices) {
                if ($device.DeviceID -like "*$script:TargetDeviceId*") {
                    Write-Log "Found target device via WMI: $($device.Name)" "INFO"
                    return @{
                        Found = $true
                        Method = "WMI"
                        Device = $device
                        Name = $device.Name
                        DeviceID = $device.DeviceID
                    }
                }
            }
        }
        
        # Method 2: Try searching registry
        $registryPaths = @(
            "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e96c-e325-11ce-bfc1-08002be10318}",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture"
        )
        
        foreach ($regPath in $registryPaths) {
            if (Test-Path $regPath) {
                $subKeys = Get-ChildItem $regPath -ErrorAction SilentlyContinue
                foreach ($key in $subKeys) {
                    $properties = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
                    if ($properties) {
                        $values = $properties.PSObject.Properties | Where-Object { $_.Value -like "*$script:TargetDeviceId*" }
                        if ($values) {
                            Write-Log "Found target device in registry: $($key.Name)" "INFO"
                            return @{
                                Found = $true
                                Method = "Registry"
                                RegistryPath = $key.PSPath
                                Key = $key
                            }
                        }
                    }
                }
            }
        }
        
        Write-Log "Target microphone device not found" "WARN"
        return @{ Found = $false }
        
    } catch {
        Write-Log "Error searching for target device: $($_.Exception.Message)" "ERROR"
        return @{ Found = $false }
    }
}

function Get-MicrophoneVolume {
    param($DeviceInfo)
    
    try {
        # Try using SoundVolumeView if available
        $svv = "${env:ProgramFiles}\SoundVolumeView\SoundVolumeView.exe"
        $svvPortable = ".\SoundVolumeView.exe"
        
        $svcPath = $null
        if (Test-Path $svv) { $svcPath = $svv }
        elseif (Test-Path $svvPortable) { $svcPath = $svvPortable }
        
        if ($svcPath) {
            # Enhanced process cleanup - kill all SoundVolumeView processes and wait
            $existingProcesses = Get-Process SoundVolumeView -ErrorAction SilentlyContinue
            if ($existingProcesses) {
                $existingProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 500  # Increased wait time
                
                # Double-check and force kill any remaining processes
                $remainingProcesses = Get-Process SoundVolumeView -ErrorAction SilentlyContinue
                if ($remainingProcesses) {
                    $remainingProcesses | ForEach-Object { 
                        try { $_.Kill() } catch { }
                    }
                    Start-Sleep -Milliseconds 200
                }
            }
            
            # Use highly unique temporary file to avoid any conflicts
            $processId = [System.Diagnostics.Process]::GetCurrentProcess().Id
            $timestamp = Get-Date -Format "yyyyMMddHHmmssfff"  # Include milliseconds
            $randomSuffix = Get-Random -Minimum 1000 -Maximum 9999
            $tempFile = "$env:TEMP\audio_vol_${processId}_${timestamp}_${randomSuffix}.txt"
            
            try {
                # Ensure temp file doesn't exist before we start
                if (Test-Path $tempFile) {
                    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Milliseconds 50
                }
                
                # Export current volume settings with timeout and proper error handling
                $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
                $processStartInfo.FileName = $svcPath
                $processStartInfo.Arguments = "/stext `"$tempFile`""
                $processStartInfo.UseShellExecute = $false
                $processStartInfo.CreateNoWindow = $true
                $processStartInfo.RedirectStandardOutput = $true
                $processStartInfo.RedirectStandardError = $true
                
                $process = [System.Diagnostics.Process]::Start($processStartInfo)
                $completed = $process.WaitForExit(8000)  # 8 second timeout
                
                if (-not $completed) {
                    Write-Log "SoundVolumeView process timeout - terminating" "WARN"
                    if (!$process.HasExited) {
                        $process.Kill()
                        Start-Sleep -Milliseconds 200
                    }
                } elseif (Test-Path $tempFile) {
                    # Wait a moment for file to be fully written
                    Start-Sleep -Milliseconds 100
                    
                    $content = Get-Content $tempFile -ErrorAction SilentlyContinue
                    if ($content) {
                        # Find device block containing our target device ID
                        $deviceBlock = @()
                        $inTargetDevice = $false
                        
                        for ($i = 0; $i -lt $content.Count; $i++) {
                            $line = $content[$i]
                            
                            # Check if we're starting a new device block
                            if ($line -eq "==================================================") {
                                if ($inTargetDevice) {
                                    # We've found our target device block, stop collecting
                                    break
                                }
                                # Reset for potential new device
                                $deviceBlock = @()
                                $inTargetDevice = $false
                            } elseif ($line -like "*Item ID*" -and $line -like "*$script:TargetDeviceId*") {
                                # Found our target device ID in the Item ID field, mark this block as target
                                $inTargetDevice = $true
                                $deviceBlock += $line
                            } elseif ($inTargetDevice) {
                                # Collect lines for the target device
                                $deviceBlock += $line
                            } elseif ($line -like "*Name*" -and -not $inTargetDevice) {
                                # Starting a new device block, start collecting
                                $deviceBlock = @($line)
                            } elseif ($deviceBlock.Count -gt 0 -and -not $inTargetDevice) {
                                # Continue collecting lines until we find Item ID
                                $deviceBlock += $line
                            }
                        }
                        
                        if ($deviceBlock.Count -gt 0) {
                            # Look for volume in the collected device block
                            $volumeLine = $deviceBlock | Where-Object { $_ -like "*Volume Percent*" }
                            if ($volumeLine) {
                                # Parse volume from SoundVolumeView output format: "Volume Percent    : XX.X%"
                                $volumeMatch = $volumeLine | Select-String "Volume Percent\s*:\s*(\d+(?:\.\d+)?)%"
                                if ($volumeMatch) {
                                    $currentVolume = [math]::Round([double]$volumeMatch.Matches[0].Groups[1].Value)
                                    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
                                    return $currentVolume
                                }
                            }
                        }
                        
                        # Fallback: search for any microphone device
                        $microphoneLines = $content | Where-Object { $_ -like "*Microphone*" }
                        if ($microphoneLines.Count -gt 0) {
                            Write-Log "Target device ID not found, checking any microphone device" "WARN"
                            # Try to find volume near microphone mentions
                            for ($i = 0; $i -lt $content.Count; $i++) {
                                if ($content[$i] -like "*Microphone*" -and $content[$i] -like "*Device Name*") {
                                    # Look for volume in the next 20 lines
                                    for ($j = $i; $j -lt [math]::Min($i + 20, $content.Count); $j++) {
                                        if ($content[$j] -like "*Volume Percent*") {
                                            $volumeMatch = $content[$j] | Select-String "Volume Percent\s*:\s*(\d+(?:\.\d+)?)%"
                                            if ($volumeMatch) {
                                                $currentVolume = [math]::Round([double]$volumeMatch.Matches[0].Groups[1].Value)
                                                Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
                                                return $currentVolume
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                
            } catch {
                Write-Log "SoundVolumeView execution error: $($_.Exception.Message)" "WARN"
            } finally {
                # Ensure cleanup happens regardless of success/failure
                if ($process -and !$process.HasExited) {
                    try {
                        $process.Kill()
                        Start-Sleep -Milliseconds 100
                    } catch { }
                }
                
                # Clean up temp file with retry
                for ($i = 0; $i -lt 3; $i++) {
                    try {
                        if (Test-Path $tempFile) {
                            Remove-Item $tempFile -Force -ErrorAction Stop
                            break
                        }
                    } catch {
                        Start-Sleep -Milliseconds (50 * ($i + 1))
                    }
                }
            }
        }
        
        # Fallback: Return unknown volume
        Write-Log "Cannot determine current volume - using fallback method" "WARN"
        return -1
        
    } catch {
        Write-Log "Error getting microphone volume: $($_.Exception.Message)" "ERROR"
        return -1
    }
}

function Set-MicrophoneVolume {
    param($DeviceInfo, [int]$VolumePercent)
    
    try {
        Write-Log "Attempting to set microphone volume to $VolumePercent%" "INFO"
        
        # Method 1: Try SoundVolumeView
        $svv = "${env:ProgramFiles}\SoundVolumeView\SoundVolumeView.exe"
        $svvPortable = ".\SoundVolumeView.exe"
        
        $svcPath = $null
        if (Test-Path $svv) { $svcPath = $svv }
        elseif (Test-Path $svvPortable) { $svcPath = $svvPortable }
        
        if ($svcPath) {
            # Clean up any existing SoundVolumeView processes first
            Get-Process SoundVolumeView -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 100
            
            try {
                # Set volume using SoundVolumeView with timeout
                $process = Start-Process -FilePath $svcPath -ArgumentList "/SetVolume", "Microphone", $VolumePercent -PassThru -NoNewWindow
                $process | Wait-Process -Timeout 10 -ErrorAction Stop
                
                Write-Log "Volume set using SoundVolumeView" "INFO"
                return $true
            } catch {
                Write-Log "SoundVolumeView volume setting timeout or error: $($_.Exception.Message)" "WARN"
                # Clean up process if it's still running
                if ($process -and !$process.HasExited) {
                    $process | Stop-Process -Force -ErrorAction SilentlyContinue
                }
                # Continue to fallback methods
            }
        }
        
        # Method 2: Try registry manipulation
        # This is more complex and device-specific
        Write-Log "SoundVolumeView not available - trying alternative methods" "WARN"
        
        # Method 3: Use Windows built-in commands (limited effectiveness)
        # This might not work for specific devices but worth trying
        try {
            # Try using PowerShell's built-in audio control (Windows 10+)
            $audioCommand = @"
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class AudioEndpointVolume { [DllImport("ole32.dll")] public static extern int CoInitialize(IntPtr pvReserved); }'
"@
            # This approach requires more complex implementation
            Write-Log "Advanced audio control methods require additional implementation" "WARN"
            
        } catch {
            Write-Log "Built-in audio commands failed: $($_.Exception.Message)" "ERROR"
        }
        
        return $false
        
    } catch {
        Write-Log "Error setting microphone volume: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Install-SoundVolumeView {
    Write-Log "Checking for SoundVolumeView utility..." "INFO"
    
    $svvUrl = "https://www.nirsoft.net/utils/soundvolumeview.zip"
    $svvZip = "$env:TEMP\soundvolumeview.zip"
    $svvDir = "$env:TEMP\SoundVolumeView"
    
    try {
        if (-not (Test-Path ".\SoundVolumeView.exe")) {
            Write-Log "Downloading SoundVolumeView utility..." "INFO"
            
            # Create download directory
            if (-not (Test-Path $svvDir)) {
                New-Item -ItemType Directory -Path $svvDir -Force | Out-Null
            }
            
            # Download SoundVolumeView
            Invoke-WebRequest -Uri $svvUrl -OutFile $svvZip -UseBasicParsing
            
            # Extract
            Expand-Archive -Path $svvZip -DestinationPath $svvDir -Force
            
            # Copy to current directory
            Copy-Item "$svvDir\SoundVolumeView.exe" ".\SoundVolumeView.exe" -Force
            
            # Cleanup
            Remove-Item $svvZip -Force -ErrorAction SilentlyContinue
            Remove-Item $svvDir -Recurse -Force -ErrorAction SilentlyContinue
            
            Write-Log "SoundVolumeView installed successfully" "INFO"
            return $true
        } else {
            Write-Log "SoundVolumeView already available" "INFO"
            return $true
        }
    } catch {
        Write-Log "Failed to install SoundVolumeView: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Add-ToStartup {
    Write-Log "Adding to Windows startup..." "INFO"
    
    try {
        $scriptPath = $MyInvocation.ScriptName
        if (-not $scriptPath) {
            $scriptPath = $PSCommandPath
        }
        
        $startupCommand = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
        
        # Method 1: Registry startup entry
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
        Set-ItemProperty -Path $regPath -Name "MicrophoneVolumeMonitor" -Value $startupCommand
        
        Write-Log "Added to Windows startup registry" "INFO"
        return $true
        
    } catch {
        Write-Log "Failed to add to startup: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Start-VolumeMonitoring {
    Write-Log "Starting Microphone Volume Monitor..." "INFO"
    Write-Log "Target Device ID: $script:TargetDeviceId" "INFO"
    Write-Log "Target Volume: $script:TargetVolume%" "INFO"
    Write-Log "Polling Interval: $script:PollingInterval seconds" "INFO"
    
    # Install SoundVolumeView if needed
    Install-SoundVolumeView | Out-Null
    
    # Add to startup
    Add-ToStartup | Out-Null
    
    $deviceInfo = $null
    $consecutiveFailures = 0
    $maxFailures = 10
    
    while ($true) {
        try {
            # Find device if not already found or after failures
            if (-not $deviceInfo -or $consecutiveFailures -gt 5) {
                $deviceInfo = Find-TargetMicrophone
            }
            
            if ($deviceInfo.Found) {
                $consecutiveFailures = 0
                
                # Get current volume
                $currentVolume = Get-MicrophoneVolume -DeviceInfo $deviceInfo
                
                if ($currentVolume -ge 0) {
                    if ($currentVolume -lt $script:TargetVolume) {
                        Write-Log "Microphone volume is $currentVolume%, adjusting to $script:TargetVolume%" "WARN"
                        
                        $success = Set-MicrophoneVolume -DeviceInfo $deviceInfo -VolumePercent $script:TargetVolume
                        if ($success) {
                            Write-Log "Successfully adjusted microphone volume to $script:TargetVolume%" "INFO"
                        } else {
                            Write-Log "Failed to adjust microphone volume" "ERROR"
                            $consecutiveFailures++
                        }
                    } else {
                        Write-Log "Microphone volume is optimal: $currentVolume%" "INFO"
                    }
                } else {
                    Write-Log "Could not determine current volume, assuming adjustment needed" "WARN"
                    Set-MicrophoneVolume -DeviceInfo $deviceInfo -VolumePercent $script:TargetVolume | Out-Null
                }
            } else {
                Write-Log "Target microphone device not found, retrying..." "WARN"
                $consecutiveFailures++
                $deviceInfo = $null
            }
            
            if ($consecutiveFailures -gt $maxFailures) {
                Write-Log "Too many consecutive failures, extending retry interval" "ERROR"
                Start-Sleep -Seconds ($script:PollingInterval * 5)
                $consecutiveFailures = 0
            } else {
                Start-Sleep -Seconds $script:PollingInterval
            }
            
        } catch {
            Write-Log "Unexpected error in monitoring loop: $($_.Exception.Message)" "ERROR"
            $consecutiveFailures++
            Start-Sleep -Seconds $script:PollingInterval
        }
    }
}

# Main execution
if ($MyInvocation.InvocationName -ne '.') {
    try {
        Write-Log "=== Microphone Volume Monitor Starting ===" "INFO"
        
        # Ensure only one instance runs at a time
        if (-not (Test-SingleInstance)) {
            Write-Log "Failed to establish single instance - exiting" "ERROR"
            exit 1
        }
        
        Start-VolumeMonitoring
    } catch {
        Write-Log "Fatal error: $($_.Exception.Message)" "ERROR"
        Write-Log "=== Microphone Volume Monitor Stopped ===" "INFO"
        exit 1
    }
}