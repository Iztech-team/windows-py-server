<#
.SYNOPSIS
    Baraka POS Printer Server - Windows Setup Script

.DESCRIPTION
    Fully self-contained Windows setup. Downloads everything from GitHub and
    installs ALL dependencies automatically:
    
     1. Downloads & installs Python 3.12 (if not found)
     2. Downloads & installs ADB Platform Tools (if not found)
     3. Downloads & installs WSA (Windows Subsystem for Android) if not found
     4. Downloads server files from GitHub (Iztech-team/windows-py-server)
     5. Installs Python packages
     6. Scans the network for thermal printers (port 9100)
     7. Enables WSA Developer Mode reminder
     8. Sets up WSA bridge (adb reverse) for Android POS access
     9. Sideloads APK into WSA (optional, via -ApkPath)
    10. Firewall, auto-start, management scripts
    11. Starts the server

    No manual installation required. Just download this single script and run it.

.PARAMETER ApkPath
    Optional path to an APK file to sideload into WSA after setup.

.NOTES
    Run as Administrator for best results (required for WSA install).
    Usage: Right-click -> "Run with PowerShell" or:
      powershell -ExecutionPolicy Bypass -File .\windows_setup.ps1
      powershell -ExecutionPolicy Bypass -File .\windows_setup.ps1 -ApkPath "C:\path\to\app.apk"
#>

param(
    [string]$ApkPath = ""
)

# ── Config ───────────────────────────────────────────────────
$ErrorActionPreference = "Continue"
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$ServerPort = 3006
$InstallDir = "$env:USERPROFILE\baraka-printer-server"
$DownloadsDir = "$env:TEMP\baraka-setup-downloads"
$TaskName = "BarakaPrinterServer"
$TotalSteps = 11

# Reboot-resume flag: saved when VM features are enabled and reboot is needed
$ResumeFlag = "$env:TEMP\baraka-setup-resume.json"

# Download URLs
$PythonVersion = "3.12.8"
$PythonInstallerURL = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe"
$ADBZipURL = "https://dl.google.com/android/repository/platform-tools-latest-windows.zip"
$ADBInstallDir = if ($isAdmin) { "C:\platform-tools" } else { "$env:USERPROFILE\platform-tools" }
$WSABuildsRepo = "MustardChef/WSABuilds"
$WSAPackagesRepo = "MustardChef/WSAPackages"
$WSATargetVersion = "2310.40000.2.0"
$WSATargetTag = "WSA_$WSATargetVersion"
$WSAMsixBundleName = "MicrosoftCorporationII.WindowsSubsystemForAndroid_${WSATargetVersion}_neutral_._8wekyb3d8bbwe.Msixbundle"
$WSADirectURL = "https://github.com/$WSAPackagesRepo/releases/download/$WSATargetTag/$WSAMsixBundleName"
$WSAInstallDir = "$env:LOCALAPPDATA\Baraka-WSA"
$SevenZipURL = "https://www.7-zip.org/a/7zr.exe"

# GitHub repo for server files
$GitHubRepo = "Iztech-team/windows-py-server"
$GitHubBranch = "main"
$GitHubRawBase = "https://raw.githubusercontent.com/$GitHubRepo/$GitHubBranch"

# ── Colors / Helpers ─────────────────────────────────────────
function Write-Banner($text) {
    Write-Host ""
    Write-Host ("=" * 56) -ForegroundColor Cyan
    Write-Host "  $text" -ForegroundColor White -BackgroundColor DarkCyan
    Write-Host ("=" * 56) -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step($num, $total, $text) {
    Write-Host ""
    Write-Host "[$num/$total] $text" -ForegroundColor Yellow
}

function Write-OK($text) {
    Write-Host "  [OK] $text" -ForegroundColor Green
}

function Write-Warn($text) {
    Write-Host "  [!]  $text" -ForegroundColor Yellow
}

function Write-Err($text) {
    Write-Host "  [X]  $text" -ForegroundColor Red
}

function Write-Info($text) {
    Write-Host "  -->  $text" -ForegroundColor Cyan
}

function Download-File($url, $dest, $label) {
    Write-Info "Downloading $label..."
    Write-Info "  URL: $url"
    try {
        # Use .NET WebClient (fast, reliable, no progress bar issues)
        try {
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
            $ProgressPreference = 'Continue'
        } catch {
            # Fallback
            $wc = New-Object System.Net.WebClient
            $wc.DownloadFile($url, $dest)
        }
        
        if (Test-Path $dest) {
            $size = [math]::Round((Get-Item $dest).Length / 1MB, 1)
            Write-OK "$label downloaded ($size MB)"
            return $true
        }
    } catch {
        Write-Err "Download failed: $_"
    }
    return $false
}

function Find-Python {
    # Check common Python commands in PATH
    foreach ($cmd in @("python", "python3", "py")) {
        try {
            $output = & $cmd --version 2>&1
            if ($output -match "Python (\d+)\.(\d+)") {
                $major = [int]$Matches[1]
                $minor = [int]$Matches[2]
                if ($major -ge 3 -and $minor -ge 8) {
                    return @{ Cmd = $cmd; Version = $output.ToString().Trim() }
                }
            }
        } catch {}
    }
    
    # Check common installation paths directly
    $commonPaths = @(
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe",
        "C:\Python312\python.exe",
        "C:\Python311\python.exe",
        "C:\Python310\python.exe",
        "$env:ProgramFiles\Python312\python.exe",
        "$env:ProgramFiles\Python311\python.exe"
    )
    foreach ($p in $commonPaths) {
        if (Test-Path $p) {
            try {
                $output = & $p --version 2>&1
                if ($output -match "Python (\d+)\.(\d+)") {
                    return @{ Cmd = $p; Version = $output.ToString().Trim() }
                }
            } catch {}
        }
    }
    
    return $null
}

function Find-ADB {
    # Check PATH first
    $adb = Get-Command adb -ErrorAction SilentlyContinue
    if ($adb) { return $adb.Source }
    
    # Check common locations
    $commonPaths = @(
        "$ADBInstallDir\adb.exe",
        "$env:LOCALAPPDATA\Android\Sdk\platform-tools\adb.exe",
        "C:\Android\platform-tools\adb.exe",
        "$env:USERPROFILE\platform-tools\adb.exe"
    )
    foreach ($p in $commonPaths) {
        if (Test-Path $p) { return $p }
    }
    
    return $null
}

function Refresh-Path {
    # Reload PATH from registry so newly installed programs are found
    $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machinePath;$userPath"
}

function Test-WSAInstalled {
    $pkg = Get-AppxPackage -Name "MicrosoftCorporationII.WindowsSubsystemForAndroid" -ErrorAction SilentlyContinue
    return ($null -ne $pkg)
}

function Test-VMFeaturesEnabled {
    try {
        $vmPlatform = (Get-WindowsOptionalFeature -Online -FeatureName "VirtualMachinePlatform" -ErrorAction Stop).State -eq "Enabled"
        $hypervisor = (Get-WindowsOptionalFeature -Online -FeatureName "HypervisorPlatform" -ErrorAction Stop).State -eq "Enabled"
        return ($vmPlatform -and $hypervisor)
    } catch {
        # Get-WindowsOptionalFeature requires elevation; assume unknown
        return $null
    }
}

function Get-WSABuildsLatestRelease {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $releaseInfo = Invoke-RestMethod -Uri "https://api.github.com/repos/$WSABuildsRepo/releases/latest" -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = 'Continue'

        # Prefer: x64, GApps, NoAmazon, .7z (the cleanest build)
        foreach ($asset in $releaseInfo.assets) {
            $name = $asset.name
            if ($name -match "x64" -and $name -match "GApps" -and $name -match "NoAmazon" -and $name -match "\.(7z|zip)$" -and $name -notmatch "magisk|KernelSU") {
                return @{
                    Url     = $asset.browser_download_url
                    Name    = $asset.name
                    Size    = [math]::Round($asset.size / 1MB, 0)
                    Tag     = $releaseInfo.tag_name
                    Is7z    = $name -match "\.7z$"
                }
            }
        }

        # Fallback: any x64 GApps build
        foreach ($asset in $releaseInfo.assets) {
            $name = $asset.name
            if ($name -match "x64" -and $name -match "GApps" -and $name -match "\.(7z|zip)$" -and $name -notmatch "magisk|KernelSU") {
                return @{
                    Url     = $asset.browser_download_url
                    Name    = $asset.name
                    Size    = [math]::Round($asset.size / 1MB, 0)
                    Tag     = $releaseInfo.tag_name
                    Is7z    = $name -match "\.7z$"
                }
            }
        }

        # Last resort: largest archive file
        $archives = $releaseInfo.assets | Where-Object { $_.name -match "\.(7z|zip)$" } | Sort-Object size -Descending
        if ($archives -and $archives.Count -gt 0) {
            $best = $archives[0]
            return @{
                Url     = $best.browser_download_url
                Name    = $best.name
                Size    = [math]::Round($best.size / 1MB, 0)
                Tag     = $releaseInfo.tag_name
                Is7z    = $best.name -match "\.7z$"
            }
        }
    } catch {
        Write-Err "Failed to query WSABuilds releases: $_"
    }
    return $null
}

function Find-7Zip {
    # Check PATH
    $sz = Get-Command "7z" -ErrorAction SilentlyContinue
    if ($sz) { return $sz.Source }

    # Common install locations
    $paths = @(
        "${env:ProgramFiles}\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe",
        "$env:LOCALAPPDATA\7-Zip\7z.exe"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) { return $p }
    }

    return $null
}

function Ensure-7Zip {
    $sz = Find-7Zip
    if ($sz) { return $sz }

    # Download 7zr.exe (standalone console extractor, ~600KB)
    $sevenZrPath = Join-Path $DownloadsDir "7zr.exe"
    if (-not (Test-Path $sevenZrPath)) {
        Write-Info "Downloading 7-Zip extractor (needed for WSA archive)..."
        $downloaded = Download-File $SevenZipURL $sevenZrPath "7-Zip extractor"
        if (-not $downloaded) { return $null }
    }
    return $sevenZrPath
}

function Save-ResumeState {
    param([hashtable]$State)
    $State | ConvertTo-Json | Set-Content $ResumeFlag -Encoding UTF8
}

function Load-ResumeState {
    if (Test-Path $ResumeFlag) {
        try {
            $state = Get-Content $ResumeFlag -Raw | ConvertFrom-Json
            return $state
        } catch {}
    }
    return $null
}

function Clear-ResumeState {
    if (Test-Path $ResumeFlag) {
        Remove-Item $ResumeFlag -Force -ErrorAction SilentlyContinue
    }
}

# ─────────────────────────────────────────────────────────────
Write-Banner "Baraka POS Printer Server - Windows Setup v3.0"
# ─────────────────────────────────────────────────────────────

$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    Write-OK "Running as Administrator"
} else {
    Write-Warn "Not running as Administrator. Some features may be limited."
    Write-Info "For full setup (WSA install, firewall), right-click PowerShell -> Run as Administrator"
}

# ── Check for reboot-resume ──────────────────────────────────
$resumeState = Load-ResumeState
$skipToWSA = $false
if ($resumeState) {
    Write-Info "Detected resume after reboot..."
    if ($resumeState.stage -eq "vm-features-enabled") {
        if (Test-VMFeaturesEnabled) {
            Write-OK "VM features are now active after reboot."
            $skipToWSA = $true
        } else {
            Write-Warn "VM features still not active. They may need another reboot."
        }
    }
    Clear-ResumeState
}

# ═════════════════════════════════════════════════════════════
# STEP 0: KILL OLD SERVER & CLEAN UP
# ═════════════════════════════════════════════════════════════
Write-Step 0 $TotalSteps "Stopping any running server and cleaning up..."

# Kill anything on port 3006
$listeners = netstat -ano 2>$null | Select-String ":$ServerPort" | Select-String "LISTENING"
if ($listeners) {
    foreach ($line in $listeners) {
        $parts = $line.ToString().Trim() -split '\s+'
        $procPid = $parts[-1]
        if ($procPid -match '^\d+$' -and [int]$procPid -gt 0) {
            try {
                $procName = (Get-Process -Id $procPid -ErrorAction SilentlyContinue).ProcessName
                taskkill /PID $procPid /F 2>$null | Out-Null
                Write-OK "Killed process $procName (PID: $procPid) on port $ServerPort"
            } catch {}
        }
    }
} else {
    Write-Info "No server running on port $ServerPort"
}

# Also kill any python printer_server.py processes
Get-Process -Name "python*" -ErrorAction SilentlyContinue | ForEach-Object {
    try {
        $cmdLine = (Get-CimInstance Win32_Process -Filter "ProcessId = $($_.Id)" -ErrorAction SilentlyContinue).CommandLine
        if ($cmdLine -match "printer_server") {
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
            Write-OK "Killed old printer_server process (PID: $($_.Id))"
        }
    } catch {}
}

# Remove old scheduled task if exists
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue 2>$null

# Remove old startup shortcut if exists
$startupFolder = [Environment]::GetFolderPath("Startup")
$oldShortcut = Join-Path $startupFolder "Baraka Printer Server.lnk"
if (Test-Path $oldShortcut) {
    Remove-Item $oldShortcut -Force -ErrorAction SilentlyContinue
    Write-OK "Removed old startup shortcut"
}

# Wipe old installation completely (fresh start)
if (Test-Path $InstallDir) {
    Write-Info "Removing old installation: $InstallDir"
    Remove-Item $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-OK "Old installation removed"
}

Write-OK "Cleanup complete"

# Create downloads directory
New-Item -ItemType Directory -Path $DownloadsDir -Force | Out-Null

# ═════════════════════════════════════════════════════════════
# STEP 1: PYTHON - Check / Download / Install
# ═════════════════════════════════════════════════════════════
Write-Step 1 $TotalSteps "Checking Python installation..."

$pythonInfo = Find-Python

if ($pythonInfo) {
    $pythonCmd = $pythonInfo.Cmd
    Write-OK "Found $($pythonInfo.Version) ($pythonCmd)"
} else {
    Write-Warn "Python 3.8+ not found. Installing automatically..."
    
    # Download Python installer
    $pythonInstaller = Join-Path $DownloadsDir "python-$PythonVersion-amd64.exe"
    
    if (-not (Test-Path $pythonInstaller)) {
        $downloaded = Download-File $PythonInstallerURL $pythonInstaller "Python $PythonVersion"
        if (-not $downloaded) {
            Write-Err "Failed to download Python installer!"
            Write-Info "Check your internet connection and try again."
            Read-Host "Press Enter to exit"
            exit 1
        }
    } else {
        Write-Info "Python installer already downloaded."
    }
    
    # Install Python silently
    Write-Info "Installing Python $PythonVersion (this may take 1-2 minutes)..."
    Write-Info "  Options: Add to PATH, pip included"
    
    $installArgs = @(
        "/quiet",
        "InstallAllUsers=0",
        "PrependPath=1",
        "Include_pip=1",
        "Include_test=0",
        "Include_launcher=1",
        "Include_doc=0"
    )
    
    # If admin, install for all users
    if ($isAdmin) {
        $installArgs[1] = "InstallAllUsers=1"
    }
    
    $process = Start-Process -FilePath $pythonInstaller -ArgumentList $installArgs -Wait -PassThru
    
    if ($process.ExitCode -eq 0) {
        Write-OK "Python $PythonVersion installed successfully"
    } else {
        Write-Warn "Installer exited with code $($process.ExitCode) (checking anyway...)"
    }
    
    # Refresh PATH and find Python again
    Refresh-Path
    Start-Sleep -Seconds 3
    
    $pythonInfo = Find-Python
    if ($pythonInfo) {
        $pythonCmd = $pythonInfo.Cmd
        Write-OK "Verified: $($pythonInfo.Version) ($pythonCmd)"
    } else {
        Write-Err "Python still not found after installation!"
        Write-Host ""
        Write-Host "  This usually means PATH was not updated." -ForegroundColor Yellow
        Write-Host "  Please:" -ForegroundColor Yellow
        Write-Host "    1. Close this window" -ForegroundColor White
        Write-Host "    2. Open a NEW PowerShell window" -ForegroundColor White
        Write-Host "    3. Run this script again" -ForegroundColor White
        Write-Host ""
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Ensure pip is available
Write-Info "Ensuring pip is up to date..."
& $pythonCmd -m ensurepip --default-pip 2>&1 | Out-Null
& $pythonCmd -m pip install --upgrade pip --quiet 2>&1 | Out-Null
Write-OK "pip is ready"

# ═════════════════════════════════════════════════════════════
# STEP 2: ADB - Check / Download / Install
# ═════════════════════════════════════════════════════════════
Write-Step 2 $TotalSteps "Checking ADB (Android Platform Tools)..."

$adbPath = Find-ADB

if ($adbPath) {
    Write-OK "ADB found: $adbPath"
} else {
    Write-Warn "ADB not found. Downloading Android Platform Tools..."
    
    $adbZip = Join-Path $DownloadsDir "platform-tools-windows.zip"
    
    if (-not (Test-Path $adbZip)) {
        $downloaded = Download-File $ADBZipURL $adbZip "Android Platform Tools"
        if (-not $downloaded) {
            Write-Warn "Failed to download ADB. WSA bridge will be unavailable."
            Write-Info "You can download manually from:"
            Write-Info "  https://developer.android.com/tools/releases/platform-tools"
        }
    } else {
        Write-Info "ADB zip already downloaded."
    }
    
    if (Test-Path $adbZip) {
        Write-Info "Extracting to $ADBInstallDir..."
        
        try {
            # Extract to temp first (zip contains a platform-tools/ folder)
            $extractTemp = Join-Path $DownloadsDir "adb-extract"
            if (Test-Path $extractTemp) { Remove-Item $extractTemp -Recurse -Force }
            
            Expand-Archive -Path $adbZip -DestinationPath $extractTemp -Force
            
            # Move platform-tools folder to final location
            if (Test-Path $ADBInstallDir) { Remove-Item $ADBInstallDir -Recurse -Force }
            Move-Item -Path (Join-Path $extractTemp "platform-tools") -Destination $ADBInstallDir -Force
            
            Write-OK "ADB extracted to $ADBInstallDir"
            
            # Add to PATH
            $targetScope = if ($isAdmin) { "Machine" } else { "User" }
            $currentPath = [Environment]::GetEnvironmentVariable("Path", $targetScope)
            
            if ($currentPath -notlike "*$ADBInstallDir*") {
                $newPath = "$currentPath;$ADBInstallDir"
                [Environment]::SetEnvironmentVariable("Path", $newPath, $targetScope)
                $env:Path = "$env:Path;$ADBInstallDir"
                Write-OK "Added $ADBInstallDir to PATH ($targetScope)"
            }
            
            $adbPath = Join-Path $ADBInstallDir "adb.exe"
            if (Test-Path $adbPath) {
                Write-OK "ADB is ready: $adbPath"
            }
            
            # Cleanup temp extraction
            Remove-Item $extractTemp -Recurse -Force -ErrorAction SilentlyContinue
        } catch {
            Write-Err "Failed to extract ADB: $_"
            $adbPath = $null
        }
    }
}

# ═════════════════════════════════════════════════════════════
# STEP 3: WSA - Check / Enable VM Features / Download / Install
# ═════════════════════════════════════════════════════════════
Write-Step 3 $TotalSteps "Checking Windows Subsystem for Android (WSA)..."

$wsaInstalled = Test-WSAInstalled

if ($wsaInstalled) {
    Write-OK "WSA is already installed"
} else {
    Write-Warn "WSA not found. Will attempt to install..."

    $canInstallWSA = $true

    # 3a: Check/enable VM features (requires admin)
    $vmStatus = Test-VMFeaturesEnabled
    if ($vmStatus -eq $false) {
        if (-not $isAdmin) {
            Write-Err "Enabling VM features requires Administrator privileges."
            Write-Info "Please re-run this script as Administrator to install WSA."
            Write-Info "  Right-click PowerShell -> Run as Administrator"
            Write-Info "Skipping WSA installation for now. Other steps will continue."
            $canInstallWSA = $false
        } else {
            Write-Info "Enabling VirtualMachinePlatform and HypervisorPlatform..."
            $needsReboot = $false

            try {
                $result1 = Enable-WindowsOptionalFeature -Online -FeatureName "VirtualMachinePlatform" -NoRestart -ErrorAction Stop
                if ($result1.RestartNeeded) { $needsReboot = $true }
                Write-OK "VirtualMachinePlatform enabled"
            } catch {
                Write-Err "Failed to enable VirtualMachinePlatform: $_"
            }

            try {
                $result2 = Enable-WindowsOptionalFeature -Online -FeatureName "HypervisorPlatform" -NoRestart -ErrorAction Stop
                if ($result2.RestartNeeded) { $needsReboot = $true }
                Write-OK "HypervisorPlatform enabled"
            } catch {
                Write-Err "Failed to enable HypervisorPlatform: $_"
            }

            if ($needsReboot) {
                Write-Warn "A reboot is required to activate VM features."
                Write-Info "After rebooting, run this script again -- it will resume automatically."

                Save-ResumeState @{
                    stage       = "vm-features-enabled"
                    scriptPath  = $MyInvocation.MyCommand.Path
                    apkPath     = $ApkPath
                }

                $scriptFullPath = $MyInvocation.MyCommand.Path
                $relaunchCmd = "powershell.exe -ExecutionPolicy Bypass -File `"$scriptFullPath`""
                if ($ApkPath) {
                    $relaunchCmd += " -ApkPath `"$ApkPath`""
                }
                try {
                    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce" `
                        -Name "BarakaSetupResume" -Value $relaunchCmd -ErrorAction Stop
                    Write-OK "Registered auto-resume after reboot"
                } catch {
                    Write-Warn "Could not register auto-resume. Please re-run the script manually after reboot."
                }

                $rebootNow = Read-Host "Reboot now? (Y/n)"
                if ($rebootNow -ne "n" -and $rebootNow -ne "N") {
                    Write-Info "Rebooting in 5 seconds..."
                    Start-Sleep -Seconds 5
                    Restart-Computer -Force
                }
                Write-Info "Please reboot manually, then run this script again."
                exit 0
            }
        }
    } elseif ($vmStatus -eq $null -and -not $isAdmin) {
        Write-Warn "Cannot check VM features without Administrator privileges."
        Write-Info "WSA install requires admin. Skipping for now."
        $canInstallWSA = $false
    } else {
        Write-OK "VM features already enabled"
    }

    # 3b: Find local WSA folder (with Install.ps1) or msixbundle
    if ($canInstallWSA) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
        $wsaInstalledNow = $false

        # Priority 1: Look for a WSA folder with Install.ps1 (MagiskOnWSA / WSABuilds)
        $wsaFolder = Get-ChildItem -Path $scriptDir -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "WSA" } |
            Select-Object -First 1

        # Check nested folder (some builds have outer/inner folder structure)
        if ($wsaFolder) {
            $installScript = Get-ChildItem -Path $wsaFolder.FullName -Filter "Install.ps1" -Recurse -Depth 2 -ErrorAction SilentlyContinue |
                Select-Object -First 1
        }

        if ($installScript) {
            $wsaInstallDir = Split-Path $installScript.FullName -Parent
            Write-OK "Found local WSA build: $($wsaFolder.Name)"
            Write-Info "Running WSA Install.ps1 from: $wsaInstallDir"
            
            try {
                $installProcess = Start-Process powershell.exe `
                    -ArgumentList "-ExecutionPolicy Bypass -File `"$($installScript.FullName)`"" `
                    -WorkingDirectory $wsaInstallDir `
                    -Wait -PassThru
                
                Start-Sleep -Seconds 5
                if (Test-WSAInstalled) {
                    $wsaInstalled = $true
                    $wsaInstalledNow = $true
                    Write-OK "WSA installed successfully via Install.ps1!"
                } else {
                    Write-Warn "Install.ps1 finished but WSA not detected. It may need a moment..."
                    Start-Sleep -Seconds 10
                    if (Test-WSAInstalled) {
                        $wsaInstalled = $true
                        $wsaInstalledNow = $true
                        Write-OK "WSA installed successfully!"
                    }
                }
            } catch {
                Write-Err "Failed to run WSA Install.ps1: $_"
            }
        } else {
            # Priority 2: Look for msixbundle
            $localMsix = Get-ChildItem -Path $scriptDir -Filter "*.Msixbundle" -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match "WindowsSubsystemForAndroid" } |
                Select-Object -First 1

            if ($localMsix) {
                Write-OK "Found local WSA package: $($localMsix.Name)"
                
                # Install VCLibs dependency first
                Write-Info "Installing VCLibs dependency..."
                try {
                    $vcLibsUrl = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
                    $vcLibsPath = Join-Path $DownloadsDir "Microsoft.VCLibs.x64.14.00.Desktop.appx"
                    if (-not (Test-Path $vcLibsPath)) {
                        $ProgressPreference = 'SilentlyContinue'
                        Invoke-WebRequest -Uri $vcLibsUrl -OutFile $vcLibsPath -UseBasicParsing
                        $ProgressPreference = 'Continue'
                    }
                    Add-AppxPackage -Path $vcLibsPath -ErrorAction Stop
                    Write-OK "VCLibs installed"
                } catch {
                    Write-Warn "VCLibs install note: $_"
                }

                Write-Info "Installing WSA from msixbundle..."
                try {
                    Add-AppxPackage -Path $localMsix.FullName -ErrorAction Stop
                    Start-Sleep -Seconds 5
                    if (Test-WSAInstalled) {
                        $wsaInstalled = $true
                        $wsaInstalledNow = $true
                        Write-OK "WSA installed successfully!"
                    }
                } catch {
                    Write-Err "Failed to install WSA msixbundle: $_"
                }
            } else {
                Write-Warn "No local WSA folder or .Msixbundle found next to the script."
                Write-Info "Place a WSA build folder (with Install.ps1) or .Msixbundle next to this script."
            }
        }

        # Launch WSA if just installed
        if ($wsaInstalledNow) {
            Write-Info "Launching WSA for first-time initialization..."
            Start-Sleep -Seconds 5
            try {
                Start-Process "shell:AppsFolder\MicrosoftCorporationII.WindowsSubsystemForAndroid_8wekyb3d8bbwe!App" -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 15
                Write-OK "WSA launched"
            } catch {
                Write-Warn "Could not auto-launch WSA. Open it from the Start Menu."
            }
        }
    }
}

# ═════════════════════════════════════════════════════════════
# STEP 4: Set up installation directory & download files from GitHub
# ═════════════════════════════════════════════════════════════
Write-Step 4 $TotalSteps "Downloading server files from GitHub..."
Write-Info "Install path: $InstallDir"
Write-Info "Source: https://github.com/$GitHubRepo"

New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
Write-OK "Created $InstallDir"

$serverFiles = @(
    "printer_server.py",
    "printer_discovery.py",
    "print_queue.py",
    "wsa_bridge.py",
    "requirements.txt",
    ".env.example"
)

$downloadedCount = 0
$ProgressPreference = 'SilentlyContinue'
foreach ($file in $serverFiles) {
    $url = "$GitHubRawBase/$file"
    $dst = Join-Path $InstallDir $file
    try {
        Invoke-WebRequest -Uri $url -OutFile $dst -UseBasicParsing -ErrorAction Stop
        $downloadedCount++
        Write-OK "  $file"
    } catch {
        Write-Err "  Failed to download $file : $_"
    }
}
$ProgressPreference = 'Continue'
Write-OK "Downloaded $downloadedCount/$($serverFiles.Count) server files from GitHub"

# Set up .env config from template
$envDst = Join-Path $InstallDir ".env"
$envExampleDst = Join-Path $InstallDir ".env.example"

if (-not (Test-Path $envDst)) {
    if (Test-Path $envExampleDst) {
        Copy-Item $envExampleDst $envDst -Force
        Write-OK "Created .env from template"
    } else {
        @"
SERVER_HOST=0.0.0.0
SERVER_PORT=$ServerPort
PRINTER_REGISTRY=printer_registry.json
SCAN_ON_STARTUP=true
TEST_PRINT_ON_STARTUP=true
WSA_BRIDGE_ENABLED=true
WSA_ADB_PORT=58526
LOG_LEVEL=INFO
"@ | Set-Content $envDst -Encoding UTF8
        Write-OK "Created default .env"
    }
} else {
    Write-OK ".env already exists (preserved)"
}

# ═════════════════════════════════════════════════════════════
# STEP 5: Install Python packages
# ═════════════════════════════════════════════════════════════
Write-Step 5 $TotalSteps "Installing Python packages..."
Write-Info "Installing Python packages (fastapi, uvicorn, python-escpos, Pillow...)..."
Write-Info "This may take 1-2 minutes on first run..."
$reqFile = Join-Path $InstallDir "requirements.txt"
if (Test-Path $reqFile) {
    $pipOutput = & $pythonCmd -m pip install -r $reqFile 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-OK "All Python packages installed"
    } else {
        Write-Err "pip install had issues:"
        $pipOutput | Select-Object -Last 5 | ForEach-Object { Write-Info $_ }
        Write-Info "Trying individual packages as fallback..."
        $packages = @("fastapi", "uvicorn", "python-escpos", "python-dotenv", "Pillow", "werkzeug", "python-multipart")
        foreach ($pkg in $packages) {
            & $pythonCmd -m pip install $pkg --quiet 2>&1 | Out-Null
            if ($LASTEXITCODE -eq 0) {
                Write-OK "  Installed $pkg"
            } else {
                Write-Err "  Failed: $pkg"
            }
        }
    }
} else {
    # No requirements.txt - install directly
    Write-Info "Installing packages directly..."
    & $pythonCmd -m pip install fastapi uvicorn python-escpos python-dotenv Pillow werkzeug python-multipart 2>&1 | Out-Null
    Write-OK "Python packages installed"
}

# ═════════════════════════════════════════════════════════════
# STEP 6: Network Printer Discovery
# ═════════════════════════════════════════════════════════════
Write-Step 6 $TotalSteps "Scanning network for thermal printers (port 9100)..."

Write-Info "This takes about 5-10 seconds..."

$localIP = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
    Where-Object { $_.IPAddress -match "^(192\.168\.|10\.|172\.(1[6-9]|2[0-9]|3[01])\.)" } |
    Select-Object -First 1).IPAddress

if ($localIP) {
    $subnet = ($localIP -split '\.')[0..2] -join '.'
    Write-Info "Local IP: $localIP"
    Write-Info "Scanning subnet: $subnet.0/24"
    
    # Use .NET RunspacePool (lightweight threads, NOT heavy Start-Job processes)
    $foundPrinters = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
    
    $scriptBlock = {
        param($targetIP, $port, $timeoutMs)
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $connect = $tcp.BeginConnect($targetIP, $port, $null, $null)
            $wait = $connect.AsyncWaitHandle.WaitOne($timeoutMs, $false)
            if ($wait -and $tcp.Connected) {
                $tcp.EndConnect($connect)
                $tcp.Close()
                return $targetIP
            }
            $tcp.Close()
            $tcp.Dispose()
        } catch {}
        return $null
    }
    
    # Create runspace pool (50 concurrent threads, very lightweight)
    $pool = [RunspaceFactory]::CreateRunspacePool(1, 50)
    $pool.Open()
    
    $runspaces = @()
    1..254 | ForEach-Object {
        $ip = "$subnet.$_"
        $ps = [PowerShell]::Create().AddScript($scriptBlock).AddArgument($ip).AddArgument(9100).AddArgument(1500)
        $ps.RunspacePool = $pool
        $runspaces += @{ Pipe = $ps; Handle = $ps.BeginInvoke() }
    }
    
    Write-Info "Waiting for scan to complete..."
    
    # Collect results
    $printerList = @()
    foreach ($rs in $runspaces) {
        try {
            $result = $rs.Pipe.EndInvoke($rs.Handle)
            if ($result -and $result[0]) {
                $printerList += $result[0]
            }
        } catch {}
        $rs.Pipe.Dispose()
    }
    $pool.Close()
    $pool.Dispose()
    
    if ($printerList.Count -gt 0) {
        Write-OK "Found $($printerList.Count) printer(s):"
        foreach ($p in $printerList) {
            $arpLine = (arp -a $p 2>$null | Select-String $p)
            $mac = "unknown"
            if ($arpLine -match "([0-9a-fA-F]{2}[:-]){5}[0-9a-fA-F]{2}") {
                $mac = $Matches[0].ToUpper() -replace '-',':'
            }
            Write-Info "$p (MAC: $mac)"
        }
    } else {
        Write-Warn "No printers found on network."
        Write-Info "Server will re-scan on startup. Or use: POST /printers/discover"
    }
} else {
    Write-Warn "Could not detect local IP. Server will scan on startup."
}

# ═════════════════════════════════════════════════════════════
# STEP 7: WSA Developer Mode (auto-enable)
# ═════════════════════════════════════════════════════════════
Write-Step 7 $TotalSteps "Configuring WSA Developer Mode..."

if ($wsaInstalled) {
    if (-not $adbPath) { $adbPath = Find-ADB }

    $wsaDevReady = $false
    $wsaPkg = Get-AppxPackage -Name "MicrosoftCorporationII.WindowsSubsystemForAndroid" -ErrorAction SilentlyContinue

    # Open WSA Settings app so user can see it and enable Developer Mode
    Write-Info "Opening WSA Settings..."
    $wsaSettingsLaunched = $false

    if ($wsaPkg) {
        # Try launching WSA Settings app directly
        $wsaSettingsExe = Join-Path $wsaPkg.InstallLocation "WsaSettings\WsaSettings.exe"
        if (Test-Path $wsaSettingsExe) {
            Start-Process $wsaSettingsExe -ErrorAction SilentlyContinue
            $wsaSettingsLaunched = $true
            Write-OK "Opened WSA Settings (WsaSettings.exe)"
        }
    }

    if (-not $wsaSettingsLaunched) {
        try {
            Start-Process "shell:AppsFolder\MicrosoftCorporationII.WindowsSubsystemForAndroid_8wekyb3d8bbwe!App" -ErrorAction SilentlyContinue
            $wsaSettingsLaunched = $true
            Write-OK "Opened WSA Settings (shell:AppsFolder)"
        } catch {}
    }

    if (-not $wsaSettingsLaunched) {
        # Last resort: try WsaClient
        if ($wsaPkg) {
            $wsaExe = Join-Path $wsaPkg.InstallLocation "WsaClient\WsaClient.exe"
            if (Test-Path $wsaExe) {
                Start-Process $wsaExe -ErrorAction SilentlyContinue
                $wsaSettingsLaunched = $true
                Write-OK "Opened WSA via WsaClient.exe"
            }
        }
    }

    Write-Host ""
    Write-Host "  +================================================+" -ForegroundColor Yellow
    Write-Host "  |  ENABLE WSA DEVELOPER MODE                      |" -ForegroundColor Yellow
    Write-Host "  +================================================+" -ForegroundColor Yellow
    Write-Host "  |                                                  |" -ForegroundColor Yellow
    Write-Host "  |  In the WSA Settings window that just opened:    |" -ForegroundColor White
    Write-Host "  |                                                  |" -ForegroundColor Yellow
    Write-Host "  |  1. Click 'Developer' on the left sidebar       |" -ForegroundColor White
    Write-Host "  |  2. Toggle 'Developer mode' to ON               |" -ForegroundColor White
    Write-Host "  |  3. Come back here and press Enter               |" -ForegroundColor White
    Write-Host "  |                                                  |" -ForegroundColor Yellow
    Write-Host "  +================================================+" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter after enabling Developer Mode"

    # Now try connecting via ADB
    if ($adbPath) {
        Write-Info "Connecting to WSA via ADB..."
        & $adbPath kill-server 2>$null | Out-Null
        Start-Sleep -Seconds 2
        & $adbPath start-server 2>$null | Out-Null

        $maxAttempts = 10
        $attempt = 0
        while ($attempt -lt $maxAttempts -and -not $wsaDevReady) {
            $attempt++
            Start-Sleep -Seconds 3
            $testConnect = & $adbPath connect 127.0.0.1:58526 2>&1 | Out-String
            if ($testConnect -match "connected|already") {
                Write-OK "WSA ADB connected on port 58526!"
                $wsaDevReady = $true
            } else {
                Write-Info "  Waiting for WSA ADB... (attempt $attempt/$maxAttempts)"
            }
        }

        if (-not $wsaDevReady) {
            Write-Warn "Could not connect to WSA on port 58526."
            Write-Info "Make sure Developer Mode is ON in WSA Settings."
            Write-Info "WSA bridge and APK sideload may not work."
        }
    } else {
        Write-Warn "ADB not available. Cannot verify WSA Developer Mode."
    }
} else {
    Write-Info "WSA not installed -- skipping Developer Mode check."
}

# ═════════════════════════════════════════════════════════════
# STEP 8: WSA Bridge (adb reverse)
# ═════════════════════════════════════════════════════════════
Write-Step 8 $TotalSteps "Setting up WSA bridge (adb reverse)..."

if (-not $adbPath) { $adbPath = Find-ADB }

if ($adbPath -and $wsaDevReady) {
    Write-OK "ADB: $adbPath"
    
    & $adbPath reverse --remove-all 2>&1 | Out-Null
    $reverseResult = & $adbPath reverse "tcp:$ServerPort" "tcp:$ServerPort" 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-OK "adb reverse tcp:$ServerPort -> tcp:$ServerPort"
        Write-OK "WSA apps can reach the server at localhost:$ServerPort"
    } else {
        Write-Warn "adb reverse failed. Server will auto-retry on startup."
    }
} elseif ($adbPath -and -not $wsaDevReady) {
    Write-Warn "WSA ADB not reachable. Bridge setup skipped."
    Write-Info "Server will auto-retry the bridge on startup."
} else {
    Write-Warn "ADB not available. WSA bridge disabled."
}

# ═════════════════════════════════════════════════════════════
# STEP 9: Sideload APK (optional)
# ═════════════════════════════════════════════════════════════
Write-Step 9 $TotalSteps "APK Sideload..."

# Auto-detect APK in script folder if none specified
if (-not $ApkPath -or $ApkPath -eq "") {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $localApk = Get-ChildItem -Path $scriptDir -Filter "*.apk" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($localApk) {
        $ApkPath = $localApk.FullName
        Write-OK "Found local APK: $($localApk.Name)"
    }
}

if ($ApkPath -and $ApkPath -ne "") {
    if (-not (Test-Path $ApkPath)) {
        Write-Err "APK file not found: $ApkPath"
    } elseif (-not $adbPath) {
        Write-Err "ADB not available -- cannot sideload APK."
    } elseif (-not $wsaDevReady) {
        Write-Err "WSA ADB not connected. Cannot sideload APK."
        Write-Info "Enable Developer Mode in WSA, then run:"
        Write-Info "  $adbPath connect 127.0.0.1:58526"
        Write-Info "  $adbPath install `"$ApkPath`""
    } else {
        Write-Info "Installing APK: $(Split-Path $ApkPath -Leaf)..."
        $installResult = & $adbPath install -r $ApkPath 2>&1
        $installOutput = ($installResult | Out-String).Trim()

        if ($installOutput -match "Success") {
            Write-OK "APK installed successfully into WSA!"
            Write-Info "The app should now appear in your Start Menu."
        } elseif ($installOutput -match "unauthorized") {
            Write-Warn "Device unauthorized. Need to allow USB debugging."
            & $adbPath kill-server 2>$null | Out-Null
            Start-Sleep -Seconds 2
            & $adbPath start-server 2>$null | Out-Null
            & $adbPath connect 127.0.0.1:58526 2>$null | Out-Null
            Start-Sleep -Seconds 3

            Write-Host ""
            Write-Host "  +================================================+" -ForegroundColor Yellow
            Write-Host "  |  A popup should appear asking to allow USB      |" -ForegroundColor White
            Write-Host "  |  debugging. Check 'Always allow' and click OK.  |" -ForegroundColor White
            Write-Host "  +================================================+" -ForegroundColor Yellow
            Write-Host ""
            Read-Host "Press Enter after allowing USB debugging"

            & $adbPath connect 127.0.0.1:58526 2>$null | Out-Null
            Start-Sleep -Seconds 2
            Write-Info "Retrying APK install..."
            $retryResult = & $adbPath install -r $ApkPath 2>&1
            $retryOutput = ($retryResult | Out-String).Trim()

            if ($retryOutput -match "Success") {
                Write-OK "APK installed successfully into WSA!"
                Write-Info "The app should now appear in your Start Menu."
            } else {
                Write-Err "APK install still failed: $retryOutput"
                Write-Info "You can try manually: $adbPath install `"$ApkPath`""
            }
        } else {
            Write-Err "APK install failed: $installOutput"
            Write-Info "You can try manually: $adbPath install `"$ApkPath`""
        }
    }
} else {
    Write-Info "No APK specified. Skipping sideload."
    Write-Info "To sideload later: .\windows_setup.ps1 -ApkPath `"C:\path\to\app.apk`""
}

# ═════════════════════════════════════════════════════════════
# STEP 10: Firewall + Auto-Start + Scripts
# ═════════════════════════════════════════════════════════════
Write-Step 10 $TotalSteps "Configuring firewall, auto-start, and management scripts..."

# ── Firewall ─────────────────────────────────────────────────
if ($isAdmin) {
    try {
        Remove-NetFirewallRule -DisplayName "Baraka Printer Server" -ErrorAction SilentlyContinue
        New-NetFirewallRule `
            -DisplayName "Baraka Printer Server" `
            -Direction Inbound `
            -Protocol TCP `
            -LocalPort $ServerPort `
            -Action Allow `
            -Profile Any | Out-Null
        Write-OK "Firewall rule added for port $ServerPort"
    } catch {
        Write-Warn "Could not add firewall rule."
    }
} else {
    # Try netsh as fallback (sometimes works without elevation)
    Write-Info "Attempting to add firewall rule via netsh..."
    $fwResult = netsh advfirewall firewall add rule name="Baraka Printer Server" dir=in action=allow protocol=TCP localport=$ServerPort 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-OK "Firewall rule added for port $ServerPort"
    } else {
        Write-Warn "Could not add firewall rule (need Administrator)."
        Write-Info "If other devices can't connect, Windows will show a popup"
        Write-Info "asking to allow access - click 'Allow' when it appears."
    }
}

# ── Resolve full python path for batch scripts ───────────────
$pythonFullPath = $pythonCmd
try {
    $resolved = Get-Command $pythonCmd -ErrorAction SilentlyContinue
    if ($resolved) { $pythonFullPath = $resolved.Source }
} catch {}

# ── start_server.bat ─────────────────────────────────────────
$startBat = Join-Path $InstallDir "start_server.bat"
@"
@echo off
title Baraka Printer Server
cd /d "$InstallDir"
echo ================================================
echo   Baraka POS Printer Server
echo ================================================
echo.
echo Starting server on port $ServerPort...
echo Press Ctrl+C to stop.
echo.
"$pythonFullPath" printer_server.py
echo.
echo Server stopped.
pause
"@ | Set-Content $startBat -Encoding ASCII
Write-OK "Created start_server.bat"

# ── start_server_hidden.vbs (silent background launcher) ─────
$hiddenVbs = Join-Path $InstallDir "start_server_hidden.vbs"
$pythonwPath = $pythonFullPath -replace '\\python\.exe$', '\pythonw.exe'
# If pythonw.exe doesn't exist, fall back to python.exe via wscript CreateObject
@"
' Baraka Printer Server - Silent Background Launcher
' Starts the server with NO visible window. Output goes to server.log.
Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = "$InstallDir"
Set objFSO = CreateObject("Scripting.FileSystemObject")

pythonwPath = "$pythonwPath"
pythonPath  = "$pythonFullPath"
logFile     = "$InstallDir\server.log"

If objFSO.FileExists(pythonwPath) Then
    objShell.Run Chr(34) & pythonwPath & Chr(34) & " printer_server.py", 0, False
Else
    objShell.Run "cmd /c " & Chr(34) & pythonPath & Chr(34) & " printer_server.py > " & Chr(34) & logFile & Chr(34) & " 2>&1", 0, False
End If
"@ | Set-Content $hiddenVbs -Encoding ASCII
Write-OK "Created start_server_hidden.vbs (silent launcher)"

# ── stop_server.bat ──────────────────────────────────────────
$stopBat = Join-Path $InstallDir "stop_server.bat"
@"
@echo off
echo Stopping Baraka Printer Server...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :$ServerPort ^| findstr LISTENING') do (
    echo Killing process %%a...
    taskkill /PID %%a /F 2>nul
)
echo Done.
timeout /t 3
"@ | Set-Content $stopBat -Encoding ASCII
Write-OK "Created stop_server.bat"

# ── Auto-start on login (HIDDEN, no console window) ─────────
if ($isAdmin) {
    # Method 1: Scheduled Task using wscript + vbs (fully hidden)
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        
        $action = New-ScheduledTaskAction `
            -Execute "wscript.exe" `
            -Argument "`"$hiddenVbs`"" `
            -WorkingDirectory $InstallDir
        
        $trigger = New-ScheduledTaskTrigger -AtLogOn
        $settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -RestartCount 3 `
            -RestartInterval (New-TimeSpan -Minutes 1)
        
        Register-ScheduledTask `
            -TaskName $TaskName `
            -Action $action `
            -Trigger $trigger `
            -Settings $settings `
            -Description "Baraka POS Printer Server - auto-start hidden on login" `
            -RunLevel Limited | Out-Null
        
        Write-OK "Scheduled task: $TaskName (auto-starts HIDDEN on login)"
    } catch {
        Write-Warn "Could not register scheduled task: $_"
    }
} else {
    # Method 2: Startup folder shortcut to hidden VBS (no admin needed)
    Write-Info "Creating auto-start shortcut in Startup folder..."
    
    $startupFolder = [Environment]::GetFolderPath("Startup")
    $shortcutPath = Join-Path $startupFolder "Baraka Printer Server.lnk"
    
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "wscript.exe"
        $shortcut.Arguments = "`"$hiddenVbs`""
        $shortcut.WorkingDirectory = $InstallDir
        $shortcut.Description = "Baraka POS Printer Server (hidden)"
        $shortcut.WindowStyle = 7  # Minimized
        $shortcut.Save()
        
        Write-OK "Auto-start shortcut created in Startup folder (HIDDEN)"
        Write-Info "Location: $shortcutPath"
        Write-Info "Server will start silently on every login (no CMD window)"
    } catch {
        Write-Warn "Could not create startup shortcut: $_"
        Write-Info "To auto-start manually: copy start_server.bat to shell:startup"
    }
}

# ═════════════════════════════════════════════════════════════
# STEP 11: Summary & Launch
# ═════════════════════════════════════════════════════════════
Write-Step 11 $TotalSteps "Setup complete!"

# Cleanup downloads
Remove-Item $DownloadsDir -Recurse -Force -ErrorAction SilentlyContinue
Write-OK "Temporary files cleaned up"

Write-Host ""
Write-Host ("=" * 56) -ForegroundColor Green
Write-Host "  SETUP COMPLETE - Baraka Printer Server Ready!" -ForegroundColor White -BackgroundColor DarkGreen
Write-Host ("=" * 56) -ForegroundColor Green
Write-Host ""
Write-Host "  Installed:" -ForegroundColor Cyan
Write-Host "  ────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Python:      $($pythonInfo.Version)" -ForegroundColor White
if ($adbPath) {
    Write-Host "  ADB:         $adbPath" -ForegroundColor White
} else {
    Write-Host "  ADB:         Not installed" -ForegroundColor Yellow
}
if ($wsaInstalled) {
    Write-Host "  WSA:         Installed" -ForegroundColor Green
} else {
    Write-Host "  WSA:         Not installed" -ForegroundColor Yellow
}
Write-Host "  Server:      $InstallDir" -ForegroundColor White
Write-Host ""
Write-Host "  Access:" -ForegroundColor Cyan
Write-Host "  ────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Local:       http://localhost:$ServerPort" -ForegroundColor Green
if ($localIP) {
    Write-Host "  Network:     http://${localIP}:$ServerPort" -ForegroundColor Green
}
Write-Host "  Health:      http://localhost:$ServerPort/health" -ForegroundColor Green
if ($wsaInstalled) {
    Write-Host "  WSA (APK):   http://localhost:$ServerPort (via adb reverse)" -ForegroundColor Green
}
Write-Host ""
Write-Host "  Scripts:" -ForegroundColor Cyan
Write-Host "  ────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Start:       $startBat" -ForegroundColor White
Write-Host "  Start (bg):  $hiddenVbs" -ForegroundColor White
Write-Host "  Stop:        $stopBat" -ForegroundColor White
Write-Host "  Log file:    $InstallDir\server.log" -ForegroundColor White
if ($isAdmin) {
    Write-Host "  Auto-start:  Scheduled Task ($TaskName) - HIDDEN" -ForegroundColor Green
} else {
    Write-Host "  Auto-start:  Startup folder shortcut - HIDDEN" -ForegroundColor Green
}
Write-Host "  ────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host ""

$startNow = Read-Host "Start the server now? (Y/n)"
if ($startNow -ne "n" -and $startNow -ne "N") {
    Write-Host ""
    Write-Host "Starting Baraka Printer Server..." -ForegroundColor Green
    Write-Host "Press Ctrl+C to stop." -ForegroundColor Yellow
    Write-Host ""
    
    Set-Location $InstallDir
    & $pythonCmd printer_server.py
} else {
    Write-Host ""
    Write-Host "To start later, double-click:" -ForegroundColor Yellow
    Write-Host "  $startBat" -ForegroundColor Cyan
    Write-Host ""
}
