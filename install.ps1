#Requires -Version 5.1
# MFP Control Center - Installer (PowerShell)
# For HP LaserJet M1536dnf

param(
    [switch]$SkipBuild,
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "MFP Control Center - Install"

# Settings
$AppName = "MFP Control Center"
$InstallDir = "$env:ProgramFiles\MFP Control Center"
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$StartMenuPath = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TempDir = "$env:TEMP\MFPControlCenter_Install"

# URLs
$BuildToolsUrl = "https://aka.ms/vs/17/release/vs_BuildTools.exe"
$BuildToolsInstaller = "$TempDir\vs_BuildTools.exe"
$NetFrameworkUrl = "https://go.microsoft.com/fwlink/?linkid=2088631"
$NetFrameworkInstaller = "$TempDir\ndp48-web.exe"

function Write-Header {
    Clear-Host
    Write-Host ""
    Write-Host "=================================================================" -ForegroundColor Cyan
    Write-Host "         MFP Control Center - Installer                          " -ForegroundColor Cyan
    Write-Host "         For HP LaserJet M1536dnf                                " -ForegroundColor Cyan
    Write-Host "=================================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step($step, $total, $message) {
    Write-Host "[$step/$total] " -ForegroundColor Yellow -NoNewline
    Write-Host $message
}

function Write-Success($message) {
    Write-Host "[OK] " -ForegroundColor Green -NoNewline
    Write-Host $message
}

function Write-Warn($message) {
    Write-Host "[!] " -ForegroundColor Yellow -NoNewline
    Write-Host $message
}

function Write-Err($message) {
    Write-Host "[X] " -ForegroundColor Red -NoNewline
    Write-Host $message
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Install-NetFramework {
    Write-Host ""
    Write-Host "=================================================================" -ForegroundColor Magenta
    Write-Host "     Installing .NET Framework 4.8                               " -ForegroundColor Magenta
    Write-Host "     (reboot required after installation)                        " -ForegroundColor Magenta
    Write-Host "=================================================================" -ForegroundColor Magenta
    Write-Host ""

    if (-not (Test-Path $TempDir)) {
        New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
    }

    Write-Host "    Downloading .NET Framework 4.8..." -ForegroundColor Gray
    Write-Host ""

    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $NetFrameworkUrl -OutFile $NetFrameworkInstaller -UseBasicParsing
        $ProgressPreference = 'Continue'
    }
    catch {
        Write-Err "Failed to download .NET Framework!"
        Write-Host "    Download manually: https://dotnet.microsoft.com/download/dotnet-framework/net48" -ForegroundColor Yellow
        return $false
    }

    Write-Success "Installer downloaded"
    Write-Host ""
    Write-Host "    Starting .NET Framework 4.8 installation..." -ForegroundColor Gray
    Write-Host ""

    $process = Start-Process -FilePath $NetFrameworkInstaller -ArgumentList "/passive /norestart" -PassThru

    $spinner = @('|', '/', '-', '\')
    $i = 0
    while (-not $process.HasExited) {
        Write-Host "`r    Installing... $($spinner[$i % 4]) " -NoNewline
        Start-Sleep -Milliseconds 500
        $i++
    }

    $process.WaitForExit()
    Write-Host ""

    if ($process.ExitCode -eq 0) {
        Write-Success ".NET Framework 4.8 installed!"
        return $true
    }
    elseif ($process.ExitCode -eq 3010) {
        Write-Success ".NET Framework 4.8 installed!"
        Write-Warn "Reboot required!"
        Write-Host ""
        Write-Host "    Run installer again after reboot." -ForegroundColor Yellow
        $restart = Read-Host "    Reboot now? (Y/N)"
        if ($restart -eq "Y" -or $restart -eq "y") {
            Write-Host "    Rebooting in 10 seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds 10
            Restart-Computer -Force
        }
        return $false
    }
    elseif ($process.ExitCode -eq 1602) {
        Write-Warn "Installation cancelled by user"
        return $false
    }
    else {
        Write-Err "Installation error (code: $($process.ExitCode))"
        return $false
    }
}

function Find-MSBuild {
    $msbuildPaths = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    )

    foreach ($path in $msbuildPaths) {
        if (Test-Path $path) {
            return $path
        }
    }
    return $null
}

function Install-BuildTools {
    Write-Host ""
    Write-Host "=================================================================" -ForegroundColor Magenta
    Write-Host "     Installing Visual Studio Build Tools 2022                   " -ForegroundColor Magenta
    Write-Host "     (free, ~2-3 GB, takes 5-15 minutes)                         " -ForegroundColor Magenta
    Write-Host "=================================================================" -ForegroundColor Magenta
    Write-Host ""

    if (-not (Test-Path $TempDir)) {
        New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
    }

    Write-Host "    Downloading Build Tools installer..." -ForegroundColor Gray
    Write-Host "    URL: $BuildToolsUrl" -ForegroundColor DarkGray
    Write-Host ""

    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $BuildToolsUrl -OutFile $BuildToolsInstaller -UseBasicParsing
        $ProgressPreference = 'Continue'
    }
    catch {
        Write-Err "Failed to download Build Tools!"
        Write-Host "    Download manually: https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022" -ForegroundColor Yellow
        return $false
    }

    Write-Success "Installer downloaded"
    Write-Host ""
    Write-Host "    Starting Build Tools installation..." -ForegroundColor Gray
    Write-Host "    This may take 5-15 minutes. Please wait." -ForegroundColor Yellow
    Write-Host ""

    $installArgs = @(
        "--quiet",
        "--wait",
        "--norestart",
        "--nocache",
        "--add", "Microsoft.VisualStudio.Workload.ManagedDesktopBuildTools",
        "--add", "Microsoft.VisualStudio.Workload.MSBuildTools",
        "--add", "Microsoft.Net.Component.4.8.SDK",
        "--add", "Microsoft.Net.Component.4.8.TargetingPack",
        "--add", "Microsoft.VisualStudio.Component.NuGet.BuildTools",
        "--includeRecommended"
    )

    $process = Start-Process -FilePath $BuildToolsInstaller -ArgumentList $installArgs -PassThru

    $spinner = @('|', '/', '-', '\')
    $i = 0
    while (-not $process.HasExited) {
        Write-Host "`r    Installing... $($spinner[$i % 4]) " -NoNewline
        Start-Sleep -Milliseconds 500
        $i++
    }

    $process.WaitForExit()
    Write-Host ""

    if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
        Write-Success "Build Tools installed!"

        if ($process.ExitCode -eq 3010) {
            Write-Warn "Reboot required to complete installation."
            $restart = Read-Host "    Reboot now? (Y/N)"
            if ($restart -eq "Y" -or $restart -eq "y") {
                Write-Host "    Rebooting in 10 seconds..." -ForegroundColor Yellow
                Write-Host "    Run installer again after reboot." -ForegroundColor Yellow
                Start-Sleep -Seconds 10
                Restart-Computer -Force
                exit
            }
        }
        return $true
    }
    else {
        Write-Err "Build Tools installation error (code: $($process.ExitCode))"
        return $false
    }
}

function Install-App {
    Write-Header

    if (-not (Test-Administrator)) {
        Write-Warn "Administrator rights required!"
        Write-Host "    Restarting with admin rights..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Start-Process PowerShell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }

    Write-Success "Running as Administrator"
    Write-Host ""

    # Step 1: Check .NET Framework
    Write-Step 1 7 "Checking .NET Framework 4.8..."

    $netVersion = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue
    if ($null -eq $netVersion -or $netVersion.Release -lt 528040) {
        Write-Warn ".NET Framework 4.8 not installed!"
        Write-Host ""
        Write-Host "    .NET Framework 4.8 is required." -ForegroundColor Yellow
        Write-Host ""

        $installChoice = Read-Host "    Install .NET Framework 4.8 automatically? (Y/N)"

        if ($installChoice -eq "Y" -or $installChoice -eq "y") {
            $success = Install-NetFramework

            if (-not $success) {
                Write-Host ""
                Write-Host "    Install .NET Framework 4.8 manually and run installer again." -ForegroundColor Yellow
                Read-Host "    Press Enter to exit"
                exit 1
            }

            $netVersion = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue
            if ($null -eq $netVersion -or $netVersion.Release -lt 528040) {
                Write-Warn "Reboot required to complete .NET Framework installation."
                Write-Host "    Run installer again after reboot." -ForegroundColor Yellow
                Read-Host "    Press Enter to exit"
                exit 1
            }
        }
        else {
            Write-Host ""
            Write-Host "    Download manually: https://dotnet.microsoft.com/download/dotnet-framework/net48" -ForegroundColor Cyan
            Read-Host "    Press Enter to exit"
            exit 1
        }
    }
    Write-Success ".NET Framework 4.8 installed (version: $($netVersion.Release))"

    if (-not $SkipBuild) {
        # Step 2: Find MSBuild
        Write-Step 2 7 "Searching for MSBuild..."

        $msbuild = Find-MSBuild

        if (-not $msbuild) {
            Write-Warn "MSBuild not found!"
            Write-Host ""
            Write-Host "    Visual Studio or Build Tools not installed." -ForegroundColor Yellow
            Write-Host "    MSBuild from Build Tools is required to build the project." -ForegroundColor Yellow
            Write-Host ""

            $installChoice = Read-Host "    Install Build Tools automatically? (Y/N)"

            if ($installChoice -eq "Y" -or $installChoice -eq "y") {
                $success = Install-BuildTools

                if ($success) {
                    $msbuild = Find-MSBuild

                    if (-not $msbuild) {
                        Write-Err "MSBuild still not found after installation!"
                        Write-Host "    Try rebooting and running installer again." -ForegroundColor Yellow
                        Read-Host "    Press Enter to exit"
                        exit 1
                    }
                }
                else {
                    Write-Host ""
                    Write-Host "    Alternative:" -ForegroundColor Yellow
                    Write-Host "    1. Download Build Tools manually:" -ForegroundColor Gray
                    Write-Host "       https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022" -ForegroundColor Cyan
                    Write-Host "    2. Select '.NET desktop build tools' during installation" -ForegroundColor Gray
                    Write-Host "    3. Run this installer again" -ForegroundColor Gray
                    Read-Host "    Press Enter to exit"
                    exit 1
                }
            }
            else {
                Write-Host ""
                Write-Host "    Manual installation:" -ForegroundColor Yellow
                Write-Host "    1. Download: https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022" -ForegroundColor Cyan
                Write-Host "    2. Select '.NET desktop build tools' during installation" -ForegroundColor Gray
                Write-Host "    3. Run this installer again" -ForegroundColor Gray
                Read-Host "    Press Enter to exit"
                exit 1
            }
        }

        Write-Success "MSBuild found: $msbuild"

        # Step 3: Restore NuGet packages
        Write-Step 3 7 "Restoring NuGet packages..."

        Set-Location $ScriptDir

        $nugetPath = Join-Path $ScriptDir "nuget.exe"
        if (-not (Test-Path $nugetPath)) {
            Write-Host "    Downloading NuGet..." -ForegroundColor Gray
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile $nugetPath -UseBasicParsing
            $ProgressPreference = 'Continue'
        }

        $nugetOutput = & $nugetPath restore "MFPControlCenter.sln" -NonInteractive 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Err "Package restore failed!"
            Write-Host $nugetOutput -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
        Write-Success "NuGet packages restored"

        # Step 4: Build project
        Write-Step 4 7 "Building project (this may take a minute)..."

        $buildOutput = & $msbuild "MFPControlCenter.sln" /p:Configuration=Release /p:Platform="Any CPU" /t:Build /v:minimal /nologo 2>&1

        if ($LASTEXITCODE -ne 0) {
            Write-Err "Build failed!"
            Write-Host ""
            Write-Host "Compiler output:" -ForegroundColor Yellow
            Write-Host $buildOutput -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
        Write-Success "Project built successfully"
    }
    else {
        Write-Step 2 7 "Skipping MSBuild search (SkipBuild mode)..."
        Write-Step 3 7 "Skipping package restore..."
        Write-Step 4 7 "Skipping build..."
    }

    # Step 5: Copy files
    Write-Step 5 7 "Copying files to $InstallDir..."

    $sourcePath = Join-Path $ScriptDir "MFPControlCenter\bin\Release"
    if (-not (Test-Path $sourcePath)) {
        $sourcePath = Join-Path $ScriptDir "bin\Release"
    }
    if (-not (Test-Path $sourcePath)) {
        $sourcePath = Join-Path $ScriptDir "Release"
    }

    if (-not (Test-Path $sourcePath)) {
        Write-Err "Build output folder not found!"
        Write-Host "    Expected: MFPControlCenter\bin\Release" -ForegroundColor Yellow
        Read-Host "    Press Enter to exit"
        exit 1
    }

    $exeFile = Join-Path $sourcePath "MFPControlCenter.exe"
    if (-not (Test-Path $exeFile)) {
        Write-Err "MFPControlCenter.exe not found in $sourcePath"
        Read-Host "Press Enter to exit"
        exit 1
    }

    if (-not (Test-Path $InstallDir)) {
        New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    }

    Copy-Item -Path "$sourcePath\*" -Destination $InstallDir -Recurse -Force
    Write-Success "Files copied"

    # Step 6: Create shortcuts
    Write-Step 6 7 "Creating shortcuts..."

    $shell = New-Object -ComObject WScript.Shell

    $shortcut = $shell.CreateShortcut("$DesktopPath\$AppName.lnk")
    $shortcut.TargetPath = "$InstallDir\MFPControlCenter.exe"
    $shortcut.WorkingDirectory = $InstallDir
    $shortcut.Description = "HP LaserJet M1536dnf MFP Control Center"
    $shortcut.IconLocation = "$InstallDir\MFPControlCenter.exe,0"
    $shortcut.Save()

    $shortcut = $shell.CreateShortcut("$StartMenuPath\$AppName.lnk")
    $shortcut.TargetPath = "$InstallDir\MFPControlCenter.exe"
    $shortcut.WorkingDirectory = $InstallDir
    $shortcut.Description = "HP LaserJet M1536dnf MFP Control Center"
    $shortcut.IconLocation = "$InstallDir\MFPControlCenter.exe,0"
    $shortcut.Save()

    Write-Success "Shortcuts created (Desktop + Start Menu)"

    # Step 7: Create uninstaller
    Write-Step 7 7 "Creating uninstaller..."

    $uninstallBat = @"
@echo off
chcp 65001 >nul
echo =================================================================
echo         Uninstalling MFP Control Center
echo =================================================================
echo.
echo Removing program files...
rmdir /S /Q "%ProgramFiles%\MFP Control Center" 2>nul
echo Removing shortcuts...
del "%USERPROFILE%\Desktop\MFP Control Center.lnk" 2>nul
del "%ProgramData%\Microsoft\Windows\Start Menu\Programs\MFP Control Center.lnk" 2>nul
echo.
echo =================================================================
echo         MFP Control Center uninstalled successfully!
echo =================================================================
pause
"@
    $uninstallBat | Out-File -FilePath "$InstallDir\Uninstall.bat" -Encoding ASCII

    Write-Success "Uninstaller created"

    try {
        if (Test-Path $TempDir) {
            Remove-Item -Path $TempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch {
        # Ignore cleanup errors
    }

    # Done!
    Write-Host ""
    Write-Host "=================================================================" -ForegroundColor Green
    Write-Host "                                                                 " -ForegroundColor Green
    Write-Host "         INSTALLATION COMPLETED SUCCESSFULLY!                    " -ForegroundColor Green
    Write-Host "                                                                 " -ForegroundColor Green
    Write-Host "=================================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "    Installed to: " -NoNewline
    Write-Host $InstallDir -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    Desktop shortcut: " -NoNewline
    Write-Host "$AppName" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    To uninstall run: " -NoNewline
    Write-Host "$InstallDir\Uninstall.bat" -ForegroundColor Gray
    Write-Host ""
    Write-Host "=================================================================" -ForegroundColor Green

    $launch = Read-Host "    Launch program now? (Y/N)"
    if ($launch -eq "Y" -or $launch -eq "y" -or $launch -eq "") {
        Write-Host ""
        Write-Host "    Launching MFP Control Center..." -ForegroundColor Cyan
        Start-Process "$InstallDir\MFPControlCenter.exe"
    }
}

function Uninstall-App {
    Write-Header
    Write-Host "Uninstalling $AppName..." -ForegroundColor Yellow

    if (-not (Test-Administrator)) {
        Start-Process PowerShell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`" -Uninstall" -Verb RunAs
        exit
    }

    Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$DesktopPath\$AppName.lnk" -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$StartMenuPath\$AppName.lnk" -Force -ErrorAction SilentlyContinue

    Write-Success "$AppName uninstalled!"
}

# Main
if ($Uninstall) {
    Uninstall-App
}
else {
    Install-App
}

Write-Host ""
Read-Host "Press Enter to exit"
