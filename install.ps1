#Requires -Version 5.1
# MFP Control Center - Установщик (PowerShell)
# Для HP LaserJet M1536dnf

param(
    [switch]$SkipBuild,
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$Host.UI.RawUI.WindowTitle = "MFP Control Center - Установка"

# Настройки
$AppName = "MFP Control Center"
$InstallDir = "$env:ProgramFiles\MFP Control Center"
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$StartMenuPath = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TempDir = "$env:TEMP\MFPControlCenter_Install"

# URL для скачивания Build Tools
$BuildToolsUrl = "https://aka.ms/vs/17/release/vs_BuildTools.exe"
$BuildToolsInstaller = "$TempDir\vs_BuildTools.exe"

# URL для скачивания .NET Framework 4.8
$NetFrameworkUrl = "https://go.microsoft.com/fwlink/?linkid=2088631"
$NetFrameworkInstaller = "$TempDir\ndp48-web.exe"

function Write-Header {
    Clear-Host
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║         MFP Control Center - Установщик                       ║" -ForegroundColor Cyan
    Write-Host "║         Для HP LaserJet M1536dnf                              ║" -ForegroundColor Cyan
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
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
    Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
    Write-Host "║     Установка .NET Framework 4.8                              ║" -ForegroundColor Magenta
    Write-Host "║     (потребуется перезагрузка после установки)                ║" -ForegroundColor Magenta
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Magenta
    Write-Host ""

    # Создание временной папки
    if (-not (Test-Path $TempDir)) {
        New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
    }

    # Скачивание установщика
    Write-Host "    Скачивание .NET Framework 4.8..." -ForegroundColor Gray
    Write-Host "    Это веб-установщик (~1.5 МБ), остальное скачается при установке." -ForegroundColor DarkGray
    Write-Host ""

    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $NetFrameworkUrl -OutFile $NetFrameworkInstaller -UseBasicParsing
        $ProgressPreference = 'Continue'
    }
    catch {
        Write-Err "Не удалось скачать .NET Framework!"
        Write-Host "    Скачайте вручную: https://dotnet.microsoft.com/download/dotnet-framework/net48" -ForegroundColor Yellow
        return $false
    }

    Write-Success "Установщик скачан"
    Write-Host ""
    Write-Host "    Запуск установки .NET Framework 4.8..." -ForegroundColor Gray
    Write-Host "    Следуйте инструкциям установщика." -ForegroundColor Yellow
    Write-Host ""

    # Запуск установщика (интерактивно, чтобы пользователь видел прогресс)
    $process = Start-Process -FilePath $NetFrameworkInstaller -ArgumentList "/passive /norestart" -PassThru

    # Анимация ожидания
    $spinner = @('|', '/', '-', '\')
    $i = 0
    while (-not $process.HasExited) {
        Write-Host "`r    Установка... $($spinner[$i % 4]) " -NoNewline
        Start-Sleep -Milliseconds 500
        $i++
    }

    $process.WaitForExit()
    Write-Host ""

    if ($process.ExitCode -eq 0) {
        Write-Success ".NET Framework 4.8 установлен!"
        return $true
    }
    elseif ($process.ExitCode -eq 3010) {
        Write-Success ".NET Framework 4.8 установлен!"
        Write-Warn "Требуется перезагрузка компьютера!"
        Write-Host ""
        Write-Host "    После перезагрузки запустите установщик снова." -ForegroundColor Yellow
        $restart = Read-Host "    Перезагрузить сейчас? (Y/N)"
        if ($restart -eq "Y" -or $restart -eq "y" -or $restart -eq "Д" -or $restart -eq "д") {
            Write-Host "    Перезагрузка через 10 секунд..." -ForegroundColor Yellow
            Start-Sleep -Seconds 10
            Restart-Computer -Force
        }
        return $false
    }
    elseif ($process.ExitCode -eq 1602) {
        Write-Warn "Установка отменена пользователем"
        return $false
    }
    else {
        Write-Err "Ошибка установки .NET Framework (код: $($process.ExitCode))"
        return $false
    }
}

function Find-MSBuild {
    $msbuildPaths = @(
        # Visual Studio 2022
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
        # Visual Studio 2019
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
    Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
    Write-Host "║     Установка Visual Studio Build Tools 2022                  ║" -ForegroundColor Magenta
    Write-Host "║     (бесплатно, ~2-3 ГБ, займёт 5-15 минут)                    ║" -ForegroundColor Magenta
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Magenta
    Write-Host ""

    # Создание временной папки
    if (-not (Test-Path $TempDir)) {
        New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
    }

    # Скачивание установщика
    Write-Host "    Скачивание установщика Build Tools..." -ForegroundColor Gray
    Write-Host "    URL: $BuildToolsUrl" -ForegroundColor DarkGray
    Write-Host ""

    try {
        # Прогресс-бар для скачивания
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $BuildToolsUrl -OutFile $BuildToolsInstaller -UseBasicParsing
        $ProgressPreference = 'Continue'
    }
    catch {
        Write-Err "Не удалось скачать Build Tools!"
        Write-Host "    Скачайте вручную: https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022" -ForegroundColor Yellow
        return $false
    }

    Write-Success "Установщик скачан"
    Write-Host ""
    Write-Host "    Запуск установки Build Tools..." -ForegroundColor Gray
    Write-Host "    Это может занять 5-15 минут. Пожалуйста, подождите." -ForegroundColor Yellow
    Write-Host ""

    # Установка с нужными компонентами для WPF/.NET Framework
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

    # Анимация ожидания
    $spinner = @('|', '/', '-', '\')
    $i = 0
    while (-not $process.HasExited) {
        Write-Host "`r    Установка... $($spinner[$i % 4]) " -NoNewline
        Start-Sleep -Milliseconds 500
        $i++
    }

    $process.WaitForExit()
    Write-Host ""

    if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
        Write-Success "Build Tools установлены!"

        if ($process.ExitCode -eq 3010) {
            Write-Warn "Требуется перезагрузка компьютера для завершения установки."
            $restart = Read-Host "    Перезагрузить сейчас? (Y/N)"
            if ($restart -eq "Y" -or $restart -eq "y" -or $restart -eq "Д" -or $restart -eq "д") {
                Write-Host "    Перезагрузка через 10 секунд..." -ForegroundColor Yellow
                Write-Host "    После перезагрузки запустите установщик снова." -ForegroundColor Yellow
                Start-Sleep -Seconds 10
                Restart-Computer -Force
                exit
            }
        }
        return $true
    }
    else {
        Write-Err "Ошибка установки Build Tools (код: $($process.ExitCode))"
        return $false
    }
}

function Install-App {
    Write-Header

    # Проверка прав администратора
    if (-not (Test-Administrator)) {
        Write-Warn "Требуются права администратора!"
        Write-Host "    Перезапуск с правами администратора..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Start-Process PowerShell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }

    Write-Success "Запущено с правами администратора"
    Write-Host ""

    # ═══════════════════════════════════════════════════════════════
    # 1. Проверка .NET Framework
    # ═══════════════════════════════════════════════════════════════
    Write-Step 1 7 "Проверка .NET Framework 4.8..."

    $netVersion = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue
    if ($null -eq $netVersion -or $netVersion.Release -lt 528040) {
        Write-Warn ".NET Framework 4.8 не установлен!"
        Write-Host ""
        Write-Host "    Для работы программы требуется .NET Framework 4.8." -ForegroundColor Yellow
        Write-Host ""

        $installChoice = Read-Host "    Установить .NET Framework 4.8 автоматически? (Y/N)"

        if ($installChoice -eq "Y" -or $installChoice -eq "y" -or $installChoice -eq "Д" -or $installChoice -eq "д") {
            $success = Install-NetFramework

            if (-not $success) {
                Write-Host ""
                Write-Host "    Установите .NET Framework 4.8 вручную и запустите установщик снова." -ForegroundColor Yellow
                Read-Host "    Нажмите Enter для выхода"
                exit 1
            }

            # Перепроверка после установки
            $netVersion = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue
            if ($null -eq $netVersion -or $netVersion.Release -lt 528040) {
                Write-Warn "Требуется перезагрузка для завершения установки .NET Framework."
                Write-Host "    После перезагрузки запустите установщик снова." -ForegroundColor Yellow
                Read-Host "    Нажмите Enter для выхода"
                exit 1
            }
        }
        else {
            Write-Host ""
            Write-Host "    Скачайте вручную: https://dotnet.microsoft.com/download/dotnet-framework/net48" -ForegroundColor Cyan
            Read-Host "    Нажмите Enter для выхода"
            exit 1
        }
    }
    Write-Success ".NET Framework 4.8 установлен (версия: $($netVersion.Release))"

    if (-not $SkipBuild) {
        # ═══════════════════════════════════════════════════════════════
        # 2. Поиск MSBuild / Установка Build Tools
        # ═══════════════════════════════════════════════════════════════
        Write-Step 2 7 "Поиск MSBuild..."

        $msbuild = Find-MSBuild

        if (-not $msbuild) {
            Write-Warn "MSBuild не найден!"
            Write-Host ""
            Write-Host "    Visual Studio или Build Tools не установлены." -ForegroundColor Yellow
            Write-Host "    Для сборки программы нужен MSBuild из Build Tools." -ForegroundColor Yellow
            Write-Host ""

            $installChoice = Read-Host "    Установить Build Tools автоматически? (Y/N)"

            if ($installChoice -eq "Y" -or $installChoice -eq "y" -or $installChoice -eq "Д" -or $installChoice -eq "д") {
                $success = Install-BuildTools

                if ($success) {
                    # Повторный поиск MSBuild после установки
                    $msbuild = Find-MSBuild

                    if (-not $msbuild) {
                        Write-Err "MSBuild всё ещё не найден после установки!"
                        Write-Host "    Попробуйте перезагрузить компьютер и запустить установщик снова." -ForegroundColor Yellow
                        Read-Host "    Нажмите Enter для выхода"
                        exit 1
                    }
                }
                else {
                    Write-Host ""
                    Write-Host "    Альтернативный вариант:" -ForegroundColor Yellow
                    Write-Host "    1. Скачайте Build Tools вручную:" -ForegroundColor Gray
                    Write-Host "       https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022" -ForegroundColor Cyan
                    Write-Host "    2. При установке выберите '.NET desktop build tools'" -ForegroundColor Gray
                    Write-Host "    3. После установки запустите этот скрипт снова" -ForegroundColor Gray
                    Read-Host "    Нажмите Enter для выхода"
                    exit 1
                }
            }
            else {
                Write-Host ""
                Write-Host "    Для установки вручную:" -ForegroundColor Yellow
                Write-Host "    1. Скачайте: https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022" -ForegroundColor Cyan
                Write-Host "    2. При установке выберите '.NET desktop build tools'" -ForegroundColor Gray
                Write-Host "    3. Запустите этот скрипт снова" -ForegroundColor Gray
                Read-Host "    Нажмите Enter для выхода"
                exit 1
            }
        }

        Write-Success "MSBuild найден: $msbuild"

        # ═══════════════════════════════════════════════════════════════
        # 3. Восстановление NuGet пакетов
        # ═══════════════════════════════════════════════════════════════
        Write-Step 3 7 "Восстановление NuGet пакетов..."

        Set-Location $ScriptDir

        $nugetPath = Join-Path $ScriptDir "nuget.exe"
        if (-not (Test-Path $nugetPath)) {
            Write-Host "    Скачивание NuGet..." -ForegroundColor Gray
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile $nugetPath -UseBasicParsing
            $ProgressPreference = 'Continue'
        }

        $nugetOutput = & $nugetPath restore "MFPControlCenter.sln" -NonInteractive 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Err "Ошибка восстановления пакетов!"
            Write-Host $nugetOutput -ForegroundColor Red
            Read-Host "Нажмите Enter для выхода"
            exit 1
        }
        Write-Success "NuGet пакеты восстановлены"

        # ═══════════════════════════════════════════════════════════════
        # 4. Сборка проекта
        # ═══════════════════════════════════════════════════════════════
        Write-Step 4 7 "Сборка проекта (это может занять минуту)..."

        $buildOutput = & $msbuild "MFPControlCenter.sln" /p:Configuration=Release /p:Platform="Any CPU" /t:Build /v:minimal /nologo 2>&1

        if ($LASTEXITCODE -ne 0) {
            Write-Err "Ошибка сборки!"
            Write-Host ""
            Write-Host "Вывод компилятора:" -ForegroundColor Yellow
            Write-Host $buildOutput -ForegroundColor Red
            Read-Host "Нажмите Enter для выхода"
            exit 1
        }
        Write-Success "Проект успешно собран"
    }
    else {
        Write-Step 2 7 "Пропуск поиска MSBuild (режим SkipBuild)..."
        Write-Step 3 7 "Пропуск восстановления пакетов..."
        Write-Step 4 7 "Пропуск сборки..."
    }

    # ═══════════════════════════════════════════════════════════════
    # 5. Копирование файлов
    # ═══════════════════════════════════════════════════════════════
    Write-Step 5 7 "Копирование файлов в $InstallDir..."

    # Поиск папки с собранными файлами
    $sourcePath = Join-Path $ScriptDir "MFPControlCenter\bin\Release"
    if (-not (Test-Path $sourcePath)) {
        $sourcePath = Join-Path $ScriptDir "bin\Release"
    }
    if (-not (Test-Path $sourcePath)) {
        $sourcePath = Join-Path $ScriptDir "Release"
    }

    if (-not (Test-Path $sourcePath)) {
        Write-Err "Не найдена папка с собранными файлами!"
        Write-Host "    Ожидалось: MFPControlCenter\bin\Release" -ForegroundColor Yellow
        Read-Host "    Нажмите Enter для выхода"
        exit 1
    }

    # Проверка наличия exe файла
    $exeFile = Join-Path $sourcePath "MFPControlCenter.exe"
    if (-not (Test-Path $exeFile)) {
        Write-Err "Файл MFPControlCenter.exe не найден в $sourcePath"
        Read-Host "Нажмите Enter для выхода"
        exit 1
    }

    # Создание папки установки
    if (-not (Test-Path $InstallDir)) {
        New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    }

    # Копирование файлов
    Copy-Item -Path "$sourcePath\*" -Destination $InstallDir -Recurse -Force
    Write-Success "Файлы скопированы"

    # ═══════════════════════════════════════════════════════════════
    # 6. Создание ярлыков
    # ═══════════════════════════════════════════════════════════════
    Write-Step 6 7 "Создание ярлыков..."

    $shell = New-Object -ComObject WScript.Shell

    # Ярлык на рабочем столе
    $shortcut = $shell.CreateShortcut("$DesktopPath\$AppName.lnk")
    $shortcut.TargetPath = "$InstallDir\MFPControlCenter.exe"
    $shortcut.WorkingDirectory = $InstallDir
    $shortcut.Description = "Центр управления МФУ HP LaserJet M1536dnf"
    $shortcut.IconLocation = "$InstallDir\MFPControlCenter.exe,0"
    $shortcut.Save()

    # Ярлык в меню Пуск
    $shortcut = $shell.CreateShortcut("$StartMenuPath\$AppName.lnk")
    $shortcut.TargetPath = "$InstallDir\MFPControlCenter.exe"
    $shortcut.WorkingDirectory = $InstallDir
    $shortcut.Description = "Центр управления МФУ HP LaserJet M1536dnf"
    $shortcut.IconLocation = "$InstallDir\MFPControlCenter.exe,0"
    $shortcut.Save()

    Write-Success "Ярлыки созданы (рабочий стол + меню Пуск)"

    # ═══════════════════════════════════════════════════════════════
    # 7. Создание деинсталлятора
    # ═══════════════════════════════════════════════════════════════
    Write-Step 7 7 "Создание файла удаления..."

    $uninstallBat = @"
@echo off
chcp 65001 >nul
echo ═══════════════════════════════════════════════════════════════
echo         Удаление MFP Control Center
echo ═══════════════════════════════════════════════════════════════
echo.
echo Удаление файлов программы...
rmdir /S /Q "%ProgramFiles%\MFP Control Center" 2>nul
echo Удаление ярлыков...
del "%USERPROFILE%\Desktop\MFP Control Center.lnk" 2>nul
del "%ProgramData%\Microsoft\Windows\Start Menu\Programs\MFP Control Center.lnk" 2>nul
echo.
echo ═══════════════════════════════════════════════════════════════
echo         MFP Control Center успешно удалён!
echo ═══════════════════════════════════════════════════════════════
pause
"@
    $uninstallBat | Out-File -FilePath "$InstallDir\Удалить программу.bat" -Encoding UTF8

    Write-Success "Файл удаления создан"

    # Очистка временных файлов
    if (Test-Path $TempDir) {
        Remove-Item -Path $TempDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # ═══════════════════════════════════════════════════════════════
    # ГОТОВО!
    # ═══════════════════════════════════════════════════════════════
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                                                               ║" -ForegroundColor Green
    Write-Host "║         УСТАНОВКА УСПЕШНО ЗАВЕРШЕНА!                          ║" -ForegroundColor Green
    Write-Host "║                                                               ║" -ForegroundColor Green
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "    Программа установлена: " -NoNewline
    Write-Host $InstallDir -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    Ярлык создан на рабочем столе: " -NoNewline
    Write-Host "$AppName" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    Для удаления запустите: " -NoNewline
    Write-Host "$InstallDir\Удалить программу.bat" -ForegroundColor Gray
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green

    $launch = Read-Host "    Запустить программу сейчас? (Y/N)"
    if ($launch -eq "Y" -or $launch -eq "y" -or $launch -eq "Д" -or $launch -eq "д" -or $launch -eq "") {
        Write-Host ""
        Write-Host "    Запуск MFP Control Center..." -ForegroundColor Cyan
        Start-Process "$InstallDir\MFPControlCenter.exe"
    }
}

function Uninstall-App {
    Write-Header
    Write-Host "Удаление $AppName..." -ForegroundColor Yellow

    if (-not (Test-Administrator)) {
        Start-Process PowerShell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`" -Uninstall" -Verb RunAs
        exit
    }

    Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$DesktopPath\$AppName.lnk" -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$StartMenuPath\$AppName.lnk" -Force -ErrorAction SilentlyContinue

    Write-Success "$AppName удалён!"
}

# ═══════════════════════════════════════════════════════════════
# ЗАПУСК
# ═══════════════════════════════════════════════════════════════
if ($Uninstall) {
    Uninstall-App
}
else {
    Install-App
}

Write-Host ""
Read-Host "Нажмите Enter для выхода"
