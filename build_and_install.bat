@echo off
chcp 65001 >nul
title MFP Control Center - Установка

echo ═══════════════════════════════════════════════════════════════
echo         MFP Control Center - Установщик
echo         Для HP LaserJet M1536dnf
echo ═══════════════════════════════════════════════════════════════
echo.

:: Проверка прав администратора
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [!] Требуются права администратора для установки.
    echo [!] Перезапуск с правами администратора...
    echo.
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Переменные
set "INSTALL_DIR=%ProgramFiles%\MFP Control Center"
set "DESKTOP=%USERPROFILE%\Desktop"
set "START_MENU=%ProgramData%\Microsoft\Windows\Start Menu\Programs"
set "SCRIPT_DIR=%~dp0"

echo [1/6] Проверка .NET Framework 4.8...
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release 2>nul | find "528040" >nul
if %errorLevel% neq 0 (
    reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release 2>nul | find "528049" >nul
)
if %errorLevel% neq 0 (
    echo [!] .NET Framework 4.8 не найден.
    echo [!] Скачайте и установите с: https://dotnet.microsoft.com/download/dotnet-framework/net48
    echo.
    pause
    start https://dotnet.microsoft.com/download/dotnet-framework/net48
    exit /b 1
)
echo [OK] .NET Framework 4.8 установлен

echo.
echo [2/6] Поиск MSBuild...

:: Поиск MSBuild в разных местах
set "MSBUILD="

:: Visual Studio 2022
if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD=%ProgramFiles%\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
)
if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD=%ProgramFiles%\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
)
if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD=%ProgramFiles%\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
)

:: Visual Studio 2019
if "%MSBUILD%"=="" (
    if exist "%ProgramFiles(x86)%\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe" (
        set "MSBUILD=%ProgramFiles(x86)%\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
    )
    if exist "%ProgramFiles(x86)%\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe" (
        set "MSBUILD=%ProgramFiles(x86)%\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe"
    )
)

:: Build Tools
if "%MSBUILD%"=="" (
    if exist "%ProgramFiles(x86)%\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" (
        set "MSBUILD=%ProgramFiles(x86)%\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    )
    if exist "%ProgramFiles(x86)%\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe" (
        set "MSBUILD=%ProgramFiles(x86)%\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    )
)

if "%MSBUILD%"=="" (
    echo [!] MSBuild не найден!
    echo [!] Установите Visual Studio 2019/2022 или Build Tools.
    echo.
    echo Скачать Build Tools: https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022
    echo.
    pause
    exit /b 1
)
echo [OK] MSBuild найден

echo.
echo [3/6] Восстановление NuGet пакетов...
cd /d "%SCRIPT_DIR%"

:: Скачивание NuGet если его нет
if not exist "%SCRIPT_DIR%nuget.exe" (
    echo      Скачивание NuGet...
    powershell -Command "Invoke-WebRequest -Uri 'https://dist.nuget.org/win-x86-commandline/latest/nuget.exe' -OutFile 'nuget.exe'"
)

"%SCRIPT_DIR%nuget.exe" restore MFPControlCenter.sln -NonInteractive
if %errorLevel% neq 0 (
    echo [!] Ошибка восстановления пакетов!
    pause
    exit /b 1
)
echo [OK] Пакеты восстановлены

echo.
echo [4/6] Сборка проекта...
"%MSBUILD%" MFPControlCenter.sln /p:Configuration=Release /p:Platform="Any CPU" /t:Build /v:minimal
if %errorLevel% neq 0 (
    echo [!] Ошибка сборки проекта!
    echo [!] Проверьте, что установлены все зависимости.
    pause
    exit /b 1
)
echo [OK] Проект собран

echo.
echo [5/6] Копирование файлов в %INSTALL_DIR%...

:: Создание папки установки
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

:: Копирование файлов
xcopy /E /Y /Q "%SCRIPT_DIR%MFPControlCenter\bin\Release\*" "%INSTALL_DIR%\"
if %errorLevel% neq 0 (
    echo [!] Ошибка копирования файлов!
    pause
    exit /b 1
)
echo [OK] Файлы скопированы

echo.
echo [6/6] Создание ярлыков...

:: Создание ярлыка на рабочем столе
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%DESKTOP%\MFP Control Center.lnk'); $s.TargetPath = '%INSTALL_DIR%\MFPControlCenter.exe'; $s.IconLocation = '%INSTALL_DIR%\MFPControlCenter.exe,0'; $s.Description = 'Центр управления МФУ HP LaserJet M1536dnf'; $s.Save()"

:: Создание ярлыка в меню Пуск
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%START_MENU%\MFP Control Center.lnk'); $s.TargetPath = '%INSTALL_DIR%\MFPControlCenter.exe'; $s.IconLocation = '%INSTALL_DIR%\MFPControlCenter.exe,0'; $s.Description = 'Центр управления МФУ HP LaserJet M1536dnf'; $s.Save()"

echo [OK] Ярлыки созданы

:: Создание файла деинсталляции
echo @echo off > "%INSTALL_DIR%\uninstall.bat"
echo chcp 65001 ^>nul >> "%INSTALL_DIR%\uninstall.bat"
echo echo Удаление MFP Control Center... >> "%INSTALL_DIR%\uninstall.bat"
echo rmdir /S /Q "%INSTALL_DIR%" >> "%INSTALL_DIR%\uninstall.bat"
echo del "%DESKTOP%\MFP Control Center.lnk" 2^>nul >> "%INSTALL_DIR%\uninstall.bat"
echo del "%START_MENU%\MFP Control Center.lnk" 2^>nul >> "%INSTALL_DIR%\uninstall.bat"
echo echo Удаление завершено! >> "%INSTALL_DIR%\uninstall.bat"
echo pause >> "%INSTALL_DIR%\uninstall.bat"

echo.
echo ═══════════════════════════════════════════════════════════════
echo         УСТАНОВКА ЗАВЕРШЕНА!
echo ═══════════════════════════════════════════════════════════════
echo.
echo   Программа установлена в: %INSTALL_DIR%
echo   Ярлык создан на рабочем столе
echo.
echo   Для удаления запустите: %INSTALL_DIR%\uninstall.bat
echo.
echo ═══════════════════════════════════════════════════════════════

set /p LAUNCH="Запустить программу сейчас? (Y/N): "
if /i "%LAUNCH%"=="Y" (
    start "" "%INSTALL_DIR%\MFPControlCenter.exe"
)

pause
