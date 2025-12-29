@echo off
:: Простой запускатор установщика
:: Для отца - просто двойной клик!

cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File "install.ps1"
