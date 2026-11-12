@echo off
chcp 65001 >nul
title جمعية إرث وحضارة — تشغيل خدمة WhatsApp
cd /d "%~dp0"

echo.
echo ============================================================
echo   جمعية إرث وحضارة بالقريات - تشغيل خدمة WhatsApp و OpenWA
echo ============================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\start-whatsapp.ps1"

pause
