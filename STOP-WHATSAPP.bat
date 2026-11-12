@echo off
chcp 65001 >nul
title جمعية إرث وحضارة — إيقاف خدمة WhatsApp
cd /d "%~dp0"

echo.
echo ============================================================
echo   جمعية إرث وحضارة بالقريات - إيقاف خدمات WhatsApp بأمان
echo ============================================================
echo.

echo [1/3] إيقاف نفق Cloudflare Tunnel...
taskkill /f /im cloudflared.exe >nul 2>&1

echo [2/3] إيقاف خادم Worker المحلي...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :8787 ^| findstr LISTENING') do (
    taskkill /f /pid %%a >nul 2>&1
)

echo [3/3] إيقاف حاوية OpenWA (اختياري)...
docker stop openwa-api >nul 2>&1

if exist "%~dp0tools\tunnel-url.txt" del /f /q "%~dp0tools\tunnel-url.txt" >nul 2>&1

echo.
echo ============================================================
echo   ✓ تم إيقاف جميع الخدمات والنفق السحابي بنجاح وأمان.
echo ============================================================
echo.

pause
