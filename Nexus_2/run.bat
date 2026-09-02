@echo off
color 0A
title Nexus Enterprise System Launcher
cls
echo.
echo   =========================================================
echo      N E X U S   E N T E R P R I S E   S Y S T E M
echo   =========================================================
echo.

:: 1. LANCEMENT APP 5 (Le dossier est 'monthly_reports_service' juste à côté du .bat)
echo   [1/2] Launching Monthly Reports Microservice (Port 5001)...
start "App 5 (Port 5001)" cmd /k "cd /d "%~dp0monthly_reports_service" && python app.py"

:: 2. LANCEMENT HUB PRINCIPAL (Le app.py est dans le même dossier que le .bat)
echo   [2/2] Launching Main Nexus Hub (Port 5000)...
start "Nexus Hub (Port 5000)" cmd /k "cd /d "%~dp0." && python app.py"

echo.
echo   Waiting for servers to bind to ports...
timeout /t 4 /nobreak >nul

echo   Launching Web Interface...
start "" "http://localhost:5000"

echo.
echo   =========================================================
echo      SYSTEM ONLINE! DO NOT CLOSE THE BLACK WINDOWS!
echo   =========================================================
echo.
echo   Appuyez sur une touche pour fermer ce message...
pause >nul