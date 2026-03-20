@echo off
chcp 65001 >nul 2>&1
title PHI Sequenz-Manager

set "ROOT=%~dp0"
set "ROOT=%ROOT:~0,-1%"
set "EXE=%ROOT%\hybrid_approach.exe"
set "PORT=8765"
set "URL=http://127.0.0.1:%PORT%"

if not exist "%EXE%" (
    echo.
    echo  FEHLER: Backend-Binary nicht gefunden:
    echo    %EXE%
    echo.
    echo  Bitte laden Sie die aktuelle Version herunter
    echo  oder bauen Sie das Projekt neu.
    echo.
    pause
    exit /b 1
)

echo.
echo  ============================================
echo   PHI Sequenz-Manager  ^|  Hybrid Runtime
echo  ============================================
echo.
echo  Starte Server auf %URL% ...
echo.

start "" "%URL%"
"%EXE%" serve --root "%ROOT%" --port %PORT%

echo.
echo  Server wurde beendet.
pause
