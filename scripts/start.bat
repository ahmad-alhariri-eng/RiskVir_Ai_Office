@echo off
:: OfficeAI Pro — Quick Start Script
:: Runs backend in one window and frontend dev server in another.

set ROOT=%~dp0

:: Load .env if it exists
if exist "%ROOT%.env" (
    for /f "usebackq tokens=1,* delims==" %%A in ("%ROOT%.env") do (
        if not "%%A"=="" if not "%%A:~0,1%"=="#" set %%A=%%B
    )
)

set HOST=%OFFICEAI_HOST%
if "%HOST%"=="" set HOST=127.0.0.1

set PORT=%OFFICEAI_PORT%
if "%PORT%"=="" set PORT=8000

echo ==============================================
echo  OfficeAI Pro — Starting Services
echo ==============================================
echo  Backend : http(s)://%HOST%:%PORT%
echo  Frontend: http://localhost:5173 (Vite dev)
echo ==============================================
echo.

:: Start backend in a new window
start "OfficeAI Backend" cmd /k "cd /d "%ROOT%backend" && python main.py"

:: Wait a moment for the backend to initialize
timeout /t 3 /nobreak >nul

:: Start frontend dev server in a new window
start "OfficeAI Frontend" cmd /k "cd /d "%ROOT%office-addin" && npm run dev"

echo Both services are starting in separate windows.
echo Close this window when done.
pause
