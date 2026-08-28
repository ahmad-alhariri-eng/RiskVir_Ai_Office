@echo off
setlocal EnableDelayedExpansion
echo ==============================================
echo       OfficeAI Pro - Executable Builder
echo ==============================================

set ROOT=%~dp0..

:: ── Step 1: Build the React frontend ──────────────────────────
echo.
echo [1/3] Building React Add-in (TypeScript + Vite)...
cd /d "%ROOT%\office-addin"
call npm run build

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Frontend build failed. Run 'npm install' first.
    pause & exit /b 1
)
echo [OK] Frontend built to office-addin\dist\

:: ── Step 2: Ensure PyInstaller is installed ───────────────────
echo.
echo [2/3] Verifying PyInstaller...
cd /d "%ROOT%\backend"
python -m pip install pyinstaller --quiet

:: ── Step 3: Compile executables using spec files ──────────────
echo.
echo [3/3] Compiling Backend + Launcher executables...
cd /d "%ROOT%"

echo   Building OfficeAI_Backend.exe ...
python -m PyInstaller OfficeAI_Backend.spec --noconfirm

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Backend build failed.
    pause & exit /b 1
)

echo   Building OfficeAI_Launcher.exe ...
python -m PyInstaller OfficeAI_Launcher.spec --noconfirm

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Launcher build failed.
    pause & exit /b 1
)

:: ── Done ───────────────────────────────────────────────────────
echo.
echo ==============================================
echo  Build Complete!
echo ==============================================
echo.
echo  Executables:
echo    dist\OfficeAI_Backend.exe
echo    dist\OfficeAI_Launcher.exe
echo.
echo  Deployment checklist:
echo    1. Copy dist\OfficeAI_Backend.exe  ^} to user machine
echo    2. Copy dist\OfficeAI_Launcher.exe ^}
echo    3. Copy models\*.gguf              (required — not bundled)
echo    4. Sideload office-addin\manifest.xml in Office
echo ==============================================
pause
