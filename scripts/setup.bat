@echo off
setlocal EnableDelayedExpansion
echo ==============================================
echo       OfficeAI Pro - Initial Setup
echo ==============================================

set ROOT=%~dp0..

:: ── Step 1: Backend dependencies ──────────────────────────────
echo.
echo [1/4] Installing Backend Python Dependencies...
cd /d "%ROOT%\backend"
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pystray Pillow pyinstaller

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to install Python dependencies.
    pause & exit /b 1
)

:: ── Step 2: Download model ─────────────────────────────────────
echo.
echo [2/4] Checking for AI Model...
set MODEL_DIR=%ROOT%\models
set MODEL_FILE=%MODEL_DIR%\microsoft_Phi-4-mini-instruct-Q4_K_M.gguf

if exist "%MODEL_FILE%" (
    echo [SKIP] Model already exists at:
    echo        %MODEL_FILE%
) else (
    echo Downloading Phi-4-mini model (approx. 2.5 GB)...
    echo This may take several minutes depending on your connection.
    pip install -U "huggingface_hub[cli]"
    python -c "from huggingface_hub import hf_hub_download; hf_hub_download(repo_id='bartowski/microsoft_Phi-4-mini-instruct-GGUF', filename='microsoft_Phi-4-mini-instruct-Q4_K_M.gguf', local_dir=r'%MODEL_DIR%')"
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Model download failed. Check your internet connection.
        pause & exit /b 1
    )
)

:: ── Step 3: Create .env if missing ────────────────────────────
echo.
echo [3/4] Configuring Environment...
if not exist "%ROOT%\.env" (
    copy "%ROOT%\.env.example" "%ROOT%\.env" >nul 2>&1
    echo [OK] Created .env from .env.example
) else (
    echo [SKIP] .env already exists
)

:: ── Step 4: Frontend dependencies ─────────────────────────────
echo.
echo [4/4] Installing Frontend (Node) Dependencies...
cd /d "%ROOT%\office-addin"
call npm install

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] npm install failed. Make sure Node.js is installed.
    pause & exit /b 1
)

:: ── Done ───────────────────────────────────────────────────────
echo.
echo ==============================================
echo  Setup Complete!
echo ==============================================
echo.
echo  To start the backend:
echo    python launcher\launcher.py
echo    OR: cd backend ^& python main.py
echo.
echo  To start the Add-in dev server:
echo    cd office-addin ^& npm run dev
echo.
echo  Model location:
echo    %MODEL_FILE%
echo.
echo  Optional: set OFFICEAI_MODEL_PATH in .env
echo    to use a different model.
echo ==============================================
pause
