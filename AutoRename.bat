@echo off
title AutoRename System Installer
cls

echo ==========================================
echo  Welcome to AutoRename Safe Installer
echo ==========================================
echo.

:: 1. Check Python
echo Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is NOT installed! 
    echo Downloading and installing Python from Microsoft Store...
    winget install --id Python.Python.3.12 --source winget --accept-package-agreements --accept-source-agreements
    if %errorlevel% neq 0 (
        echo Failed to install Python automatically. Please install Python manually from python.org
        pause
        exit /b
    )
    echo Python installed successfully! Please restart this Setup.bat file.
    pause
    exit /b
) else (
    echo Python is already installed.
)

echo.
echo Installing required AI components (PyMuPDF, EasyOCR, OpenCV, NumPy)...
echo (This may take a few minutes depending on your internet speed...)
python -m pip install --upgrade pip
python -m pip install pymupdf easyocr numpy opencv-python

echo.
echo Deploying script to a safe location...
SET "SAFE_DIR=%ProgramData%\AutoRename"
if not exist "%SAFE_DIR%" mkdir "%SAFE_DIR%"

SET "CURRENT_DIR=%~dp0"
if exist "%CURRENT_DIR%rename_docs.py" (
    copy /Y "%CURRENT_DIR%rename_docs.py" "%SAFE_DIR%\rename_docs.py" >nul
    echo Script deployed successfully to %SAFE_DIR%
) else (
    echo Error: rename_docs.py not found in the current folder!
    pause
    exit /b
)

echo.
echo Configuring SendTo Shortcut...
SET "PY_SCRIPT_PATH=%SAFE_DIR%\rename_docs.py"
SET "SENDTO_DIR=%APPDATA%\Microsoft\Windows\SendTo"
SET "BAT_SHORTCUT=%SENDTO_DIR%\AutoRename.bat"

(
echo @echo off
echo python "%PY_SCRIPT_PATH%" "%%~1"
) > "%BAT_SHORTCUT%"

echo.
echo ===================================================
echo SETUP COMPLETED SUCCESSFULLY!
echo settings are deployed safely in C:\ProgramData\AutoRename
echo You can now delete the downloaded zip/folder if you want.
echo ===================================================
echo.
pause
