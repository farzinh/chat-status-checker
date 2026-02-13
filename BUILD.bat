@echo off
echo ============================================================
echo   Chat Status Monitor - Build Standalone EXE
echo ============================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Step 1: Installing dependencies...
pip install pyinstaller pyautogui opencv-python numpy pytesseract mss Pillow tzdata

echo.
echo Step 2: Building executable...
pyinstaller --name=ChatStatusMonitor --onefile --windowed --noconfirm --clean chat_monitor_gui.py

echo.
echo ============================================================
if exist "dist\ChatStatusMonitor.exe" (
    echo   BUILD SUCCESSFUL!
    echo ============================================================
    echo.
    echo Your executable is ready at:
    echo   dist\ChatStatusMonitor.exe
    echo.
    echo You can copy this file anywhere and run it!
    echo Make sure Tesseract OCR is installed on the target computer.
) else (
    echo   BUILD FAILED
    echo ============================================================
    echo Please check the error messages above.
)

echo.
pause