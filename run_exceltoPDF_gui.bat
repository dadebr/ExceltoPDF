@echo off
:: Windows Batch Launcher for ExceltoPDF GUI
:: Tries to activate local venv if present and launches exceltopdf-gui
:: Falls back to python -m exceltopdf.gui if exceltopdf-gui command fails

echo Starting ExceltoPDF GUI...
echo.

:: Check if local venv exists and activate it
if exist "%~dp0venv\Scripts\activate.bat" (
    echo Activating local virtual environment...
    call "%~dp0venv\Scripts\activate.bat"
) else if exist "%~dp0.venv\Scripts\activate.bat" (
    echo Activating local virtual environment (.venv)...
    call "%~dp0.venv\Scripts\activate.bat"
) else (
    echo No local virtual environment found, using system Python...
)

echo.

:: Try to launch exceltopdf-gui command first
echo Attempting to launch exceltopdf-gui...
exceltopdf-gui 2>nul
if %ERRORLEVEL% neq 0 (
    echo exceltopdf-gui command not found, trying fallback...
    echo Launching via python -m exceltopdf.gui
    python -m exceltopdf.gui
    if %ERRORLEVEL% neq 0 (
        echo.
        echo Error: Could not launch ExceltoPDF GUI.
        echo Please ensure ExceltoPDF is properly installed.
        echo.
        echo Try one of these commands manually:
        echo   exceltopdf-gui
        echo   python -m exceltopdf.gui
        echo.
        pause
        exit /b 1
    )
)

:: If we get here, the GUI should have started successfully
echo GUI launched successfully!
