@echo off
echo Starting ExceltoPDF GUI...
echo.

echo Checking if ExceltoPDF is installed...
exceltopdf-gui --help >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo ExceltoPDF found, launching GUI...
    exceltopdf-gui
) else (
    echo ExceltoPDF command not found, trying fallback...
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

echo.
echo GUI closed.
