@echo off
:: Windows Batch Launcher for ExceltoPDF CLI
:: Tries to activate local venv if present and launches exceltopdf with all arguments
:: Falls back to python -m exceltopdf.cli if exceltopdf command fails

setlocal

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

:: Try to launch exceltopdf command first, forwarding all arguments
exceltopdf %* 2>nul
if %ERRORLEVEL% neq 0 (
    echo exceltopdf command not found, trying fallback...
    python -m exceltopdf.cli %*
    if %ERRORLEVEL% neq 0 (
        echo.
        echo Error: Could not launch ExceltoPDF CLI.
        echo Please ensure ExceltoPDF is properly installed.
        echo.
        echo Try one of these commands manually:
        echo   exceltopdf [args]
        echo   python -m exceltopdf.cli [args]
        echo.
        pause
        exit /b 1
    )
)

endlocal
