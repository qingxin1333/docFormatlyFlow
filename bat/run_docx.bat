@echo off
echo ========================================
echo Document Formatter Tool - Generate Formatted DOCX
echo ========================================
echo.

REM Check if shared virtual environment exists - use this first
if exist "D:\pythonVenv\3.11\docFormatlyFlow\Scripts\activate.bat" (
    echo Using shared virtual environment: D:\pythonVenv\3.11\docFormatlyFlow
    call "D:\pythonVenv\3.11\docFormatlyFlow\Scripts\activate.bat"
) else (
    REM Check if conda environment exists
    call conda info --envs | findstr /C:"docformatter" >nul 2>&1
    if %errorlevel%==0 (
        echo Using conda environment: docformatter
        call conda activate docformatter
    ) else (
        echo ERROR: No virtual environment found
        echo Please run setup_env.bat first
        pause
        exit /b 1
    )
)

echo.
echo Starting formatted document formatter to generate DOCX files...
echo.

REM Run formatted docx formatter
python docx_formatter.py

echo.
echo Program execution completed.
