@echo off
echo ========================================
echo Document Formatter Tool - Environment Setup
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found, please install Python 3.7+
    pause
    exit /b 1
)

echo Python detected
python --version

REM Check if conda is available
call conda --version >nul 2>&1
if errorlevel 1 (
    echo WARNING: conda not found, using system Python environment
    set USE_CONDA=0
) else (
    echo conda detected
    set USE_CONDA=1
)

echo.
echo Starting to create Python environment...
echo.

if %USE_CONDA%==1 (
    echo Using conda to create environment...
    echo Creating virtual environment: docformatter
    conda create -n docformatter python=3.11 -y
    
    echo Activating environment...
    call conda activate docformatter
    
    echo Upgrading pip...
    python -m pip install --upgrade pip
    
) else (
    echo Using venv to create environment...
    echo Creating virtual environment: venv
    python -m venv venv
    
    echo Activating environment...
    call venv\Scripts\activate.bat
    
    echo Upgrading pip...
    python -m pip install --upgrade pip
)

echo.
echo Installing project dependencies...
echo.

REM Install dependencies
pip install -r requirements.txt

echo.
echo Checking Ollama connection...
echo.

python -c "import ollama; print('Ollama library installed successfully')" 2>nul
if errorlevel 1 (
    echo WARNING: Ollama library installation failed, please check network connection
) else (
    echo Ollama library installed successfully
)

echo.
echo Installation completed!
echo.
echo Usage:
echo 1. Put docx files in source directory
echo 2. Run python doc_formatter.py
echo 3. Formatted markdown files will be saved in target directory
echo.

if %USE_CONDA%==1 (
    echo Activate environment command: conda activate docformatter
) else (
    echo Activate environment command: venv\Scripts\activate.bat
)

echo.
pause
