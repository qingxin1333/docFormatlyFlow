@echo off
echo ========================================
echo Document Formatter Tool
echo ========================================
echo.
echo 请选择输出格式：
echo 1. 生成格式化的 DOCX 文件
echo 2. 生成 Markdown 文件
echo 3. 同时生成两种格式
echo 4. 退出
echo.

set /p choice="请输入选择 (1-4): "

if "%choice%"=="1" goto run_docx
if "%choice%"=="2" goto run_md
if "%choice%"=="3" goto run_both
if "%choice%"=="4" goto end
echo 无效选择，请重新运行
pause
goto end

:run_docx
echo.
echo 正在生成格式化的 DOCX 文件...
goto setup_env

:run_md
echo.
echo 正在生成 Markdown 文件...
goto setup_env

:run_both
echo.
echo 正在生成两种格式的文件...
goto setup_env

:setup_env
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
echo 开始处理文件...

if "%choice%"=="1" (
    python docx_formatter.py
) else if "%choice%"=="2" (
    python doc_formatter_to_md.py
) else if "%choice%"=="3" (
    echo 正在生成 DOCX 格式...
    python docx_formatter.py
    echo.
    echo 正在生成 Markdown 格式...
    python doc_formatter_to_md.py
)

echo.
echo 程序执行完成。
goto end

:end
pause
