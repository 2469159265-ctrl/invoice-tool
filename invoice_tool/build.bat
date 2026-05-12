@echo off
chcp 65001 >nul
echo ========================================
echo   InvoiceTool Windows 本地打包脚本
echo ========================================
echo.

:: 1. 检查 Python
echo [1/5] 检查 Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo   错误：未找到 Python，请先安装 Python 3.9+
    echo   下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)
for /f "delims=" %%v in ('python --version 2^>^&1') do echo   %%v
echo.

:: 2. 安装依赖
echo [2/5] 安装依赖...
pip install --upgrade pip -q
if errorlevel 1 (
    echo   错误：pip 安装失败
    pause
    exit /b 1
)
pip install -r requirements.txt -q
if errorlevel 1 (
    echo   错误：依赖安装失败
    pause
    exit /b 1
)
echo   依赖安装完成
echo.

:: 3. 下载并解压 7z extra 包
echo [3/5] 下载 7z extra 包...
if not exist "invoice_tool\7z" mkdir "invoice_tool\7z"
:: 优先用 PowerShell 的 Invoke-WebRequest（兼容性好）
powershell -Command "Invoke-WebRequest -Uri 'https://7-zip.org/a/7z2301-extra.7z' -OutFile 'invoice_tool\7z\7z-extra.7z'"
if errorlevel 1 (
    echo   错误：下载失败
    pause
    exit /b 1
)
echo   下载完成
echo.

:: 4. 解压 extra 包得到 7z.exe
echo [4/5] 解压 7z extra 包...
:: 检查系统 7z 是否存在
where 7z >nul 2>&1
if not errorlevel 1 (
    echo   使用系统 7z 解压
    7z x "invoice_tool\7z\7z-extra.7z" -oinvoice_tool\7z -y
) else (
    :: 没有系统 7z 用 Python py7zr 库解压
    echo   使用 Python 解压...
    pip install py7zr -q
    python -c "import py7zr; py7zr.SevenZipFile('invoice_tool\\7z\\7z-extra.7z', mode='r').extractall('invoice_tool\\7z')"
)
del "invoice_tool\7z\7z-extra.7z"
if not exist "invoice_tool\7z\7z.exe" (
    echo   错误：7z.exe 解压失败
    dir invoice_tool\7z\
    pause
    exit /b 1
)
echo   7z.exe 准备完成
echo.

:: 5. PyInstaller 打包
echo [5/5] 开始打包...
if not exist "dist" mkdir "dist"
pyinstaller invoice_extractor_windows.spec --noconfirm --clean
if errorlevel 1 (
    echo   错误：打包失败
    pause
    exit /b 1
)
echo.
echo ========================================
echo   打包完成！
echo   EXE 文件：dist\InvoiceTool.exe
echo ========================================
pause
