@echo off
chcp 65001 >nul 2>&1
echo ========================================
echo   InvoiceTool Windows Build Script
echo ========================================
echo.

:: 1. Check Python
echo [1/5] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo   ERROR: Python not found. Please install Python 3.9+
    echo   Download: https://www.python.org/downloads/
    pause
    exit /b 1
)
for /f "delims=" %%v in ('python --version 2^>^&1') do echo   %%v
echo.

:: 2. Install dependencies
echo [2/5] Installing dependencies...
python -m pip install --upgrade pip -q
if errorlevel 1 (
    echo   ERROR: pip upgrade failed
    pause
    exit /b 1
)
python -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo   ERROR: requirements install failed
    pause
    exit /b 1
)
echo   Dependencies installed
echo.

:: 3. Download 7z extra package
echo [3/5] Downloading 7z extra...
if not exist "invoice_tool\7z" mkdir "invoice_tool\7z"
powershell -Command "Invoke-WebRequest -Uri 'https://7-zip.org/a/7z2301-extra.7z' -OutFile 'invoice_tool\7z\7z-extra.7z'"
if errorlevel 1 (
    echo   ERROR: Download failed
    pause
    exit /b 1
)
echo   Download complete
echo.

:: 4. Extract 7z.exe from extra package
echo [4/5] Extracting 7z...
where 7z >nul 2>&1
if not errorlevel 1 (
    echo   Using system 7z
    7z x "invoice_tool\7z\7z-extra.7z" -oinvoice_tool\7z -y
) else (
    echo   Using Python to extract...
    python -m pip install py7zr -q
    python -c "import py7zr; py7zr.SevenZipFile('invoice_tool\\7z\\7z-extra.7z', mode='r').extractall('invoice_tool\\7z')"
)
del "invoice_tool\7z\7z-extra.7z"
if not exist "invoice_tool\7z\7z.exe" (
    echo   ERROR: 7z.exe not found
    dir invoice_tool\7z\
    pause
    exit /b 1
)
echo   7z.exe ready
echo.

:: 5. PyInstaller build
echo [5/5] Building EXE...
if not exist "dist" mkdir "dist"
pyinstaller invoice_extractor_windows.spec --noconfirm --clean
if errorlevel 1 (
    echo   ERROR: Build failed
    pause
    exit /b 1
)
echo.
echo ========================================
echo   Build SUCCESS!
echo   EXE: dist\InvoiceTool.exe
echo ========================================
pause
