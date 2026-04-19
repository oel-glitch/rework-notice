@echo off
setlocal enabledelayedexpansion

echo =================================================================
echo   PDF Parser OATI - Build EXE (NO INSTALLATION REQUIRED!)
echo =================================================================
echo.

REM Check if Python is already installed
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] Python found in system
    python --version
    set "PYTHON_CMD=python"
    goto :check_pip
)

echo [INFO] Python not found - downloading portable version...
echo.

REM Download portable Python if not exists
if exist "python_portable\python.exe" (
    echo [OK] Portable Python already downloaded
    set "PYTHON_CMD=python_portable\python.exe"
    goto :check_pip
)

echo [1/6] Downloading portable Python 3.11 (~13 MB)...
powershell -ExecutionPolicy Bypass -Command "& {$ProgressPreference='SilentlyContinue';[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;(New-Object System.Net.WebClient).DownloadFile('https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip','python_embed.zip')}"

if not exist "python_embed.zip" (
    echo [ERROR] Failed to download Python
    echo Please check internet connection or download manually:
    echo https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip
    pause
    exit /b 1
)

echo [OK] Downloaded
echo.

echo Extracting Python...
powershell -ExecutionPolicy Bypass -Command "Expand-Archive -Path 'python_embed.zip' -DestinationPath 'python_portable' -Force"
del python_embed.zip

if not exist "python_portable\python.exe" (
    echo [ERROR] Failed to extract Python
    pause
    exit /b 1
)

echo [OK] Python extracted
echo.

echo Setting up pip...
powershell -ExecutionPolicy Bypass -Command "(Get-Content 'python_portable\python311._pth') -replace '#import site','import site' | Set-Content 'python_portable\python311._pth'"

powershell -ExecutionPolicy Bypass -Command "& {$ProgressPreference='SilentlyContinue';[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile 'get-pip.py'}"
python_portable\python.exe get-pip.py >nul 2>&1
del get-pip.py

echo [OK] pip ready
echo.

set "PYTHON_CMD=python_portable\python.exe"

:check_pip
echo [2/6] Updating pip...
%PYTHON_CMD% -m pip install --quiet --upgrade pip

echo [OK] pip updated
echo.

echo [3/6] Installing project dependencies from requirements.txt...
%PYTHON_CMD% -m pip install --quiet -r requirements.txt

if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)
echo [OK] Dependencies installed
echo.

echo [4/6] Installing PyInstaller...
%PYTHON_CMD% -m pip install --quiet pyinstaller

if %errorlevel% neq 0 (
    echo [ERROR] Failed to install PyInstaller
    pause
    exit /b 1
)
echo [OK] PyInstaller installed
echo.

echo [5/6] Building executable...
echo This will take 2-3 minutes, please wait...
echo.

%PYTHON_CMD% build.py

if %errorlevel% neq 0 (
    echo [ERROR] Build failed
    pause
    exit /b 1
)

if not exist "dist\PDF_Parser_OATI.exe" (
    echo [ERROR] EXE file not created
    pause
    exit /b 1
)

echo [OK] Build completed
echo.

echo [6/6] Creating final package...

REM Create distribution folder
if exist "PDF_Parser_OATI_Final" rd /s /q "PDF_Parser_OATI_Final"
mkdir "PDF_Parser_OATI_Final"
mkdir "PDF_Parser_OATI_Final\templates"
mkdir "PDF_Parser_OATI_Final\database"

REM Copy files
copy /Y "dist\PDF_Parser_OATI.exe" "PDF_Parser_OATI_Final\" >nul
xcopy /E /I /Y "templates\*.docx" "PDF_Parser_OATI_Final\templates\" >nul
if exist "INSTRUCTION_RU.txt" copy /Y "INSTRUCTION_RU.txt" "PDF_Parser_OATI_Final\" >nul

REM Create database readme
echo Database will be created automatically on first run. > "PDF_Parser_OATI_Final\database\README.txt"

REM Create ZIP archive
powershell -ExecutionPolicy Bypass -Command "Compress-Archive -Path 'PDF_Parser_OATI_Final\*' -DestinationPath 'PDF_Parser_OATI_Standalone.zip' -Force"

if not exist "PDF_Parser_OATI_Standalone.zip" (
    echo [ERROR] Failed to create ZIP
    pause
    exit /b 1
)

echo [OK] Package created
echo.

echo Cleaning up temporary files...
rd /s /q "build" >nul 2>&1
rd /s /q "dist" >nul 2>&1
rd /s /q "PDF_Parser_OATI_Final" >nul 2>&1
rd /s /q "python_portable" >nul 2>&1
del /f /q "*.spec~" >nul 2>&1

echo [OK] Cleanup completed
echo.

echo =================================================================
echo                   BUILD COMPLETED SUCCESSFULLY!
echo =================================================================
echo.
echo   Package ready: PDF_Parser_OATI_Standalone.zip
echo.

REM Get file size
for %%A in (PDF_Parser_OATI_Standalone.zip) do (
    set size=%%~zA
    set /a sizeMB=!size! / 1048576
    echo   Size: !sizeMB! MB
)

echo.
echo   Now you can:
echo   1. Extract PDF_Parser_OATI_Standalone.zip
echo   2. Run PDF_Parser_OATI.exe
echo   3. Use without Python installation!
echo.
echo   [INFO] Portable Python was removed after build
echo.
pause
