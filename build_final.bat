@echo off
echo ========================================
echo Building DQA Dashboard Standalone EXE
echo ========================================
echo.

REM Clean previous builds
echo Cleaning previous builds...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "__pycache__" rmdir /s /q __pycache__
if exist "*.spec" del /q *.spec

REM Install required packages
echo Installing PyInstaller...
pip install pyinstaller --quiet --upgrade

echo Installing dependencies...
pip install -r requirements_standalone.txt --quiet

REM Verify data files exist
echo.
echo Verifying data files...
if not exist "data\results.json" (
    echo ERROR: data\results.json not found!
    echo Make sure your facility data exists.
    pause
    exit /b 1
)

echo Found data\results.json
dir "data\results.json"

REM Build command
echo.
echo Building executable...
echo This may take a few minutes...
echo.

pyinstaller ^
  --name="DQA_Dashboard" ^
  --onefile ^
  --windowed ^
  --add-data="app;app" ^
  --add-data="data;data" ^
  --add-data="config;config" ^
  --add-data="assets;assets" ^
  --add-data="icon.ico;." ^
  --hidden-import="flask" ^
  --hidden-import="werkzeug" ^
  --hidden-import="jinja2" ^
  --hidden-import="openpyxl" ^
  --hidden-import="pandas" ^
  --hidden-import="numpy" ^
  --hidden-import="app.routes" ^
  --hidden-import="app.analysis" ^
  --hidden-import="app.storage" ^
  --hidden-import="app.__init__" ^
  --hidden-import="flask.cli" ^
  --hidden-import="jinja2.ext" ^
  --hidden-import="datetime" ^
  --hidden-import="uuid" ^
  --hidden-import="json" ^
  --hidden-import="os" ^
  --hidden-import="sys" ^
  --hidden-import="logging" ^
  --hidden-import="shutil" ^
  --hidden-import="socket" ^
  --hidden-import="webbrowser" ^
  --hidden-import="threading" ^
  --clean ^
  --icon="icon.ico" ^
  run_standalone.py

REM Check if successful
if exist "dist\DQA_Dashboard.exe" (
    echo.
    echo ========================================
    echo BUILD SUCCESSFUL!
    echo ========================================
    echo.
    echo Executable: dist\DQA_Dashboard.exe
    echo Size: 
    for %%F in ("dist\DQA_Dashboard.exe") do echo   %%~zF bytes ^(%%~zF/1048576 MB^)
    echo.
    echo To distribute:
    echo   1. Copy dist\DQA_Dashboard.exe to any folder
    echo   2. Run it - it will create DQA_Data folder automatically
    echo   3. No installation needed!
    echo.
    
    REM Create README
    echo Creating README.txt...
    (
    echo DQA Dashboard - Standalone Application
    echo ======================================
    echo.
    echo How to Use:
    echo 1. Run DQA_Dashboard.exe
    echo 2. Wait for browser to open automatically
    echo 3. View facility performance data
    echo 4. Close by closing the application window
    echo.
    echo Data Storage:
    echo - Application creates "DQA_Data" folder in same directory
    echo - Contains results.json with all facility data
    echo - You can backup or restore data by copying this folder
    echo.
    echo Features:
    echo - View-only mode (analysis features disabled)
    echo - All data included and preserved
    echo - No installation or admin rights required
    echo - Runs on your local computer
    echo.
    echo Troubleshooting:
    echo - If browser doesn't open, visit: http://127.0.0.1:8000
    echo - If port is busy, restart the application
    echo - Check DQA_Data folder exists and contains results.json
    ) > "dist\README.txt"
    
    echo.
    echo Would you like to test now? (y/n)
    set /p test=
    if /i "!test!"=="y" (
        echo Testing application...
        start "" "dist\DQA_Dashboard.exe"
        timeout /t 3
        echo Application started! Check your browser.
    )
) else (
    echo.
    echo ========================================
    echo BUILD FAILED!
    echo ========================================
)

echo.
pause