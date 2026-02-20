@echo off
echo ========================================
echo Building DQA Dashboard Standalone EXE
echo ========================================
echo.

REM Check if required files exist
if not exist "run_standalone.py" (
    echo ERROR: run_standalone.py not found!
    pause
    exit /b 1
)

if not exist "standalone.spec" (
    echo ERROR: standalone.spec not found!
    pause
    exit /b 1
)

REM Clean previous builds
echo Cleaning previous builds...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "__pycache__" rmdir /s /q __pycache__

REM Install/upgrade PyInstaller
echo Installing PyInstaller...
python -m pip install --upgrade pyinstaller --quiet

REM Install required packages
echo Installing dependencies...
python -m pip install -r requirements_standalone.txt --quiet

REM Build the executable
echo Building executable...
echo.
pyinstaller standalone.spec --clean --noconfirm

REM Check if build was successful
if exist "dist\DQA_Dashboard\DQA_Dashboard.exe" (
    echo.
    echo ========================================
    echo BUILD SUCCESSFUL!
    echo ========================================
    echo.
    echo Executable created at: dist\DQA_Dashboard\DQA_Dashboard.exe
    echo.
    echo Folder size: 
    for /f "tokens=3" %%a in ('dir /s "dist\DQA_Dashboard" ^| find "File(s)"') do echo %%a bytes
    echo.
    echo To distribute:
    echo 1. Copy the entire "dist\DQA_Dashboard" folder
    echo 2. Users run "DQA_Dashboard.exe"
    echo 3. No installation needed!
    echo.
    
    REM Create a simple README
    echo Creating README.txt...
    echo DQA Dashboard Standalone Application>dist\DQA_Dashboard\README.txt
    echo ===============================>>dist\DQA_Dashboard\README.txt
    echo.>>dist\DQA_Dashboard\README.txt
    echo How to use:>>dist\DQA_Dashboard\README.txt
    echo 1. Double-click DQA_Dashboard.exe>>dist\DQA_Dashboard\README.txt
    echo 2. Wait for browser to open automatically>>dist\DQA_Dashboard\README.txt
    echo 3. View facility performance data>>dist\DQA_Dashboard\README.txt
    echo 4. Close application by closing the window>>dist\DQA_Dashboard\README.txt
    echo.>>dist\DQA_Dashboard\README.txt
    echo Note: The application runs on your local computer>>dist\DQA_Dashboard\README.txt
    echo and does not require internet connection.>>dist\DQA_Dashboard\README.txt
    
    echo.
    echo Would you like to test the application now? (y/n)
    set /p test=
    if /i "!test!"=="y" (
        echo Starting application...
        start "" "dist\DQA_Dashboard\DQA_Dashboard.exe"
        timeout /t 5 /nobreak > nul
        echo Application started! Check your browser.
    )
) else (
    echo.
    echo ========================================
    echo BUILD FAILED!
    echo ========================================
    echo Check the error messages above.
)

echo.
echo Press any key to exit...
pause > nul