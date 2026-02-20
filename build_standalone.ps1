# build_standalone.ps1
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Building DQA Dashboard Standalone EXE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if required files exist
if (-not (Test-Path "run_standalone.py")) {
    Write-Host "ERROR: run_standalone.py not found!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

if (-not (Test-Path "standalone.spec")) {
    Write-Host "ERROR: standalone.spec not found!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Clean previous builds
Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
if (Test-Path "build") { Remove-Item -Path "build" -Recurse -Force }
if (Test-Path "dist") { Remove-Item -Path "dist" -Recurse -Force }
if (Test-Path "__pycache__") { Remove-Item -Path "__pycache__" -Recurse -Force }

# Install/upgrade PyInstaller
Write-Host "Installing PyInstaller..." -ForegroundColor Yellow
python -m pip install --upgrade pyinstaller --quiet

# Install required packages
Write-Host "Installing dependencies..." -ForegroundColor Yellow
python -m pip install -r requirements_standalone.txt --quiet

# Build the executable
Write-Host "Building executable..." -ForegroundColor Yellow
Write-Host ""
pyinstaller standalone.spec --clean --noconfirm

# Check if build was successful
if (Test-Path "dist\DQA_Dashboard\DQA_Dashboard.exe") {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "BUILD SUCCESSFUL!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Executable created at: dist\DQA_Dashboard\DQA_Dashboard.exe" -ForegroundColor Green
    
    $size = (Get-ChildItem "dist\DQA_Dashboard" -Recurse | Measure-Object -Property Length -Sum).Sum
    Write-Host "Folder size: $([math]::Round($size/1MB, 2)) MB" -ForegroundColor Green
    Write-Host ""
    Write-Host "To distribute:" -ForegroundColor Yellow
    Write-Host "1. Copy the entire 'dist\DQA_Dashboard' folder"
    Write-Host "2. Users run 'DQA_Dashboard.exe'"
    Write-Host "3. No installation needed!"
    Write-Host ""
    
    # Create README
    Write-Host "Creating README.txt..." -ForegroundColor Yellow
    @"
DQA Dashboard Standalone Application
===============================

How to use:
1. Double-click DQA_Dashboard.exe
2. Wait for browser to open automatically
3. View facility performance data
4. Close application by closing the window

Note: The application runs on your local computer
and does not require internet connection.

Features:
- View facility performance scores
- See detailed results and rankings
- Export data to Excel/JSON
- No installation required
"@ | Out-File -FilePath "dist\DQA_Dashboard\README.txt" -Encoding UTF8
    
    $test = Read-Host "Would you like to test the application now? (y/n)"
    if ($test -eq "y") {
        Write-Host "Starting application..." -ForegroundColor Yellow
        Start-Process "dist\DQA_Dashboard\DQA_Dashboard.exe"
        Start-Sleep -Seconds 3
        Write-Host "Application started! Check your browser." -ForegroundColor Green
    }
} else {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "BUILD FAILED!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Check the error messages above." -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to exit"