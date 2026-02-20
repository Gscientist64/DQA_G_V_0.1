# build_final.ps1
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Building DQA Dashboard Standalone EXE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Clean previous builds
Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
Remove-Item -Path "build", "dist", "__pycache__" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "*.spec" -Force -ErrorAction SilentlyContinue

# Install packages
Write-Host "Installing PyInstaller..." -ForegroundColor Yellow
pip install pyinstaller --quiet --upgrade

Write-Host "Installing dependencies..." -ForegroundColor Yellow
pip install -r requirements_standalone.txt --quiet

# Verify data
Write-Host "`nVerifying data files..." -ForegroundColor Yellow
if (-not (Test-Path "data\results.json")) {
    Write-Host "ERROR: data\results.json not found!" -ForegroundColor Red
    Write-Host "Make sure your facility data exists in data\ folder" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

$fileInfo = Get-Item "data\results.json"
Write-Host "Found data\results.json ($([math]::Round($fileInfo.Length/1KB, 2)) KB)" -ForegroundColor Green

# Build command
Write-Host "`nBuilding executable..." -ForegroundColor Yellow
Write-Host "This may take a few minutes..." -ForegroundColor Yellow

$buildArgs = @(
    "--name=DQA_Dashboard",
    "--onefile",
    "--windowed",
    "--add-data=app;app",
    "--add-data=data;data",
    "--add-data=config;config",
    "--add-data=assets;assets",
    "--add-data=icon.ico;.",
    "--hidden-import=flask",
    "--hidden-import=werkzeug",
    "--hidden-import=jinja2",
    "--hidden-import=openpyxl",
    "--hidden-import=pandas",
    "--hidden-import=numpy",
    "--hidden-import=app.routes",
    "--hidden-import=app.analysis",
    "--hidden-import=app.storage",
    "--hidden-import=app.__init__",
    "--hidden-import=flask.cli",
    "--hidden-import=jinja2.ext",
    "--hidden-import=datetime",
    "--hidden-import=uuid",
    "--hidden-import=json",
    "--hidden-import=os",
    "--hidden-import=sys",
    "--hidden-import=logging",
    "--hidden-import=shutil",
    "--hidden-import=socket",
    "--hidden-import=webbrowser",
    "--hidden-import=threading",
    "--clean",
    "--icon=icon.ico",
    "run_standalone.py"
)

pyinstaller @buildArgs

# Check result
if (Test-Path "dist\DQA_Dashboard.exe") {
    Write-Host "`n" + "="*40 -ForegroundColor Green
    Write-Host "BUILD SUCCESSFUL!" -ForegroundColor Green
    Write-Host "="*40 -ForegroundColor Green
    Write-Host ""
    
    $exePath = "dist\DQA_Dashboard.exe"
    $size = (Get-Item $exePath).Length
    Write-Host "Executable: $exePath" -ForegroundColor Green
    Write-Host "Size: $([math]::Round($size/1MB, 2)) MB" -ForegroundColor Green
    
    # Create README
    @"
DQA Dashboard - Standalone Application
======================================

How to Use:
1. Run DQA_Dashboard.exe
2. Wait for browser to open automatically
3. View facility performance data
4. Close by closing the application window

Data Storage:
- Application creates "DQA_Data" folder in same directory
- Contains results.json with all facility data
- You can backup or restore data by copying this folder

Features:
- View-only mode (analysis features disabled)
- All data included and preserved
- No installation or admin rights required
- Runs on your local computer

Troubleshooting:
- If browser doesn't open, visit: http://127.0.0.1:8000
- If port is busy, restart the application
- Check DQA_Data folder exists and contains results.json
"@ | Out-File -FilePath "dist\README.txt" -Encoding UTF8
    
    Write-Host "`nDistribution ready!" -ForegroundColor Green
    Write-Host "To distribute: Copy dist\DQA_Dashboard.exe to any folder" -ForegroundColor Yellow
    Write-Host "It will create DQA_Data folder automatically on first run" -ForegroundColor Yellow
    
    $test = Read-Host "`nTest the application now? (y/n)"
    if ($test -eq 'y') {
        Write-Host "Starting application..." -ForegroundColor Yellow
        Start-Process $exePath
        Start-Sleep -Seconds 2
        Write-Host "Application started! Check your browser." -ForegroundColor Green
    }
} else {
    Write-Host "`n" + "="*40 -ForegroundColor Red
    Write-Host "BUILD FAILED!" -ForegroundColor Red
    Write-Host "="*40 -ForegroundColor Red
}

Write-Host "`nPress Enter to exit"
Read-Host