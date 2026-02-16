# Check if script is running with necessary permissions
$ErrorActionPreference = "Stop"

Write-Host "Running Automated Setup Script v2" -ForegroundColor Magenta

function Check-Command($cmd, $name) {
    if (Get-Command $cmd -ErrorAction SilentlyContinue) {
        Write-Host "Found $name in PATH." -ForegroundColor Green
        return $true
    }
    return $false
}

Write-Host "Checking System Requirements..." -ForegroundColor Cyan

# 1. Tesseract Check & Auto-Configuration
$tesseractFound = Check-Command "tesseract" "Tesseract-OCR"

if (-not $tesseractFound) {
    # Check common locations
    $commonPaths = @(
        "C:\Program Files\Tesseract-OCR",
        "C:\Program Files (x86)\Tesseract-OCR",
        "$env:LOCALAPPDATA\Tesseract-OCR"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path "$path\tesseract.exe") {
            Write-Host "Found Tesseract at $path. Adding to PATH for this session..." -ForegroundColor Green
            $env:Path = "$path;" + $env:Path
            $tesseractFound = $true
            break
        }
    }
}

if (-not $tesseractFound) {
    Write-Warning "Tesseract-OCR not found."
    Write-Host "Downloading Tesseract Installer..." -ForegroundColor Yellow
    
    $installerUrl = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.3/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
    $installerPath = "$PWD\tesseract_installer.exe"
    
    try {
        Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -ErrorAction Stop
        Write-Host "Installer downloaded to $installerPath" -ForegroundColor Green
        Write-Host "Launching installer... Please complete the installation." -ForegroundColor Yellow
        Write-Host "IMPORTANT: You do NOT need to add to PATH manually, just install dependencies." -ForegroundColor Yellow
        $process = Start-Process $installerPath -Wait -PassThru
        
        # Re-check after install
        foreach ($path in $commonPaths) {
            if (Test-Path "$path\tesseract.exe") {
                Write-Host "Found Tesseract at $path after install. Adding to PATH..." -ForegroundColor Green
                $env:Path = "$path;" + $env:Path
                $tesseractFound = $true
                break
            }
        }
    } catch {
        Write-Warning "Could not auto-download installer. Please download from: https://github.com/UB-Mannheim/tesseract/wiki"
    }
}

# 2. Poppler Check
if (-not (Check-Command "pdftoppm" "Poppler")) {
    Write-Warning "Poppler not found. Downloading..."
    
    $popplerUrl = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.02.0-0/Release-24.02.0-0.zip"
    $zipPath = "$PWD\poppler.zip"
    $extractPath = "$PWD\poppler_lib"
    
    try {
        if (-not (Test-Path $zipPath)) {
            Write-Host "Downloading Poppler..." -ForegroundColor Yellow
            Invoke-WebRequest -Uri $popplerUrl -OutFile $zipPath -ErrorAction Stop
        }
        
        if (-not (Test-Path $extractPath)) {
            Write-Host "Extracting Poppler..." -ForegroundColor Yellow
            Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
        }
        
        # Find the bin folder. Usually poppler_lib/poppler-xx/Library/bin
        $binPath = Get-ChildItem -Path $extractPath -Recurse -Filter "pdftoppm.exe" | Select-Object -First 1 -ExpandProperty DirectoryName
        
        if ($binPath) {
            Write-Host "Found Poppler bin at $binPath. Adding to PATH..." -ForegroundColor Green
            $env:Path = "$binPath;" + $env:Path
        } else {
            Write-Error "Could not find pdftoppm.exe in extracted archive."
        }
        
    } catch {
        Write-Warning "Failed to setup Poppler: $_"
        Write-Host "PDF Support might not work."
    }
}

# 3. Python Check
$pyCommand = "python"
if (Get-Command "py" -ErrorAction SilentlyContinue) {
    try {
        $versions = py --list
        if ($versions -match "3.10") {
            $pyCommand = "py -3.10"
        }
    } catch {}
}

# 4. Environment Setup
if (-not (Test-Path ".venv")) {
    Write-Host "Creating virtual environment..." -ForegroundColor Cyan
    Invoke-Expression "$pyCommand -m venv .venv"
}

Write-Host "Installing dependencies..." -ForegroundColor Cyan
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\pip install -r requirements.txt

# 5. Model Download
Write-Host "Verifying OCR Models..." -ForegroundColor Cyan
.\.venv\Scripts\python.exe scripts/download_models.py

# 6. Run App
Write-Host "Starting App..." -ForegroundColor Cyan
Start-Job -ScriptBlock {
    Start-Sleep -Seconds 5
    Start-Process "http://127.0.0.1:5000"
} | Out-Null

.\.venv\Scripts\python.exe app.py
