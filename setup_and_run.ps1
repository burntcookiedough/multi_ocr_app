# Check if script is running with necessary permissions
$ErrorActionPreference = "Stop"

Write-Host "Running Automated Setup Script v3 (Portable Mode)" -ForegroundColor Magenta

function Check-Command($cmd, $name) {
    if (Get-Command $cmd -ErrorAction SilentlyContinue) {
        Write-Host "Found $name in PATH." -ForegroundColor Green
        return $true
    }
    return $false
}

Write-Host "Checking System Requirements..." -ForegroundColor Cyan

# 1. Tesseract Check & Auto-Install
$tesseractFound = Check-Command "tesseract" "Tesseract-OCR"

# Check local portable path first
$localTesseractPath = "$PWD\tesseract-ocr"
if (-not $tesseractFound -and (Test-Path "$localTesseractPath\tesseract.exe")) {
    Write-Host "Found local Tesseract at $localTesseractPath. Configuring..." -ForegroundColor Green
    $env:Path = "$localTesseractPath;" + $env:Path
    $env:TESSDATA_PREFIX = "$localTesseractPath\tessdata"
    $tesseractFound = $true
}

if (-not $tesseractFound) {
    # Check common system locations
    $commonPaths = @(
        "C:\Program Files\Tesseract-OCR",
        "C:\Program Files (x86)\Tesseract-OCR",
        "$env:LOCALAPPDATA\Tesseract-OCR"
    )
    
    foreach ($path in $commonPaths) {
        if (Test-Path "$path\tesseract.exe") {
            Write-Host "Found properly installed Tesseract at $path. Adding to PATH..." -ForegroundColor Green
            $env:Path = "$path;" + $env:Path
            $tesseractFound = $true
            break
        }
    }
}

if (-not $tesseractFound) {
    Write-Warning "Tesseract-OCR not found."
    Write-Host "Downloading Tesseract Installer..." -ForegroundColor Yellow
    
    # Using specific version from UB-Mannheim
    $installerUrl = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.3/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
    $installerPath = "$PWD\tesseract_installer.exe"
    
    try {
        # GitHub requires UserAgent
        Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UserAgent "PowerShell" -ErrorAction Stop
        Write-Host "Installer downloaded to $installerPath" -ForegroundColor Green
        
        Write-Host "Installing Tesseract locally to $localTesseractPath..." -ForegroundColor Yellow
        # Silent install to local directory
        $proc = Start-Process -FilePath $installerPath -ArgumentList "/S", "/D=$localTesseractPath" -Wait -PassThru
        
        if (Test-Path "$localTesseractPath\tesseract.exe") {
            Write-Host "Tesseract installed successfully!" -ForegroundColor Green
            $env:Path = "$localTesseractPath;" + $env:Path
            $env:TESSDATA_PREFIX = "$localTesseractPath\tessdata"
            
            # Clean up installer
            Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
        } else {
            Write-Error "Silent install failed or required elevation. Please check $installerPath manually."
        }
        
    } catch {
        Write-Warning "Auto-download/install failed: $_"
        Write-Host "Please manually download from: https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor Red
    }
}

# 2. Poppler Check & Auto-Install
if (-not (Check-Command "pdftoppm" "Poppler")) {
    Write-Warning "Poppler not found. Downloading..."
    
    $popplerUrl = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.02.0-0/Release-24.02.0-0.zip"
    $zipPath = "$PWD\poppler.zip"
    $extractPath = "$PWD\poppler_lib"
    
    try {
        if (-not (Test-Path $zipPath)) {
            Write-Host "Downloading Poppler..." -ForegroundColor Yellow
            Invoke-WebRequest -Uri $popplerUrl -OutFile $zipPath -UserAgent "PowerShell" -ErrorAction Stop
        }
        
        if (-not (Test-Path $extractPath)) {
            Write-Host "Extracting Poppler..." -ForegroundColor Yellow
            Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
        }
        
        $binPath = Get-ChildItem -Path $extractPath -Recurse -Filter "pdftoppm.exe" | Select-Object -First 1 -ExpandProperty DirectoryName
        if ($binPath) {
            Write-Host "Found Poppler bin at $binPath. Adding to PATH..." -ForegroundColor Green
            $env:Path = "$binPath;" + $env:Path
        }
        
    } catch {
        Write-Warning "Failed to setup Poppler: $_"
    }
}

# 3. Python Check (Enforcing 3.10 - 3.12)
$pyCommand = $null
$supportedVersions = @("3.10", "3.11", "3.12")
$targetVersion = $null

if (Get-Command "py" -ErrorAction SilentlyContinue) {
    try {
        $versions = py --list
        foreach ($ver in $supportedVersions) {
            if ($versions -match " $ver") {
                $pyCommand = "py -$ver"
                $targetVersion = $ver
                Write-Host "Found Python $ver via py launcher." -ForegroundColor Green
                break
            }
        }
    } catch {}
}

# Fallback: Check if 'python' in PATH is a supported version
if (-not $pyCommand) {
    try {
        $currentVer = python --version 2>&1
        foreach ($ver in $supportedVersions) {
            if ($currentVer -match " $ver") {
                $pyCommand = "python"
                $targetVersion = $ver
                Write-Host "Default 'python' is version $ver. Using it." -ForegroundColor Green
                break
            }
        }
    } catch {}
}

if (-not $pyCommand) {
    Write-Error "CRITICAL: No compatible Python version found (3.10, 3.11, or 3.12)."
    Write-Host "PaddleOCR requires Numpy < 2.0.0, which does not support Python 3.13 yet." -ForegroundColor Red
    Write-Host "Please install Python 3.10 from python.org and try again." -ForegroundColor Yellow
    exit 1
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
