# Check if script is running with necessary permissions, though usually not required for local venv
$ErrorActionPreference = "Stop"

Write-Host "Checking for Python 3.10..." -ForegroundColor Cyan

# Try to find Python 3.10 using py launcher
$pyCommand = "python"
if (Get-Command "py" -ErrorAction SilentlyContinue) {
    try {
        # Check if 3.10 is available
        $versions = py --list
        if ($versions -match "3.10") {
            $pyCommand = "py -3.10"
            Write-Host "Found Python 3.10 via py launcher." -ForegroundColor Green
        } else {
            Write-Warning "Python 3.10 not found in py launcher. Trying default python..."
        }
    } catch {
        Write-Warning "Error checking py launcher. Using default python..."
    }
}

# Create Virtual Environment if it doesn't exist
if (-not (Test-Path ".venv")) {
    Write-Host "Creating virtual environment (.venv)..." -ForegroundColor Cyan
    Invoke-Expression "$pyCommand -m venv .venv"
} else {
    Write-Host "Virtual environment (.venv) already exists." -ForegroundColor Green
}

# Activate and Install Requirements
Write-Host "Installing/Updating dependencies..." -ForegroundColor Cyan
# We use the full path to the venv python/pip to ensure we use the venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\pip install -r requirements.txt

# Start the App
Write-Host "Starting Flask Application..." -ForegroundColor Cyan
Write-Host "The application will run in this window. Press Ctrl+C to stop." -ForegroundColor Yellow

# Launch Browser in background after a short delay to allow Flask to start
Start-Job -ScriptBlock {
    Start-Sleep -Seconds 5
    Start-Process "http://127.0.0.1:5000"
} | Out-Null

# Run App
.\.venv\Scripts\python.exe app.py
