#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

Write-Host "============================================================" -ForegroundColor Green
Write-Host "       AI File Sorter - Installation for End Users          " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host

# Get current script directory
$installDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Verify files exist
$mainDll = Join-Path -Path $installDir -ChildPath "AIFileSorterShellExtension.dll"
if (-not (Test-Path $mainDll)) {
    Write-Host "ERROR: Main DLL not found at: $mainDll" -ForegroundColor Red
    Write-Host "Please make sure all files are in the same directory as this script." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

try {
    # 1. Close Explorer to free any locked files
    Write-Host "1. Closing Windows Explorer (will restart automatically)..." -ForegroundColor Cyan
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    # 2. Register the DLL
    Write-Host "2. Registering extension with Windows..." -ForegroundColor Cyan
    $regResult = Start-Process "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" -ArgumentList "`"$mainDll`" /codebase" -Wait -PassThru
    if ($regResult.ExitCode -ne 0) {
        throw "DLL registration failed with exit code: $($regResult.ExitCode)"
    }
    Write-Host "   [OK] Registration successful" -ForegroundColor Green

    # 3. Create registry entries
    Write-Host "3. Adding to context menu..." -ForegroundColor Cyan
    New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue | Out-Null
    New-Item -Path "HKCR:\Directory\shellex\ContextMenuHandlers\AIFileSorter" -Force | Out-Null
    Set-ItemProperty -Path "HKCR:\Directory\shellex\ContextMenuHandlers\AIFileSorter" -Name "(Default)" -Value "{0702a7bb-5000-58ca-256a-7306f3614924}" -Force
    New-Item -Path "HKCR:\Directory\Background\shellex\ContextMenuHandlers\AIFileSorter" -Force | Out-Null
    Set-ItemProperty -Path "HKCR:\Directory\Background\shellex\ContextMenuHandlers\AIFileSorter" -Name "(Default)" -Value "{0702a7bb-5000-58ca-256a-7306f3614924}" -Force
    Write-Host "   [OK] Added to context menu" -ForegroundColor Green

    # 4. Create destination directory and copy files
    Write-Host "4. Installing files to Program Files..." -ForegroundColor Cyan
    $programFilesDir = Join-Path -Path ${env:ProgramFiles} -ChildPath "AI File Sorter"
    if (-not (Test-Path $programFilesDir)) {
        New-Item -Path $programFilesDir -ItemType Directory -Force | Out-Null
    }

    # Copy all files to program files
    Get-ChildItem -Path $installDir -File | ForEach-Object {
        Copy-Item -Path $_.FullName -Destination $programFilesDir -Force
        Write-Host "   [OK] Installed: $($_.Name)" -ForegroundColor Green
    }

    # 5. Restart Explorer
    Write-Host "5. Restarting Windows Explorer..." -ForegroundColor Cyan
    Start-Process explorer
    Write-Host "   [OK] Explorer restarted" -ForegroundColor Green

    Write-Host
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "Installation complete! You can now right-click on any folder" -ForegroundColor Green
    Write-Host "and use the 'Sort Files' option in the context menu." -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    
    Write-Host
    Write-Host "If the option doesn't appear in the context menu, try restarting your computer." -ForegroundColor Yellow
}
catch {
    Write-Host "ERROR: Installation failed with the following error:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    # Restart Explorer in case of failure
    Start-Process explorer
    
    Write-Host "Windows Explorer has been restarted." -ForegroundColor Yellow
}

Write-Host
Read-Host "Press Enter to exit"
