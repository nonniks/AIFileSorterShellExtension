#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

Write-Host "============================================================" -ForegroundColor Red
Write-Host "       AI File Sorter - Uninstallation for End Users        " -ForegroundColor Red
Write-Host "============================================================" -ForegroundColor Red
Write-Host

# Get installation directory
$programFilesDir = Join-Path -Path ${env:ProgramFiles} -ChildPath "AI File Sorter"
$mainDll = Join-Path -Path $programFilesDir -ChildPath "AIFileSorterShellExtension.dll"

try {
    # 1. Close Explorer to free any locked files
    Write-Host "1. Closing Windows Explorer (will restart automatically)..." -ForegroundColor Cyan
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    # 2. Unregister the DLL if it exists
    Write-Host "2. Unregistering extension from Windows..." -ForegroundColor Cyan
    if (Test-Path $mainDll) {
        $regResult = Start-Process "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" -ArgumentList "/unregister `"$mainDll`"" -Wait -PassThru
        if ($regResult.ExitCode -ne 0) {
            Write-Host "   [WARNING] Unregistration may have issues, but continuing anyway..." -ForegroundColor Yellow
        } else {
            Write-Host "   [OK] Successfully unregistered DLL" -ForegroundColor Green
        }
    } else {
        Write-Host "   [INFO] Main DLL not found, skipping unregistration" -ForegroundColor Yellow
    }

    # 3. Remove registry entries
    Write-Host "3. Removing from context menu..." -ForegroundColor Cyan
    New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue | Out-Null
    Remove-Item -Path "HKCR:\Directory\shellex\ContextMenuHandlers\AIFileSorter" -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "HKCR:\Directory\Background\shellex\ContextMenuHandlers\AIFileSorter" -Force -ErrorAction SilentlyContinue
    Write-Host "   [OK] Removed from context menu" -ForegroundColor Green

    # 4. Remove installation directory
    Write-Host "4. Removing installed files..." -ForegroundColor Cyan
    if (Test-Path $programFilesDir) {
        Remove-Item -Path $programFilesDir -Recurse -Force
        Write-Host "   [OK] Removed installation directory" -ForegroundColor Green
    } else {
        Write-Host "   [INFO] Installation directory not found, nothing to remove" -ForegroundColor Yellow
    }

    # 5. Restart Explorer
    Write-Host "5. Restarting Windows Explorer..." -ForegroundColor Cyan
    Start-Process explorer
    Write-Host "   [OK] Explorer restarted" -ForegroundColor Green

    Write-Host
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host "Uninstallation complete! AI File Sorter has been removed" -ForegroundColor Red
    Write-Host "from your system." -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Red
}
catch {
    Write-Host "ERROR: Uninstallation encountered the following error:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    # Restart Explorer in case of failure
    Start-Process explorer
    
    Write-Host "Windows Explorer has been restarted." -ForegroundColor Yellow
}

Write-Host
Read-Host "Press Enter to exit"
