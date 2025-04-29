#Requires -RunAsAdministrator

Write-Host "==================================================" -ForegroundColor Red
Write-Host "      AI File Sorter Extension Uninstallation     " -ForegroundColor Red
Write-Host "==================================================" -ForegroundColor Red
Write-Host

# 1. Завершаем работу проводника для освобождения DLL
Write-Host "1. Stopping Windows Explorer to release DLL files..." -ForegroundColor Yellow
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# 2. Определяем пути к файлам
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$dllPath = Join-Path -Path $scriptPath -ChildPath "bin\Release\AIFileSorterShellExtension.dll"
$debugDllPath = Join-Path -Path $scriptPath -ChildPath "bin\Debug\AIFileSorterShellExtension.dll"

# 3. Отменяем регистрацию DLL
Write-Host "2. Unregistering shell extension from system..." -ForegroundColor Yellow
& "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" /unregister "$dllPath" /silent 2>$null
& "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" /unregister "$debugDllPath" /silent 2>$null

# 4. Удаляем записи реестра
Write-Host "3. Removing registry entries..." -ForegroundColor Yellow
New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue
Remove-Item -Path "HKCR:\Directory\shellex\ContextMenuHandlers\AIFileSorter" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCR:\Directory\Background\shellex\ContextMenuHandlers\AIFileSorter" -Force -ErrorAction SilentlyContinue

# 5. Удаляем файлы DLL, чтобы Visual Studio могла создать новые
Write-Host "4. Deleting DLL files to allow recompilation..." -ForegroundColor Yellow
Remove-Item -Path $dllPath -Force -ErrorAction SilentlyContinue
Remove-Item -Path $debugDllPath -Force -ErrorAction SilentlyContinue

# 6. Очищаем временные файлы компиляции
Write-Host "5. Cleaning build artifacts..." -ForegroundColor Yellow
$objPath = Join-Path -Path $scriptPath -ChildPath "obj"
if (Test-Path $objPath) {
    Remove-Item -Path "$objPath\*" -Recurse -Force -ErrorAction SilentlyContinue
}

# 7. Перезапускаем проводник
Write-Host "6. Restarting Windows Explorer..." -ForegroundColor Yellow
Start-Process explorer

Write-Host "Uninstallation complete! You can now rebuild the project in Visual Studio." -ForegroundColor Green
Write-Host
