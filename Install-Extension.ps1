#Requires -RunAsAdministrator

Write-Host "==================================================" -ForegroundColor Green
Write-Host "      AI File Sorter Extension Installation       " -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
Write-Host

# 1. Завершаем работу проводника для освобождения DLL
Write-Host "1. Stopping Windows Explorer to release DLL files..." -ForegroundColor Cyan
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# 2. Определяем пути к файлам
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputPath = Join-Path -Path $scriptPath -ChildPath "bin\Release"
$dllPath = Join-Path -Path $outputPath -ChildPath "AIFileSorterShellExtension.dll"

# 3. Проверяем наличие DLL
Write-Host "2. Checking DLL existence..." -ForegroundColor Cyan
if (-not (Test-Path $dllPath)) {
    Write-Host "[ERROR] DLL not found at: $dllPath" -ForegroundColor Red
    Write-Host "Please compile the project in Release mode first" -ForegroundColor Yellow
    
    # Перезапустим проводник и выйдем
    Start-Process explorer
    Write-Host "Explorer restarted. Installation aborted." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# 4. Копируем необходимые библиотеки зависимостей
Write-Host "3. Copying dependency libraries..." -ForegroundColor Cyan
$dependencies = @(
    "Newtonsoft.Json.dll",
    "SharpShell.dll"
)

foreach ($dep in $dependencies) {
    # Проверяем наличие библиотеки в выходной папке
    $depPath = Join-Path -Path $outputPath -ChildPath $dep
    if (-not (Test-Path $depPath)) {
        Write-Host "   [WARNING] Dependency not found: $dep" -ForegroundColor Yellow
    } else {
        Write-Host "   [OK] Found dependency: $dep" -ForegroundColor Green
    }
}

# 5. Создаем конфигурационный файл
Write-Host "4. Creating configuration file..." -ForegroundColor Cyan
$configContent = @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <runtime>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Newtonsoft.Json" publicKeyToken="30ad4fe6b2a6aeed" culture="neutral" />
        <bindingRedirect oldVersion="0.0.0.0-13.0.0.0" newVersion="13.0.0.0" />
      </dependentAssembly>
    </assemblyBinding>
  </runtime>
</configuration>
"@

$configPath = Join-Path -Path $outputPath -ChildPath "AIFileSorterShellExtension.dll.config"
Set-Content -Path $configPath -Value $configContent -Force
Write-Host "   [OK] Created config file: $configPath" -ForegroundColor Green

# 6. Регистрируем COM-объект
Write-Host "5. Registering COM server..." -ForegroundColor Cyan
& "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "$dllPath" /codebase
if ($LASTEXITCODE -ne 0) {
    Write-Host "   [ERROR] Registration failed!" -ForegroundColor Red
    
    # Перезапустим проводник и выйдем
    Start-Process explorer
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "   [OK] Registration successful" -ForegroundColor Green

# 7. Создаем записи в реестре
Write-Host "6. Creating registry entries..." -ForegroundColor Cyan
New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue
New-Item -Path "HKCR:\Directory\shellex\ContextMenuHandlers\AIFileSorter" -Force -ErrorAction SilentlyContinue | Out-Null
Set-ItemProperty -Path "HKCR:\Directory\shellex\ContextMenuHandlers\AIFileSorter" -Name "(Default)" -Value "{0702a7bb-5000-58ca-256a-7306f3614924}" -Force
New-Item -Path "HKCR:\Directory\Background\shellex\ContextMenuHandlers\AIFileSorter" -Force -ErrorAction SilentlyContinue | Out-Null
Set-ItemProperty -Path "HKCR:\Directory\Background\shellex\ContextMenuHandlers\AIFileSorter" -Name "(Default)" -Value "{0702a7bb-5000-58ca-256a-7306f3614924}" -Force
Write-Host "   [OK] Registry entries created" -ForegroundColor Green

# 8. Перезапускаем проводник
Write-Host "7. Restarting Windows Explorer..." -ForegroundColor Cyan
Start-Process explorer
Write-Host "   [OK] Explorer restarted" -ForegroundColor Green

Write-Host
Write-Host "Installation complete! Try right-clicking on a folder to see the 'Sort Files' option." -ForegroundColor Green
Write-Host "If you don't see the option, try restarting your computer." -ForegroundColor Yellow
Write-Host
