param(
    [switch]$SkipInstaller
)
$ErrorActionPreference = "Stop"
function Find-InnoSetupCompiler {
    $command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if (-not $command) {
        $command = Get-Command iscc -ErrorAction SilentlyContinue
    }
    if ($command) {
        return $command.Source
    }
    $candidatePaths = @(
        Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"
        Join-Path $env:ProgramFiles "Inno Setup 6\ISCC.exe"
        Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe"
    )
    foreach ($candidate in $candidatePaths) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }
    return $null
}
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot
$PythonCommand = if (Get-Command py -ErrorAction SilentlyContinue) { "py" } else { "python" }
Write-Host "Installing Python dependencies into the active Python environment..."
& $PythonCommand -m pip install -r requirements.txt pyinstaller
Write-Host "Installing Playwright Chromium into the package-local browser cache..."
$env:PLAYWRIGHT_BROWSERS_PATH = "0"
& $PythonCommand -m playwright install chromium
Write-Host "Cleaning previous PyInstaller output..."
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
Write-Host "Building CustomerLetterGenerator..."
& $PythonCommand -m PyInstaller CustomerLetterGenerator.spec
Write-Host "Copying editable runtime files..."
Copy-Item config.yaml dist\CustomerLetterGenerator\ -Force
Copy-Item customer_letter_template.docx dist\CustomerLetterGenerator\ -Force
New-Item -ItemType Directory -Force -Path dist\CustomerLetterGenerator\out | Out-Null
$CliPath = "dist\CustomerLetterGenerator\_internal\playwright\driver\package\cli.js"
$ChromiumPath = "dist\CustomerLetterGenerator\_internal\playwright\driver\package\.local-browsers\chromium-1134"
Write-Host "Verifying bundled Playwright files..."
Write-Host "$CliPath => $(Test-Path $CliPath)"
Write-Host "$ChromiumPath => $(Test-Path $ChromiumPath)"
if (-not (Test-Path $CliPath)) {
    throw "Playwright cli.js was not bundled at $CliPath"
}
if (-not (Test-Path $ChromiumPath)) {
    throw "Playwright Chromium was not bundled at $ChromiumPath"
}
if ($SkipInstaller) {
    Write-Host "Skipping installer creation because -SkipInstaller was provided."
    Write-Host "Build completed successfully."
    return
}
$InstallerCompiler = Find-InnoSetupCompiler
if (-not $InstallerCompiler) {
    throw @"
Inno Setup compiler (ISCC.exe) was not found, so the installer was not created.

Install Inno Setup 6 on this Windows build machine, then rerun this script:
  winget install JRSoftware.InnoSetup

If Inno Setup is already installed, add its folder to PATH or verify this file exists:
  C:\Program Files (x86)\Inno Setup 6\ISCC.exe

To intentionally build only the dist folder without an installer, run:
  powershell -ExecutionPolicy Bypass -File ./build_windows.ps1 -SkipInstaller
"@
}
$Version = if ($env:APP_VERSION) { $env:APP_VERSION } else { "1.0.0" }
Write-Host "Building Windows installer version $Version with $InstallerCompiler..."
& $InstallerCompiler "/DMyAppVersion=$Version" CustomerLetterGeneratorInstaller.iss
Write-Host "Installer created in the installer folder."
Write-Host "Build completed successfully."