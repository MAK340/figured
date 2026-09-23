# Figured build: tests, PyInstaller bundle, installer, portable zip.
#   .\build.ps1            build dist\Figured-<ver>-setup.exe and dist\Figured-<ver>-portable.zip
#   .\build.ps1 -Install   ...then install it on this PC (silent, keeps settings)
param([switch]$Install)
# native tools write progress to stderr; failures are caught via $LASTEXITCODE instead
$ErrorActionPreference = "Continue"
$env:PYTHONHOME = $null; $env:PYTHONPATH = $null
Set-Location $PSScriptRoot
$py = ".\.venv\Scripts\python.exe"
if (-not (Test-Path $py)) { python -m venv .venv }
& $py -m pip install -q comtypes pystray pillow pyinstaller
# pre-generate the UI Automation COM wrapper so the frozen app never has to
& $py -c "import comtypes.client as c; c.GetModule('UIAutomationCore.dll'); from comtypes.gen import UIAutomationClient; print('uia ok')"
$env:PYTHONIOENCODING = "utf-8"
& $py figured.py --test | Out-Null
if ($LASTEXITCODE) { throw "self-test failed" }
$ver = (Select-String -Path figured.py -Pattern '^VERSION = "(.+)"').Matches[0].Groups[1].Value

& ".\.venv\Scripts\pyinstaller.exe" --noconfirm --onedir --windowed --noupx --name Figured `
    --hidden-import comtypes.gen.UIAutomationClient --collect-submodules comtypes.gen `
    --exclude-module keyboard --icon icon.ico --add-data "icon.png;." figured.py
if ($LASTEXITCODE) { throw "pyinstaller failed" }

$iscc = @("$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe", "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe") |
    Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw "Inno Setup 6 not found: winget install JRSoftware.InnoSetup" }
& $iscc /Q "/DAppVersion=$ver" installer.iss
if ($LASTEXITCODE) { throw "installer build failed" }

$zip = "dist\Figured-$ver-portable.zip"
Remove-Item $zip -ErrorAction SilentlyContinue
Compress-Archive -Path dist\Figured -DestinationPath $zip
Get-Item "dist\Figured-$ver-setup.exe", $zip | Select-Object Name, @{n = "MB"; e = { [math]::Round($_.Length / 1MB, 1) } }

if ($Install) {
    Start-Process "dist\Figured-$ver-setup.exe" -ArgumentList "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART" -Wait
    "INSTALLED $ver"
}