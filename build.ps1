# SumSelect v2 build + install. Run from C:\SumSelect.
$ErrorActionPreference = "Stop"
$env:PYTHONHOME = $null
Set-Location C:\SumSelect
$py = ".\.venv\Scripts\python.exe"
& $py -m pip install -q comtypes pystray pillow pyinstaller
# pre-generate the UI Automation COM wrapper so the frozen app never has to
& $py -c "import comtypes.client as c; c.GetModule('UIAutomationCore.dll'); from comtypes.gen import UIAutomationClient; print('uia ok')"
$env:PYTHONIOENCODING="utf-8"; & $py sumselect.py --test | Out-Null
& ".\.venv\Scripts\pyinstaller.exe" --noconfirm --onedir --windowed --noupx --name SumSelect `
    --hidden-import comtypes.gen.UIAutomationClient --collect-submodules comtypes.gen `
    --exclude-module keyboard --icon icon.ico --add-data "icon.png;." sumselect.py
# install: stop old copy, replace, point autostart at the new exe, launch
Get-Process SumSelect -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 500
$dest = "$env:LOCALAPPDATA\SumSelect"
Remove-Item "$dest\app" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$dest\SumSelect.exe" -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force "$dest\app" | Out-Null
Copy-Item ".\dist\SumSelect\*" "$dest\app" -Recurse -Force
Set-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Run -Name SumSelect -Value "`"$dest\app\SumSelect.exe`""
Start-Process "$dest\app\SumSelect.exe"
"INSTALLED"
