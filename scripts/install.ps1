param(
    [string]$Python = 'python',
    [string]$PluginRoot = "$env:APPDATA\FlowLauncher\Plugins"
)
$ErrorActionPreference = 'Stop'
$source = Split-Path $PSScriptRoot -Parent
$destination = Join-Path $PluginRoot 'OutlookWorkspace-1.0.0'
& $Python -c "import sys, tkinter; assert sys.version_info >= (3, 10), 'Python 3.10 or newer required'"
if ($LASTEXITCODE -ne 0) { throw 'Install Python 3.10+ with Tcl/Tk first.' }
New-Item -ItemType Directory -Force -Path $destination | Out-Null
foreach ($name in @('main.py', 'ui.py', 'plugin.json', 'requirements.txt', 'README.md', 'outlook_flow', 'Images')) {
    Copy-Item -Path (Join-Path $source $name) -Destination $destination -Recurse -Force
}
& $Python -m pip install --upgrade --target (Join-Path $destination 'lib') -r (Join-Path $destination 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
Write-Host "Installed to $destination. Set Flow Launcher's Python interpreter to the same Python executable, restart Flow, and type: outlook settings"
