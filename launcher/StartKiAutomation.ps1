$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$nodeDir = Join-Path $repoRoot "tools\node-v24.11.0-win-x64"
$npmCmd = Join-Path $nodeDir "npm.cmd"

if (-not (Test-Path $npmCmd)) {
  [System.Windows.Forms.MessageBox]::Show(
    "Die portable Node-Umgebung wurde nicht gefunden.`nErwartet: $npmCmd",
    "StartKiAutomation",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error
  ) | Out-Null
  exit 1
}

$env:PATH = "$nodeDir;$env:PATH"
Set-Location $repoRoot

& $npmCmd run dev -- --login-ki
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
  Read-Host "Der Startlauf ist mit Exit-Code $exitCode beendet worden. Mit Enter schliessen"
}

exit $exitCode
