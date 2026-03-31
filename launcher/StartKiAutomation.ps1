$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$nodeDir = Join-Path $repoRoot "tools\node-v24.11.0-win-x64"
$npmCmd = Join-Path $nodeDir "npm.cmd"
$resolvedNpm = $null

try {
  $resolvedNpm = (Get-Command npm.cmd -ErrorAction Stop).Source
} catch {
  if (Test-Path $npmCmd) {
    $resolvedNpm = $npmCmd
    $env:PATH = "$nodeDir;$env:PATH"
  }
}

if (-not $resolvedNpm) {
  [System.Windows.Forms.MessageBox]::Show(
    "Weder eine globale noch die portable Node-Umgebung wurde gefunden.`nErwartet lokal: $npmCmd",
    "StartKiAutomation",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error
  ) | Out-Null
  exit 1
}

Set-Location $repoRoot

& $resolvedNpm run dev -- --login-ki
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
  Read-Host "Der Startlauf ist mit Exit-Code $exitCode beendet worden. Mit Enter schliessen"
}

exit $exitCode
