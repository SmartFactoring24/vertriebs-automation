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

function Start-BotMode {
  param(
    [string[]]$Arguments
  )

  & $resolvedNpm run dev -- @Arguments
  return $LASTEXITCODE
}

function Show-DebugMenu {
  Write-Host ""
  Write-Host "============================================================" -ForegroundColor Yellow
  Write-Host "Debug- und Troubleshooting-Menü" -ForegroundColor Yellow
  Write-Host "============================================================" -ForegroundColor Yellow
  Write-Host "1 - Recovery- und Crash-Test starten"
  Write-Host "2 - KI-Statusdiagnose anzeigen"
  Write-Host "3 - Alle KI-Prozesse sauber beenden"
  Write-Host "0 - Normalen Start fortsetzen"
  Write-Host ""

  do {
    $selection = Read-Host "Bitte Auswahl eingeben"

    switch ($selection) {
      "1" {
        Write-Host ""
        Write-Host "Recovery- und Crash-Test wird gestartet..." -ForegroundColor Yellow
        return Start-BotMode @("--test-recovery")
      }
      "2" {
        Write-Host ""
        Write-Host "KI-Statusdiagnose wird gestartet..." -ForegroundColor Yellow
        return Start-BotMode @("--inspect-ki")
      }
      "3" {
        Write-Host ""
        Write-Host "KI-Prozesse werden beendet..." -ForegroundColor Yellow
        return Start-BotMode @("--close-ki")
      }
      "0" {
        Write-Host ""
        Write-Host "Normaler Start wird fortgesetzt..." -ForegroundColor Green
        return Start-BotMode @("--login-ki")
      }
      default {
        Write-Host "Ungültige Eingabe. Bitte 0, 1, 2 oder 3 wählen." -ForegroundColor Red
      }
    }
  } while ($true)
}

function Read-DebugHotkey {
  param(
    [int]$TimeoutSeconds = 4
  )

  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  while ((Get-Date) -lt $deadline) {
    if ([Console]::KeyAvailable) {
      $keyInfo = [Console]::ReadKey($true)
      if ($keyInfo.Key -eq [ConsoleKey]::F4) {
        return $true
      }
    }

    Start-Sleep -Milliseconds 100
  }

  return $false
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Vertriebs-Automation" -ForegroundColor Cyan
Write-Host "Launcher-Modus: Login-Testlauf" -ForegroundColor Cyan
Write-Host ""
Write-Host "Hinweis:" -ForegroundColor Yellow
Write-Host "- Für den laufenden Automationsbetrieb muss das PowerShell-Fenster geöffnet bleiben."
Write-Host "- Dieser Launcher startet einen einmaligen Login-Testlauf und endet danach automatisch."
Write-Host "- Crash- und Diagnoseprotokolle werden im Ordner .\data\logs gespeichert."
Write-Host "- Die DVAG-2FA muss während des Anmeldevorgangs manuell freigegeben werden."
Write-Host "- Bitte halten Sie Ihr Freigabegerät bereit, damit der Bot den Startvorgang fortsetzen kann."
Write-Host "- Für Debug und Troubleshooting direkt beim Start innerhalb von 4 Sekunden F4 drücken." -ForegroundColor Yellow
Write-Host ""
Write-Host "Firmeninterne Nutzung:" -ForegroundColor Yellow
Write-Host "- Dieses Programm wurde von Leo Mitteneder und Moritz Rolle mit Unterstützung durch Codex erstellt."
Write-Host "- Es ist ausschließlich für die firmeninterne Nutzung bestimmt."
Write-Host "- Eine Vervielfältigung oder Weitergabe ist nicht zulässig."
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$enterDebugMenu = Read-DebugHotkey -TimeoutSeconds 4

if ($enterDebugMenu) {
  $exitCode = Show-DebugMenu
} else {
  $exitCode = Start-BotMode @("--login-ki")
}

if ($exitCode -ne 0) {
  Read-Host "Der Startlauf wurde mit Exit-Code $exitCode beendet. Mit Enter schließen"
}

exit $exitCode
