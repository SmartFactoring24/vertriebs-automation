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

function Get-KiInstallRoot {
  $envPath = Join-Path $repoRoot ".env"

  if (-not (Test-Path $envPath)) {
    return $null
  }

  $kiAppLine = Get-Content $envPath | Where-Object { $_ -like 'KI_APP_PATH=*' } | Select-Object -First 1
  if (-not $kiAppLine) {
    return $null
  }

  $kiAppPath = $kiAppLine.Substring('KI_APP_PATH='.Length).Trim()
  if (-not $kiAppPath) {
    return $null
  }

  return Split-Path (Split-Path $kiAppPath -Parent) -Parent
}

function Stop-ProcessTree {
  param(
    [int]$RootProcessId
  )

  $allProcesses = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue)
  $killIds = New-Object System.Collections.Generic.HashSet[int]
  $queue = New-Object System.Collections.Generic.Queue[int]

  if ($killIds.Add($RootProcessId)) {
    $queue.Enqueue($RootProcessId)
  }

  while ($queue.Count -gt 0) {
    $parentId = $queue.Dequeue()
    $children = @($allProcesses | Where-Object { $_.ParentProcessId -eq $parentId })
    foreach ($child in $children) {
      if ($killIds.Add([int]$child.ProcessId)) {
        $queue.Enqueue([int]$child.ProcessId)
      }
    }
  }

  foreach ($processId in @($killIds)) {
    Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
  }
}

function Stop-KiProcessesDirect {
  $installRoot = Get-KiInstallRoot
  if (-not $installRoot) {
    Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match 'smartclient|javaw' } | Stop-Process -Force -ErrorAction SilentlyContinue
    return
  }

  Get-Process -ErrorAction SilentlyContinue | Where-Object {
    $_.ProcessName -match 'smartclient|javaw' -and $_.Path -and $_.Path.StartsWith($installRoot, [System.StringComparison]::OrdinalIgnoreCase)
  } | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Test-LauncherPrerequisites {
  $envPath = Join-Path $repoRoot ".env"
  if (-not (Test-Path $envPath)) {
    [System.Windows.Forms.MessageBox]::Show(
      "Die Datei .env wurde nicht gefunden.`nBitte zuerst die Übergabedokumentation befolgen und eine lokale .env anlegen.",
      "StartKiAutomation",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
  }

  $tsxPath = Join-Path $repoRoot "node_modules\\tsx"
  if (-not (Test-Path $tsxPath)) {
    [System.Windows.Forms.MessageBox]::Show(
      "Die Projektabhängigkeiten wurden noch nicht installiert.`nBitte im Projektordner einmal 'npm install' ausführen.",
      "StartKiAutomation",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
  }
}

function Start-BotMode {
  param(
    [string[]]$Arguments,
    [switch]$AllowAbortHotkey
  )

  $argumentList = @("run", "dev", "--") + $Arguments
  $escapedArguments = $argumentList | ForEach-Object {
    if ($_ -match '\s') {
      '"' + $_ + '"'
    } else {
      $_
    }
  }

  $process = Start-Process -FilePath $resolvedNpm -ArgumentList $escapedArguments -WorkingDirectory $repoRoot -PassThru

  while (-not $process.HasExited) {
    if ($AllowAbortHotkey -and [Console]::KeyAvailable) {
      $keyInfo = [Console]::ReadKey($true)
      if ($keyInfo.Key -eq [ConsoleKey]::F12) {
        Write-Host ""
        Write-Host "Abbruch durch Benutzer erkannt. Bot und KI-Prozesse werden beendet..." -ForegroundColor Yellow
        Stop-ProcessTree -RootProcessId $process.Id
        Start-Sleep -Milliseconds 300
        Stop-KiProcessesDirect
        return 130
      }
    }

    Start-Sleep -Milliseconds 100
  }

  return $process.ExitCode
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
Write-Host "- Während des laufenden Startvorgangs kann jederzeit F12 gedrückt werden, um den Bot sauber abzubrechen." -ForegroundColor Yellow
Write-Host ""
Write-Host "Firmeninterne Nutzung:" -ForegroundColor Yellow
Write-Host "- Dieses Programm wurde von Leo Mitteneder und Moritz Rolle mit Unterstützung durch Codex erstellt."
Write-Host "- Es ist ausschließlich für die firmeninterne Nutzung bestimmt."
Write-Host "- Eine Vervielfältigung oder Weitergabe ist nicht zulässig."
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$enterDebugMenu = Read-DebugHotkey -TimeoutSeconds 4
Test-LauncherPrerequisites

if ($enterDebugMenu) {
  $exitCode = Show-DebugMenu
} else {
  $exitCode = Start-BotMode @("--login-ki") -AllowAbortHotkey
}

if ($exitCode -ne 0) {
  Read-Host "Der Startlauf wurde mit Exit-Code $exitCode beendet. Mit Enter schließen"
}

exit $exitCode
