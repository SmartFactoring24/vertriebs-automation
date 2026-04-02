$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$bundledNodeDir = Join-Path $repoRoot "tools\node-v24.11.0-win-x64"
$bundledNodeExe = Join-Path $bundledNodeDir "node.exe"
$bundledNpmCmd = Join-Path $bundledNodeDir "npm.cmd"
$defaultTesseractExe = "C:\Program Files\Tesseract-OCR\tesseract.exe"
$defaultTessdataDir = Join-Path $repoRoot "data\tessdata"
$deuTrainedDataPath = Join-Path $defaultTessdataDir "deu.traineddata"
$deuTrainedDataUrl = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/deu.traineddata"

function Show-InstallerMessage {
  param(
    [string]$Message,
    [string]$Title = "InstallKiAutomation",
    [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
  )

  [System.Windows.Forms.MessageBox]::Show(
    $Message,
    $Title,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    $Icon
  ) | Out-Null
}

function Resolve-NodeAndNpm {
  $resolvedNode = $null
  $resolvedNpm = $null

  try {
    $resolvedNode = (Get-Command node.exe -ErrorAction Stop).Source
  } catch {
    if (Test-Path $bundledNodeExe) {
      $resolvedNode = $bundledNodeExe
      $env:PATH = "$bundledNodeDir;$env:PATH"
    }
  }

  try {
    $resolvedNpm = (Get-Command npm.cmd -ErrorAction Stop).Source
  } catch {
    if (Test-Path $bundledNpmCmd) {
      $resolvedNpm = $bundledNpmCmd
    }
  }

  if (-not $resolvedNode -or -not $resolvedNpm) {
    throw "Es konnte weder eine globale noch eine portable Node/NPM-Umgebung gefunden werden."
  }

  return [PSCustomObject]@{
    Node = $resolvedNode
    Npm = $resolvedNpm
  }
}

function Invoke-Step {
  param(
    [string]$Label,
    [scriptblock]$Action
  )

  Write-Host ""
  Write-Host ">> $Label" -ForegroundColor Cyan
  & $Action
}

function Ensure-Directory {
  param([string]$Path)
  New-Item -ItemType Directory -Force -Path $Path | Out-Null
}

function Ensure-EnvFile {
  $envPath = Join-Path $repoRoot ".env"
  $envExamplePath = Join-Path $repoRoot ".env.example"

  if (Test-Path $envPath) {
    return $false
  }

  if (-not (Test-Path $envExamplePath)) {
    throw "Die Datei .env.example wurde nicht gefunden."
  }

  Copy-Item -LiteralPath $envExamplePath -Destination $envPath -Force
  return $true
}

function Get-EnvMap {
  param([string]$Path)

  $map = @{}
  if (-not (Test-Path $Path)) {
    return $map
  }

  foreach ($line in Get-Content $Path) {
    if ($line -match "^\s*#" -or $line -notmatch "=") {
      continue
    }

    $parts = $line -split "=", 2
    if ($parts.Count -eq 2) {
      $map[$parts[0]] = $parts[1]
    }
  }

  return $map
}

function Set-EnvValue {
  param(
    [string]$Path,
    [string]$Key,
    [string]$Value
  )

  $lines = if (Test-Path $Path) { Get-Content $Path } else { @() }
  $updated = $false

  for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match "^$([regex]::Escape($Key))=") {
      $lines[$i] = "$Key=$Value"
      $updated = $true
      break
    }
  }

  if (-not $updated) {
    $lines += "$Key=$Value"
  }

  Set-Content -LiteralPath $Path -Value $lines
}

function Convert-SecureStringToPlainText {
  param([Security.SecureString]$SecureString)

  $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
  try {
    return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
  } finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
  }
}

function Prompt-LoginData {
  param([string]$EnvPath)

  $envMap = Get-EnvMap -Path $EnvPath
  $existingUsername = $envMap["KI_USERNAME"]
  $existingPassword = $envMap["KI_PASSWORD"]

  Write-Host ""
  Write-Host "Bitte jetzt die lokalen KI-Logindaten hinterlegen." -ForegroundColor Yellow

  $usernamePrompt = "KI_USERNAME"
  if ($existingUsername) {
    $usernamePrompt += " [$existingUsername]"
  }

  $enteredUsername = Read-Host $usernamePrompt
  if ([string]::IsNullOrWhiteSpace($enteredUsername)) {
    $enteredUsername = $existingUsername
  }

  if ([string]::IsNullOrWhiteSpace($enteredUsername)) {
    throw "KI_USERNAME darf nicht leer bleiben."
  }

  $finalPassword = $existingPassword
  $keepExistingPassword = $false
  if (-not [string]::IsNullOrWhiteSpace($existingPassword)) {
    $keepResponse = Read-Host "Vorhandenes KI_PASSWORD beibehalten? [J/n]"
    $keepExistingPassword = [string]::IsNullOrWhiteSpace($keepResponse) -or $keepResponse -match "^(j|ja|y|yes)$"
  }

  if (-not $keepExistingPassword) {
    $securePassword = Read-Host "KI_PASSWORD" -AsSecureString
    $finalPassword = Convert-SecureStringToPlainText -SecureString $securePassword
  }

  if ([string]::IsNullOrWhiteSpace($finalPassword)) {
    throw "KI_PASSWORD darf nicht leer bleiben."
  }

  Set-EnvValue -Path $EnvPath -Key "KI_USERNAME" -Value $enteredUsername
  Set-EnvValue -Path $EnvPath -Key "KI_PASSWORD" -Value $finalPassword
}

function Ensure-ProjectDirectories {
  $dirs = @(
    (Join-Path $repoRoot "data"),
    (Join-Path $repoRoot "data\exports"),
    (Join-Path $repoRoot "data\logs"),
    (Join-Path $repoRoot "data\screenshots"),
    (Join-Path $repoRoot "data\state"),
    (Join-Path $repoRoot "data\vision"),
    (Join-Path $repoRoot "data\tessdata"),
    (Join-Path $repoRoot "browser-profiles"),
    (Join-Path $repoRoot "browser-profiles\whatsapp-profile")
  )

  foreach ($dir in $dirs) {
    Ensure-Directory -Path $dir
  }
}

function Ensure-TesseractInstalled {
  if (Test-Path $defaultTesseractExe) {
    return
  }

  $winget = $null
  try {
    $winget = (Get-Command winget.exe -ErrorAction Stop).Source
  } catch {
    throw "Tesseract ist nicht installiert und winget ist auf diesem Rechner nicht verfuegbar."
  }

  Write-Host "Tesseract wurde nicht gefunden. Installation via winget wird gestartet..." -ForegroundColor Yellow
  & $winget install --id UB-Mannheim.TesseractOCR -e --accept-package-agreements --accept-source-agreements

  if (-not (Test-Path $defaultTesseractExe)) {
    throw "Tesseract konnte nicht installiert werden oder wurde danach nicht gefunden."
  }
}

function Ensure-GermanTessdata {
  Ensure-Directory -Path $defaultTessdataDir

  if (Test-Path $deuTrainedDataPath) {
    return
  }

  Write-Host "Deutsche OCR-Sprachdatei wird heruntergeladen..." -ForegroundColor Yellow
  Invoke-WebRequest -Uri $deuTrainedDataUrl -OutFile $deuTrainedDataPath
}

function Install-NpmDependencies {
  param([string]$NpmCmd)

  Set-Location $repoRoot
  & $NpmCmd install
}

function Install-PlaywrightChromium {
  param([string]$NodeExe)

  $playwrightCli = Join-Path $repoRoot "node_modules\playwright\cli.js"
  if (-not (Test-Path $playwrightCli)) {
    throw "Die Playwright-CLI wurde nicht gefunden. npm install scheint nicht vollstaendig durchgelaufen zu sein."
  }

  Set-Location $repoRoot
  & $NodeExe $playwrightCli install chromium
}

function Test-InstallerScript {
  param([string]$ScriptPath)

  $tokens = $null
  $errors = $null
  [void][System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errors)
  if ($errors.Count -gt 0) {
    throw "Das Installer-Skript enthaelt Syntaxfehler."
  }
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Vertriebs-Automation" -ForegroundColor Cyan
Write-Host "Installer-Modus" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Dieser Installer bereitet einen neuen Rechner fuer den Bot vor." -ForegroundColor Yellow
Write-Host "Automatisch erledigt werden:" -ForegroundColor Yellow
Write-Host "- lokale Projektordner anlegen"
Write-Host "- Node/NPM aufloesen (global oder portable Repo-Version)"
Write-Host "- npm-Abhaengigkeiten installieren"
Write-Host "- Playwright Chromium installieren"
Write-Host "- Tesseract pruefen bzw. via winget installieren"
Write-Host "- deutsche OCR-Sprachdatei lokal im Repo ablegen"
Write-Host "- .env aus .env.example anlegen, falls sie noch fehlt"
Write-Host "- KI_USERNAME und KI_PASSWORD interaktiv in die lokale .env schreiben"
Write-Host ""

try {
  $runtime = Resolve-NodeAndNpm

  Invoke-Step "Projektordner vorbereiten" {
    Ensure-ProjectDirectories
  }

  Invoke-Step "Tesseract pruefen" {
    Ensure-TesseractInstalled
  }

  Invoke-Step "Deutsche OCR-Sprachdatei pruefen" {
    Ensure-GermanTessdata
  }

  Invoke-Step "Node-Abhaengigkeiten installieren" {
    Install-NpmDependencies -NpmCmd $runtime.Npm
  }

  Invoke-Step "Playwright Chromium installieren" {
    Install-PlaywrightChromium -NodeExe $runtime.Node
  }

  $envCreated = $false
  Invoke-Step ".env pruefen" {
    $script:envCreated = Ensure-EnvFile
  }

  Invoke-Step "KI-Logindaten erfassen" {
    Prompt-LoginData -EnvPath (Join-Path $repoRoot ".env")
  }

  Test-InstallerScript -ScriptPath $PSCommandPath

  Write-Host ""
  Write-Host "Installation abgeschlossen." -ForegroundColor Green
  Write-Host ""
  Write-Host "Projektordner: $repoRoot"
  Write-Host "Node: $($runtime.Node)"
  Write-Host "NPM: $($runtime.Npm)"
  Write-Host "Tesseract: $defaultTesseractExe"
  Write-Host "Tessdata: $defaultTessdataDir"
  Write-Host ""

  if ($envCreated) {
    Show-InstallerMessage -Title "InstallKiAutomation" -Icon Warning -Message (
      "Die Datei .env wurde neu aus .env.example angelegt.`n`n" +
      "KI_USERNAME und KI_PASSWORD wurden direkt im Installer abgefragt.`n" +
      "Bitte jetzt nur noch die uebrigen geraetespezifischen Werte pruefen, bevor du den Bot startest."
    )
  } else {
    Show-InstallerMessage -Title "InstallKiAutomation" -Message (
      "Die benoetigten Ressourcen wurden vorbereitet und die KI-Logindaten in der lokalen .env hinterlegt.`n`n" +
      "Du kannst danach den Bot ueber StartKiAutomation.exe oder das PowerShell-Skript starten."
    )
  }
} catch {
  Show-InstallerMessage -Title "InstallKiAutomation" -Icon Error -Message (
    "Die Installation konnte nicht vollstaendig abgeschlossen werden.`n`n" +
    $_.Exception.Message
  )
  exit 1
}
