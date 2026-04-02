import { mkdir, rm } from "node:fs/promises";
import path from "node:path";
import type { AppConfig } from "../config.js";
import type { KiWindowInfo } from "../connectors/ki-desktop.js";
import type { VisionCaptureArtifact, VisionCaptureRegion } from "./types.js";

function buildTreeRegion(config: AppConfig, window: KiWindowInfo): VisionCaptureRegion {
  const x = Math.round(window.x + window.width * config.kiTreeRegionRelX);
  const y = Math.round(window.y + window.height * config.kiTreeRegionRelY);
  const width = Math.max(1, Math.round(window.width * config.kiTreeRegionRelWidth));
  const height = Math.max(1, Math.round(window.height * config.kiTreeRegionRelHeight));

  return { x, y, width, height };
}

function buildHeaderRegion(config: AppConfig, window: KiWindowInfo): VisionCaptureRegion {
  const x = Math.round(window.x + window.width * config.kiHeaderRegionRelX);
  const y = Math.round(window.y + window.height * config.kiHeaderRegionRelY);
  const width = Math.max(1, Math.round(window.width * config.kiHeaderRegionRelWidth));
  const height = Math.max(1, Math.round(window.height * config.kiHeaderRegionRelHeight));

  return { x, y, width, height };
}

function buildTableRegion(config: AppConfig, window: KiWindowInfo): VisionCaptureRegion {
  const x = Math.round(window.x + window.width * config.kiTableRegionRelX);
  const y = Math.round(window.y + window.height * config.kiTableRegionRelY);
  const width = Math.max(1, Math.round(window.width * config.kiTableRegionRelWidth));
  const height = Math.max(1, Math.round(window.height * config.kiTableRegionRelHeight));

  return { x, y, width, height };
}

function buildFullWindowRegion(window: KiWindowInfo): VisionCaptureRegion {
  return {
    x: window.x,
    y: window.y,
    width: Math.max(1, window.width),
    height: Math.max(1, window.height)
  };
}

async function captureRegionToPng(region: VisionCaptureRegion, outputPath: string): Promise<void> {
  const escapedPath = outputPath.replace(/'/g, "''");
  const script = `
Add-Type -AssemblyName System.Drawing
$bmp = New-Object System.Drawing.Bitmap ${region.width}, ${region.height}
$graphics = [System.Drawing.Graphics]::FromImage($bmp)
$graphics.CopyFromScreen(${region.x}, ${region.y}, 0, 0, $bmp.Size)
$bmp.Save('${escapedPath}', [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bmp.Dispose()
`;

  const { execFile } = await import("node:child_process");
  await new Promise<void>((resolve, reject) => {
    execFile(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
      { windowsHide: true },
      (error) => {
        if (error) {
          reject(error);
          return;
        }

        resolve();
      }
    );
  });
}

async function cropImageRegionToPng(
  sourcePath: string,
  region: VisionCaptureRegion,
  outputPath: string
): Promise<void> {
  const escapedSource = sourcePath.replace(/'/g, "''");
  const escapedOutput = outputPath.replace(/'/g, "''");
  const script = `
Add-Type -AssemblyName System.Drawing
$src = [System.Drawing.Bitmap]::FromFile('${escapedSource}')
$rect = New-Object System.Drawing.Rectangle ${region.x}, ${region.y}, ${region.width}, ${region.height}
$dst = $src.Clone($rect, $src.PixelFormat)
$dst.Save('${escapedOutput}', [System.Drawing.Imaging.ImageFormat]::Png)
$dst.Dispose()
$src.Dispose()
`;

  const { execFile } = await import("node:child_process");
  await new Promise<void>((resolve, reject) => {
    execFile(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
      { windowsHide: true },
      (error) => {
        if (error) {
          reject(error);
          return;
        }

        resolve();
      }
    );
  });
}

async function activateWindowByTitle(windowTitle: string): Promise<void> {
  const escapedTitle = windowTitle.replace(/'/g, "''");
  const verificationHint = windowTitle.includes("DVAG Online-System") ? "DVAG Online-System" : windowTitle;
  const escapedVerificationHint = verificationHint.replace(/'/g, "''");
  const script = `
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class WinApi {
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
}
"@

function Get-ForegroundTitle {
  $handle = [WinApi]::GetForegroundWindow()
  if ($handle -eq [IntPtr]::Zero) {
    return ""
  }

  $builder = New-Object System.Text.StringBuilder 512
  [void][WinApi]::GetWindowText($handle, $builder, $builder.Capacity)
  return $builder.ToString()
}

$wshell = New-Object -ComObject WScript.Shell
for ($attempt = 0; $attempt -lt 4; $attempt++) {
  if (-not $wshell.AppActivate('${escapedTitle}')) {
    Start-Sleep -Milliseconds 250
    continue
  }

  Start-Sleep -Milliseconds 350
  $foregroundTitle = Get-ForegroundTitle
  if ($foregroundTitle -like '*${escapedTitle}*' -or $foregroundTitle -like '*${escapedVerificationHint}*') {
    return
  }
}

throw 'Das KI-Hauptfenster konnte vor dem Screenshot nicht zuverlässig in den Vordergrund gebracht werden.'
`;

  const { execFile } = await import("node:child_process");
  await new Promise<void>((resolve, reject) => {
    execFile(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
      { windowsHide: true },
      (error) => {
        if (error) {
          reject(error);
          return;
        }

        resolve();
      }
    );
  });
}

async function captureKiTreeRegion(
  config: AppConfig,
  window: KiWindowInfo,
  options?: { keepArtifact?: boolean; fileStem?: string }
): Promise<VisionCaptureArtifact> {
  return captureKiRegion(config, window, buildTreeRegion(config, window), options?.fileStem ?? `ki-tree-${Date.now()}`);
}

async function captureKiHeaderRegion(
  config: AppConfig,
  window: KiWindowInfo,
  options?: { keepArtifact?: boolean; fileStem?: string }
): Promise<VisionCaptureArtifact> {
  return captureKiRegion(config, window, buildHeaderRegion(config, window), options?.fileStem ?? `ki-header-${Date.now()}`);
}

async function captureKiTableRegion(
  config: AppConfig,
  window: KiWindowInfo,
  options?: { keepArtifact?: boolean; fileStem?: string }
): Promise<VisionCaptureArtifact> {
  return captureKiRegion(config, window, buildTableRegion(config, window), options?.fileStem ?? `ki-table-${Date.now()}`);
}

async function captureKiFullWindowRegion(
  config: AppConfig,
  window: KiWindowInfo,
  options?: { keepArtifact?: boolean; fileStem?: string }
): Promise<VisionCaptureArtifact> {
  return captureKiRegion(config, window, buildFullWindowRegion(window), options?.fileStem ?? `ki-full-${Date.now()}`);
}

async function captureKiRegion(
  config: AppConfig,
  window: KiWindowInfo,
  region: VisionCaptureRegion,
  fileStem: string
): Promise<VisionCaptureArtifact> {
  await mkdir(config.visionDirectory, { recursive: true });
  const createdAt = new Date().toISOString();
  const imagePath = path.join(config.visionDirectory, `${fileStem}.png`);

  await activateWindowByTitle(window.title);
  await captureRegionToPng(region, imagePath);

  return {
    imagePath,
    region,
    createdAt
  };
}

async function disposeVisionArtifact(config: AppConfig, artifact: VisionCaptureArtifact): Promise<void> {
  if (config.visionKeepDebugArtifacts) {
    return;
  }

  await rm(artifact.imagePath, { force: true });
}

export {
  buildFullWindowRegion,
  buildHeaderRegion,
  buildTableRegion,
  buildTreeRegion,
  captureKiFullWindowRegion,
  captureKiHeaderRegion,
  captureKiTableRegion,
  captureKiTreeRegion,
  cropImageRegionToPng,
  disposeVisionArtifact
};
