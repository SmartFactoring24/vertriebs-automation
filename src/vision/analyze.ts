import { execFile } from "node:child_process";
import { promisify } from "node:util";
import type { AppConfig } from "../config.js";
import type { VisionCaptureArtifact } from "./types.js";

const execFileAsync = promisify(execFile);

type VisionRegionProbe = {
  name: string;
  relX: number;
  relY: number;
  relWidth: number;
  relHeight: number;
};

type VisionRegionAnalysis = {
  name: string;
  averageR: number;
  averageG: number;
  averageB: number;
  averageLuma: number;
  colorSpread: number;
};

type HeaderVisionAssessment = {
  imagePath: string;
  probes: VisionRegionAnalysis[];
  looksLikeGruppenAkte: boolean;
};

type TreeVisionAssessment = {
  imagePath: string;
  probes: VisionRegionAnalysis[];
  looksLikeSubmittedUnitsSelection: boolean;
};

function escapeSingleQuotedPowerShell(value: string): string {
  return value.replace(/'/g, "''");
}

async function analyzeImageRegions(
  artifact: VisionCaptureArtifact,
  probes: VisionRegionProbe[]
): Promise<VisionRegionAnalysis[]> {
  const escapedPath = escapeSingleQuotedPowerShell(artifact.imagePath);
  const probeObjects = probes
    .map((probe) => {
      const x = Math.round(probe.relX * artifact.region.width);
      const y = Math.round(probe.relY * artifact.region.height);
      const width = Math.max(1, Math.round(probe.relWidth * artifact.region.width));
      const height = Math.max(1, Math.round(probe.relHeight * artifact.region.height));
      return `@{ Name='${escapeSingleQuotedPowerShell(probe.name)}'; X=${x}; Y=${y}; Width=${width}; Height=${height} }`;
    })
    .join(",\n");

  const script = `
Add-Type -AssemblyName System.Drawing
$image = [System.Drawing.Bitmap]::FromFile('${escapedPath}')
$probes = @(
${probeObjects}
)
$results = @()
foreach ($probe in $probes) {
  $sumR = 0.0
  $sumG = 0.0
  $sumB = 0.0
  $count = 0.0
  $minLuma = 255.0
  $maxLuma = 0.0
  for ($x = $probe.X; $x -lt [Math]::Min($image.Width, $probe.X + $probe.Width); $x++) {
    for ($y = $probe.Y; $y -lt [Math]::Min($image.Height, $probe.Y + $probe.Height); $y++) {
      $pixel = $image.GetPixel($x, $y)
      $sumR += $pixel.R
      $sumG += $pixel.G
      $sumB += $pixel.B
      $luma = (($pixel.R * 0.299) + ($pixel.G * 0.587) + ($pixel.B * 0.114))
      if ($luma -lt $minLuma) { $minLuma = $luma }
      if ($luma -gt $maxLuma) { $maxLuma = $luma }
      $count += 1.0
    }
  }
  if ($count -eq 0) { continue }
  $avgR = $sumR / $count
  $avgG = $sumG / $count
  $avgB = $sumB / $count
  $avgLuma = (($avgR * 0.299) + ($avgG * 0.587) + ($avgB * 0.114))
  $channelSpread = [Math]::Max($avgR, [Math]::Max($avgG, $avgB)) - [Math]::Min($avgR, [Math]::Min($avgG, $avgB))
  $results += [PSCustomObject]@{
    name = $probe.Name
    averageR = [Math]::Round($avgR, 2)
    averageG = [Math]::Round($avgG, 2)
    averageB = [Math]::Round($avgB, 2)
    averageLuma = [Math]::Round($avgLuma, 2)
    colorSpread = [Math]::Round([Math]::Max($channelSpread, ($maxLuma - $minLuma)), 2)
  }
}
$image.Dispose()
$results | ConvertTo-Json -Compress
`;

  const { stdout } = await execFileAsync("powershell.exe", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    script
  ], { windowsHide: true });

  const raw = stdout.trim();

  if (!raw) {
    return [];
  }

  const parsed = JSON.parse(raw) as VisionRegionAnalysis[] | VisionRegionAnalysis;
  return Array.isArray(parsed) ? parsed : [parsed];
}

async function analyzeHeaderCapture(_config: AppConfig, artifact: VisionCaptureArtifact): Promise<HeaderVisionAssessment> {
  const probes = await analyzeImageRegions(artifact, [
    { name: "gruppenAkteToolbar", relX: 0.03, relY: 0.10, relWidth: 0.16, relHeight: 0.30 },
    { name: "contentHeader", relX: 0.18, relY: 0.48, relWidth: 0.34, relHeight: 0.24 },
    { name: "contentTitleArea", relX: 0.21, relY: 0.52, relWidth: 0.24, relHeight: 0.18 }
  ]);

  const toolbar = probes.find((probe) => probe.name === "gruppenAkteToolbar");
  const contentHeader = probes.find((probe) => probe.name === "contentHeader");
  const contentTitleArea = probes.find((probe) => probe.name === "contentTitleArea");

  const looksLikeGruppenAkte =
    (toolbar ? toolbar.averageLuma < 235 || toolbar.colorSpread > 25 : false) &&
    (
      (contentHeader ? contentHeader.colorSpread > 18 : false) ||
      (contentTitleArea ? contentTitleArea.averageLuma < 225 : false)
    );

  return {
    imagePath: artifact.imagePath,
    probes,
    looksLikeGruppenAkte
  };
}

async function analyzeTreeCapture(_config: AppConfig, artifact: VisionCaptureArtifact): Promise<TreeVisionAssessment> {
  const probes = await analyzeImageRegions(artifact, [
    { name: "selectedRowBackground", relX: 0.22, relY: 0.39, relWidth: 0.40, relHeight: 0.04 },
    { name: "selectedRowText", relX: 0.26, relY: 0.39, relWidth: 0.28, relHeight: 0.04 },
    { name: "neighborRowText", relX: 0.26, relY: 0.43, relWidth: 0.28, relHeight: 0.04 }
  ]);

  const selectedRowBackground = probes.find((probe) => probe.name === "selectedRowBackground");
  const selectedRowText = probes.find((probe) => probe.name === "selectedRowText");

  const looksLikeSubmittedUnitsSelection =
    (selectedRowBackground ? selectedRowBackground.averageR > 175 && selectedRowBackground.averageG > 150 : false) &&
    (selectedRowBackground ? selectedRowBackground.averageB < 140 : false) &&
    (selectedRowText ? selectedRowText.averageLuma < 210 : true);

  return {
    imagePath: artifact.imagePath,
    probes,
    looksLikeSubmittedUnitsSelection
  };
}

export { analyzeHeaderCapture, analyzeTreeCapture };
export type { HeaderVisionAssessment, TreeVisionAssessment, VisionRegionAnalysis };
