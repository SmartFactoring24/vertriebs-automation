import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

type TemplateSearchRegion = {
  x: number;
  y: number;
  width: number;
  height: number;
};

type TemplateMatchResult = {
  found: boolean;
  x: number;
  y: number;
  width: number;
  height: number;
  score: number;
};

function escapeSingleQuotedPowerShell(value: string): string {
  return value.replace(/'/g, "''");
}

async function matchTemplateInImage(options: {
  sourcePath: string;
  templatePath: string;
  searchRegion?: TemplateSearchRegion;
  sampleCols?: number;
  sampleRows?: number;
  threshold?: number;
}): Promise<TemplateMatchResult> {
  const escapedSource = escapeSingleQuotedPowerShell(options.sourcePath);
  const escapedTemplate = escapeSingleQuotedPowerShell(options.templatePath);
  const sampleCols = options.sampleCols ?? 10;
  const sampleRows = options.sampleRows ?? 4;
  const threshold = options.threshold ?? 35;
  const region = options.searchRegion ?? { x: 0, y: 0, width: -1, height: -1 };

  const script = `
Add-Type -AssemblyName System.Drawing
$source = [System.Drawing.Bitmap]::FromFile('${escapedSource}')
$template = [System.Drawing.Bitmap]::FromFile('${escapedTemplate}')
$regionX = ${region.x}
$regionY = ${region.y}
$regionWidth = ${region.width}
$regionHeight = ${region.height}
if ($regionWidth -lt 0) { $regionWidth = $source.Width - $regionX }
if ($regionHeight -lt 0) { $regionHeight = $source.Height - $regionY }
$maxX = [Math]::Min($source.Width - $template.Width, $regionX + $regionWidth - $template.Width)
$maxY = [Math]::Min($source.Height - $template.Height, $regionY + $regionHeight - $template.Height)
$sampleCols = ${sampleCols}
$sampleRows = ${sampleRows}
$samplePoints = @()
for ($cx = 0; $cx -lt $sampleCols; $cx++) {
  for ($cy = 0; $cy -lt $sampleRows; $cy++) {
    $tx = [Math]::Min($template.Width - 1, [Math]::Round((($cx + 0.5) / $sampleCols) * ($template.Width - 1)))
    $ty = [Math]::Min($template.Height - 1, [Math]::Round((($cy + 0.5) / $sampleRows) * ($template.Height - 1)))
    $samplePoints += [PSCustomObject]@{ X = $tx; Y = $ty }
  }
}
$bestScore = [double]::PositiveInfinity
$bestX = -1
$bestY = -1
for ($x = $regionX; $x -le $maxX; $x++) {
  for ($y = $regionY; $y -le $maxY; $y++) {
    $score = 0.0
    foreach ($point in $samplePoints) {
      $sp = $source.GetPixel($x + $point.X, $y + $point.Y)
      $tp = $template.GetPixel($point.X, $point.Y)
      $score += [Math]::Abs($sp.R - $tp.R)
      $score += [Math]::Abs($sp.G - $tp.G)
      $score += [Math]::Abs($sp.B - $tp.B)
    }
    $score = $score / ($samplePoints.Count * 3.0)
    if ($score -lt $bestScore) {
      $bestScore = $score
      $bestX = $x
      $bestY = $y
    }
  }
}
$source.Dispose()
$template.Dispose()
[PSCustomObject]@{
  found = ($bestScore -le ${threshold})
  x = $bestX
  y = $bestY
  width = 0
  height = 0
  score = [Math]::Round($bestScore, 2)
} | ConvertTo-Json -Compress
`;

  const { stdout } = await execFileAsync(
    "powershell.exe",
    ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
    { windowsHide: true, maxBuffer: 10 * 1024 * 1024 }
  );

  const raw = stdout.trim();
  const parsed = JSON.parse(raw) as TemplateMatchResult;
  return parsed;
}

export { matchTemplateInImage };
export type { TemplateMatchResult, TemplateSearchRegion };
