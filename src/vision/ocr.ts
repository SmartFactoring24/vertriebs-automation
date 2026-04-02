import { execFile } from "node:child_process";
import { mkdir, rm, access } from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";
import type { AppConfig } from "../config.js";
import type { VisionCaptureRegion } from "./types.js";
import { cropImageRegionToPng } from "./tree-capture.js";

const execFileAsync = promisify(execFile);

type OcrTableReadResult = {
  imagePath: string;
  preprocessedImagePath: string;
  rawText: string;
  normalizedLines: string[];
  parsedRows: ParsedTableRow[];
};

type ParsedTableCell = {
  column: string;
  value: string;
};

type ParsedTableRow = {
  stage: string;
  name: string;
  cells: ParsedTableCell[];
  sum?: string;
};

type ParsedTableRowWithRegion = ParsedTableRow & {
  nameRegion?: VisionCaptureRegion;
};

type TesseractTsvWord = {
  text: string;
  left: number;
  top: number;
  width: number;
  height: number;
  conf: number;
};

async function preprocessImageForOcr(sourcePath: string, outputPath: string): Promise<void> {
  const escapedSource = sourcePath.replace(/'/g, "''");
  const escapedOutput = outputPath.replace(/'/g, "''");
  const script = `
Add-Type -AssemblyName System.Drawing
$src = [System.Drawing.Bitmap]::FromFile('${escapedSource}')
$dst = New-Object System.Drawing.Bitmap ($src.Width * 3), ($src.Height * 3)
$graphics = [System.Drawing.Graphics]::FromImage($dst)
$graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$graphics.DrawImage($src, 0, 0, $dst.Width, $dst.Height)
$dst.Save('${escapedOutput}', [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$dst.Dispose()
$src.Dispose()
`;

  await execFileAsync(
    "powershell.exe",
    ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
    { windowsHide: true }
  );
}

async function runTesseractText(
  config: AppConfig,
  imagePath: string,
  options?: { psm?: string; lang?: string; extraArgs?: string[] }
): Promise<string> {
  await access(config.tesseractPath);
  const args = [
    imagePath,
    "stdout",
    "--tessdata-dir",
    config.tesseractDataDirectory,
    "--psm",
    options?.psm ?? "6",
    "-l",
    options?.lang ?? "deu+eng",
    ...(options?.extraArgs ?? [])
  ];
  const { stdout } = await execFileAsync(
    config.tesseractPath,
    args,
    {
      windowsHide: true,
      maxBuffer: 10 * 1024 * 1024
    }
  );

  return stdout;
}

async function runTesseractTsv(config: AppConfig, imagePath: string): Promise<string> {
  await access(config.tesseractPath);
  const { stdout } = await execFileAsync(
    config.tesseractPath,
    [imagePath, "stdout", "--psm", "6", "-l", "deu+eng", "tsv"],
    { windowsHide: true, maxBuffer: 10 * 1024 * 1024 }
  );

  return stdout;
}

function normalizeOcrLines(rawText: string): string[] {
  return rawText
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter((line) => line.length > 0);
}

function parseTesseractTsv(tsv: string): TesseractTsvWord[] {
  const lines = tsv.split(/\r?\n/).filter((line) => line.trim().length > 0);
  if (lines.length <= 1) {
    return [];
  }

  return lines
    .slice(1)
    .map((line) => line.split("\t"))
    .filter((parts) => parts.length >= 12)
    .map((parts) => ({
      text: parts[11]?.trim() ?? "",
      left: Number.parseInt(parts[6] ?? "0", 10),
      top: Number.parseInt(parts[7] ?? "0", 10),
      width: Number.parseInt(parts[8] ?? "0", 10),
      height: Number.parseInt(parts[9] ?? "0", 10),
      conf: Number.parseFloat(parts[10] ?? "-1")
    }))
    .filter((word) => word.text.length > 0 && Number.isFinite(word.left) && Number.isFinite(word.top));
}

function clusterRows(words: TesseractTsvWord[]): TesseractTsvWord[][] {
  const sorted = [...words].sort((a, b) => a.top - b.top || a.left - b.left);
  const rows: TesseractTsvWord[][] = [];

  for (const word of sorted) {
    const existing = rows.find((row) => Math.abs(row[0].top - word.top) <= 18);
    if (existing) {
      existing.push(word);
      existing.sort((a, b) => a.left - b.left);
    } else {
      rows.push([word]);
    }
  }

  return rows;
}

function isNumericToken(value: string): boolean {
  return /^-?\d+(?:[.,]\d+)?$/.test(value);
}

function normalizeNumericToken(value: string): string {
  if (/^0\d$/.test(value)) {
    return `0,${value.slice(1)}`;
  }

  return value;
}

function looksLikePersonName(value: string): boolean {
  return /^[A-Za-zÄÖÜäöüß.-]+,\s*[A-Za-zÄÖÜäöüß.-]+$/.test(value);
}

function detectHeaderColumns(rows: TesseractTsvWord[][]): Array<{ name: string; left: number }> {
  const headerRow = rows.find((row) => row.some((word) => /Stufe|Name|Leben|Sach|Bausparen|Invest|Banken|Sonstige|Summe/i.test(word.text)));
  if (!headerRow) {
    return [];
  }

  const headerNames = ["Stufe", "Name", "Leben", "Sach", "KV", "Bausparen", "Invest", "Banken", "Sonstige", "Summe"];
  return headerRow
    .map((word) => {
      const normalized = word.text.replace(/[+v]/g, "");
      const matched = headerNames.find((name) => normalized.localeCompare(name, "de", { sensitivity: "base" }) === 0);
      return matched ? { name: matched, left: word.left } : null;
    })
    .filter((entry): entry is { name: string; left: number } => entry !== null)
    .sort((a, b) => a.left - b.left);
}

function parseRowsFromTsv(tsv: string): ParsedTableRowWithRegion[] {
  const words = parseTesseractTsv(tsv);
  const rows = clusterRows(words);
  const columns = detectHeaderColumns(rows);
  if (columns.length === 0) {
    return [];
  }

  const dataRows = rows.filter((row) => {
    const first = row[0]?.text ?? "";
    return /^(AL|VBAS|VM|HGS|RGS|VB|GS|RD\d?)$/i.test(first);
  });

  return dataRows.map((row) => {
    const stage = row[0]?.text ?? "";
    const remaining = row.slice(1);
    const numericStartIndex = remaining.findIndex((word) => isNumericToken(word.text));
    const nameWords = numericStartIndex === -1 ? remaining : remaining.slice(0, numericStartIndex);
    const valueWords = numericStartIndex === -1 ? [] : remaining.slice(numericStartIndex);
    const name = nameWords.map((word) => word.text).join(" ").replace(/\s+,/g, ",").trim();
    const nameRegion = nameWords.length > 0
      ? {
          x: Math.max(0, Math.min(...nameWords.map((word) => word.left)) - 10),
          y: Math.max(0, Math.min(...nameWords.map((word) => word.top)) - 6),
          width: Math.max(...nameWords.map((word) => word.left + word.width)) - Math.min(...nameWords.map((word) => word.left)) + 20,
          height: Math.max(...nameWords.map((word) => word.top + word.height)) - Math.min(...nameWords.map((word) => word.top)) + 12
        }
      : undefined;

    const cells: ParsedTableCell[] = [];
    let sum: string | undefined;

    for (const word of valueWords) {
      const nearestColumn = columns.reduce<{ name: string; left: number } | null>((best, column) => {
        if (!best) {
          return column;
        }

        return Math.abs(column.left - word.left) < Math.abs(best.left - word.left) ? column : best;
      }, null);

      if (!nearestColumn || nearestColumn.name === "Stufe" || nearestColumn.name === "Name") {
        continue;
      }

      if (nearestColumn.name === "Summe") {
        sum = normalizeNumericToken(word.text);
      } else {
        cells.push({
          column: nearestColumn.name,
          value: normalizeNumericToken(word.text)
        });
      }
    }

    return {
      stage,
      name,
      cells,
      sum,
      nameRegion
    };
  });
}

async function refineRowNames(
  config: AppConfig,
  imagePath: string,
  rows: ParsedTableRowWithRegion[]
): Promise<ParsedTableRow[]> {
  const refinedRows: ParsedTableRow[] = [];

  for (let index = 0; index < rows.length; index++) {
    const row = rows[index];
    let refinedName = row.name;

    if (row.nameRegion) {
      const cropPath = path.join(config.visionDirectory, `ki-name-ocr-${Date.now()}-${index}.png`);
      await cropImageRegionToPng(imagePath, row.nameRegion, cropPath);
      try {
        const refinedText = await runTesseractText(config, cropPath, {
          psm: "7",
          lang: "deu",
          extraArgs: ["-c", "tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzÄÖÜäöüß,.- "]
        });
        const candidate = refinedText.replace(/\s+/g, " ").trim();
        if (
          candidate.length > 0 &&
          looksLikePersonName(candidate) &&
          candidate.length >= row.name.length - 1
        ) {
          refinedName = candidate;
        }
      } finally {
        if (!config.visionKeepDebugArtifacts) {
          await rm(cropPath, { force: true });
        }
      }
    }

    refinedRows.push({
      stage: row.stage,
      name: refinedName,
      cells: row.cells,
      sum: row.sum
    });
  }

  return refinedRows;
}

function extractNameFromNormalizedLines(lines: string[], stage: string): string | null {
  const line = lines.find((entry) => entry.startsWith(`${stage} `));
  if (!line) {
    return null;
  }

  const withoutStage = line.slice(stage.length).trim();
  const match = withoutStage.match(/^(.*?)(?=\s-?\d+(?:[.,]\d+)?(?:\s|$))/);
  if (!match) {
    return null;
  }

  const candidate = match[1]?.trim() ?? "";
  return looksLikePersonName(candidate) ? candidate : null;
}

function mergeNamesFromRawLines(lines: string[], rows: ParsedTableRow[]): ParsedTableRow[] {
  return rows.map((row) => ({
    ...row,
    name: extractNameFromNormalizedLines(lines, row.stage) ?? row.name
  }));
}

async function readTableCaptureWithOcr(config: AppConfig, imagePath: string): Promise<OcrTableReadResult> {
  await mkdir(config.visionDirectory, { recursive: true });
  const preprocessedImagePath = path.join(config.visionDirectory, `ki-table-ocr-${Date.now()}.png`);
  await preprocessImageForOcr(imagePath, preprocessedImagePath);
  const rawText = await runTesseractText(config, preprocessedImagePath);
  const normalizedLines = normalizeOcrLines(rawText);
  const tsv = await runTesseractTsv(config, preprocessedImagePath);
  const parsedRows = mergeNamesFromRawLines(
    normalizedLines,
    await refineRowNames(config, preprocessedImagePath, parseRowsFromTsv(tsv))
  );

  if (!config.visionKeepDebugArtifacts) {
    await rm(preprocessedImagePath, { force: true });
  }

  return {
    imagePath,
    preprocessedImagePath,
    rawText,
    normalizedLines,
    parsedRows
  };
}

export { readTableCaptureWithOcr };
export type { OcrTableReadResult };
