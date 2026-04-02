import type { AppConfig } from "../config.js";
import type { SalesRecord } from "../state/types.js";
import { readTableCaptureWithOcr } from "../vision/ocr.js";
import { captureCurrentKiTableRegion, ensureKiDesktopReady, performKiLogin } from "./ki-desktop.js";

export async function fetchSalesRecords(config: AppConfig): Promise<SalesRecord[]> {
  const state = await ensureKiDesktopReady(config);

  if (state.stage === "login") {
    await performKiLogin(config);
  }

  return collectSalesRecords(config);
}

async function collectSalesRecords(config: AppConfig): Promise<SalesRecord[]> {
  const artifact = await captureCurrentKiTableRegion(config);
  const ocr = await readTableCaptureWithOcr(config, artifact.imagePath);
  const now = new Date().toISOString();
  const evaluationPeriod = extractValue(ocr.normalizedLines, /^Auswertungszeitraum\s+(.+)$/i) ?? "";
  const businessScope = extractValue(ocr.normalizedLines, /^Gewertetes Geschäft\s+(.+)$/i) ?? "";
  const sourcePage = extractValue(
    ocr.normalizedLines,
    /^VB\d+\s+Eingereichtes Geschäft:\s+(.+?)\s+Erstellungsdatum:/i
  ) ?? "Einheiten nach Sparten der Gruppe";

  return ocr.parsedRows.flatMap((row) =>
    row.cells.map((cell) => ({
      businessId: buildBusinessId(row.stage, row.name, cell.column),
      partnerStage: row.stage,
      partnerName: row.name,
      productName: cell.column,
      status: "eingereicht",
      unitsValue: parseGermanNumber(cell.value),
      totalUnits: parseGermanNumber(row.sum ?? cell.value),
      evaluationPeriod,
      businessScope,
      sourcePage,
      submittedAt: now,
      updatedAt: now,
      source: "KI" as const
    }))
  );
}

function parseGermanNumber(value: string): number {
  const normalized = value.replace(/\./g, "").replace(",", ".");
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function buildBusinessId(stage: string, name: string, column: string): string {
  return `${stage}:${name}:${column}`
    .normalize("NFKD")
    .replace(/[^\w:.-]+/g, "_");
}

function extractValue(lines: string[], pattern: RegExp): string | null {
  const line = lines.find((entry) => pattern.test(entry));
  if (!line) {
    return null;
  }

  const match = line.match(pattern);
  return match?.[1]?.trim() ?? null;
}
