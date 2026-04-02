import fs from "node:fs/promises";
import path from "node:path";
import { createObjectCsvWriter } from "csv-writer";
import type { SalesChange, SalesRecord } from "../state/types.js";

async function ensureUtf8Bom(filePath: string): Promise<void> {
  const content = await fs.readFile(filePath, "utf8");
  if (content.startsWith("\uFEFF")) {
    return;
  }

  await fs.writeFile(filePath, `\uFEFF${content}`, "utf8");
}

export async function exportSnapshotToCsv(records: SalesRecord[], exportDirectory: string): Promise<string> {
  await fs.mkdir(exportDirectory, { recursive: true });
  const filePath = path.join(exportDirectory, "snapshot-current.csv");

  const writer = createObjectCsvWriter({
    path: filePath,
    header: [
      { id: "businessId", title: "businessId" },
      { id: "partnerStage", title: "partnerStage" },
      { id: "partnerName", title: "partnerName" },
      { id: "productName", title: "productName" },
      { id: "status", title: "status" },
      { id: "unitsValue", title: "unitsValue" },
      { id: "totalUnits", title: "totalUnits" },
      { id: "evaluationPeriod", title: "evaluationPeriod" },
      { id: "businessScope", title: "businessScope" },
      { id: "sourcePage", title: "sourcePage" },
      { id: "submittedAt", title: "submittedAt" },
      { id: "updatedAt", title: "updatedAt" },
      { id: "source", title: "source" }
    ]
  });

  await writer.writeRecords(records);
  await ensureUtf8Bom(filePath);
  return filePath;
}

export async function exportChangesToCsv(changes: SalesChange[], exportDirectory: string): Promise<string | null> {
  if (changes.length === 0) {
    return null;
  }

  await fs.mkdir(exportDirectory, { recursive: true });
  const filePath = path.join(exportDirectory, "changes-current.csv");

  const writer = createObjectCsvWriter({
    path: filePath,
    header: [
      { id: "eventId", title: "eventId" },
      { id: "type", title: "type" },
      { id: "detectedAt", title: "detectedAt" },
      { id: "businessId", title: "businessId" },
      { id: "partnerStage", title: "partnerStage" },
      { id: "partnerName", title: "partnerName" },
      { id: "productName", title: "productName" },
      { id: "status", title: "status" },
      { id: "unitsValue", title: "unitsValue" },
      { id: "totalUnits", title: "totalUnits" },
      { id: "evaluationPeriod", title: "evaluationPeriod" },
      { id: "businessScope", title: "businessScope" },
      { id: "updatedAt", title: "updatedAt" }
    ]
  });

  await writer.writeRecords(
    changes.map((change) => ({
      eventId: change.eventId,
      type: change.type,
      detectedAt: change.detectedAt,
      businessId: change.record.businessId,
      partnerStage: change.record.partnerStage,
      partnerName: change.record.partnerName,
      productName: change.record.productName,
      status: change.record.status,
      unitsValue: change.record.unitsValue,
      totalUnits: change.record.totalUnits,
      evaluationPeriod: change.record.evaluationPeriod,
      businessScope: change.record.businessScope,
      updatedAt: change.record.updatedAt
    }))
  );

  await ensureUtf8Bom(filePath);
  return filePath;
}
