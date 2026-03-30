import fs from "node:fs/promises";
import path from "node:path";
import { createObjectCsvWriter } from "csv-writer";
import type { SalesChange, SalesRecord } from "../state/types.js";

export async function exportSnapshotToCsv(records: SalesRecord[], exportDirectory: string): Promise<string> {
  await fs.mkdir(exportDirectory, { recursive: true });
  const filePath = path.join(exportDirectory, `snapshot-${Date.now()}.csv`);

  const writer = createObjectCsvWriter({
    path: filePath,
    header: [
      { id: "businessId", title: "businessId" },
      { id: "customerName", title: "customerName" },
      { id: "productName", title: "productName" },
      { id: "status", title: "status" },
      { id: "salesValue", title: "salesValue" },
      { id: "submittedAt", title: "submittedAt" },
      { id: "updatedAt", title: "updatedAt" },
      { id: "source", title: "source" }
    ]
  });

  await writer.writeRecords(records);
  return filePath;
}

export async function exportChangesToCsv(changes: SalesChange[], exportDirectory: string): Promise<string | null> {
  if (changes.length === 0) {
    return null;
  }

  await fs.mkdir(exportDirectory, { recursive: true });
  const filePath = path.join(exportDirectory, `changes-${Date.now()}.csv`);

  const writer = createObjectCsvWriter({
    path: filePath,
    header: [
      { id: "eventId", title: "eventId" },
      { id: "type", title: "type" },
      { id: "detectedAt", title: "detectedAt" },
      { id: "businessId", title: "businessId" },
      { id: "customerName", title: "customerName" },
      { id: "productName", title: "productName" },
      { id: "status", title: "status" },
      { id: "salesValue", title: "salesValue" },
      { id: "updatedAt", title: "updatedAt" }
    ]
  });

  await writer.writeRecords(
    changes.map((change) => ({
      eventId: change.eventId,
      type: change.type,
      detectedAt: change.detectedAt,
      businessId: change.record.businessId,
      customerName: change.record.customerName,
      productName: change.record.productName,
      status: change.record.status,
      salesValue: change.record.salesValue,
      updatedAt: change.record.updatedAt
    }))
  );

  return filePath;
}
