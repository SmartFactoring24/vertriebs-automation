import fs from "node:fs/promises";
import path from "node:path";
import * as XLSX from "xlsx";
import type { SalesRecord } from "../state/types.js";

export async function exportSnapshotToXlsx(records: SalesRecord[], exportDirectory: string): Promise<string> {
  await fs.mkdir(exportDirectory, { recursive: true });

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(records);
  XLSX.utils.book_append_sheet(workbook, worksheet, "SalesSnapshot");

  const filePath = path.join(exportDirectory, `snapshot-${Date.now()}.xlsx`);
  XLSX.writeFile(workbook, filePath);
  return filePath;
}
