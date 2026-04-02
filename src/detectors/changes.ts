import type { SalesChange, SalesRecord } from "../state/types.js";

function buildEventId(type: SalesChange["type"], record: SalesRecord): string {
  return `${type}:${record.businessId}:${record.updatedAt}`;
}

export function detectChanges(previous: SalesRecord[], current: SalesRecord[]): SalesChange[] {
  const previousById = new Map(previous.map((record) => [record.businessId, record]));
  const currentById = new Map(current.map((record) => [record.businessId, record]));
  const detectedAt = new Date().toISOString();
  const changes: SalesChange[] = [];

  for (const record of current) {
    const previousRecord = previousById.get(record.businessId);

    if (!previousRecord) {
      changes.push({
        eventId: buildEventId("created", record),
        type: "created",
        record,
        detectedAt
      });
      continue;
    }

    if (previousRecord.status !== record.status) {
      changes.push({
        eventId: buildEventId("status_changed", record),
        type: "status_changed",
        record,
        previousRecord,
        detectedAt
      });
    }

    if (previousRecord.unitsValue !== record.unitsValue) {
      changes.push({
        eventId: buildEventId("units_value_changed", record),
        type: "units_value_changed",
        record,
        previousRecord,
        detectedAt
      });
    }

  }

  for (const record of previous) {
    if (!currentById.has(record.businessId)) {
      changes.push({
        eventId: buildEventId("removed", record),
        type: "removed",
        record,
        previousRecord: record,
        detectedAt
      });
    }
  }

  return changes;
}
