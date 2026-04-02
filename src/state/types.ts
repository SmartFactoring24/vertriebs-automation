export type SalesRecord = {
  businessId: string;
  partnerStage: string;
  partnerName: string;
  productName: string;
  status: string;
  unitsValue: number;
  totalUnits: number;
  evaluationPeriod: string;
  businessScope: string;
  sourcePage: string;
  submittedAt: string;
  updatedAt: string;
  source: "KI";
};

export type ChangeType =
  | "created"
  | "status_changed"
  | "units_value_changed"
  | "updated"
  | "removed";

export type SalesChange = {
  eventId: string;
  type: ChangeType;
  record: SalesRecord;
  previousRecord?: SalesRecord;
  detectedAt: string;
};

export type PersistedState = {
  records: SalesRecord[];
  sentEventIds: string[];
  updatedAt: string | null;
};
