export type SalesRecord = {
  businessId: string;
  customerName: string;
  productName: string;
  status: string;
  salesValue: number;
  submittedAt: string;
  updatedAt: string;
  source: "KI";
};

export type ChangeType =
  | "created"
  | "status_changed"
  | "sales_value_changed"
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
