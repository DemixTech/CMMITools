/**
 * Shared TypeScript interfaces for CAS Form Populator
 */

export interface FieldMapping {
  Row: number;
  Sheet: string;
  FieldLabel: string;
  CAS_Page: string;
  CAS_Selector: string;
  CAS_FieldName: string;
  CAS_Type: string;
  Notes: string;
}

export interface ExcelData {
  [sheet: string]: {
    [row: number]: string;
  };
}

export interface PopulateLog {
  timestamp: string;
  page: string;
  fieldsAttempted: number;
  fieldsSuccessful: number;
  fieldsFailed: number;
  details: {
    field: string;
    selector: string;
    value: string;
    status: 'success' | 'failed' | 'skipped';
    error?: string;
  }[];
  userFeedback?: string;
}
