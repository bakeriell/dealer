export interface DealerData {
  id: string;
  [key: string]: string | number | boolean;
}

export interface ColumnDefinition {
  key: string;
  label: string;
  type: 'text' | 'number' | 'boolean' | 'date';
}

export interface ChartDataPoint {
  name: string;
  value: number;
}

export enum AppView {
  DASHBOARD = 'DASHBOARD',
  DATA_GRID = 'DATA_GRID',
}

export type LeadFilterType = 'ALL' | 'CONTACTED' | 'NOT_CONTACTED';