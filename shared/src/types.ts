/**
 * Shared type definitions for Excel and Google Sheets add-ons
 */

export interface ApiResponse<T = any> {
  data: T;
  status: 'success' | 'error';
  message?: string;
}

export interface DateRange {
  startDate: string;
  endDate: string;
}

export interface DataPoint {
  date: string;
  value: number;
  category?: string;
}

export interface Config {
  baseRoot: string;
  token?: string;
}

export interface ApiPayload {
  start_date: string | null;
  end_date: string | null;
  great_or_equal: boolean | null;
  less_or_equal: boolean | null;
  columns: string[] | null;
  unique_identifier_list: string[] | null;
}

export interface StorageAdapter {
  getItem(key: string): Promise<string | null> | string | null;
  setItem(key: string, value: string): Promise<void> | void;
}

export interface HttpAdapter {
  fetch(url: string, options: any): Promise<any> | any;
}
