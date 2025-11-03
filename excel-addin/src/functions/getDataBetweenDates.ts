/* eslint-disable no-var */
/* global OfficeRuntime */

import {
  getDataBetweenDatesFromApi,
  setBaseRoot,
  StorageAdapter,
  HttpAdapter
} from '@excel-plugin/shared';
import { config } from '../config';

// Excel-specific storage adapter
const excelStorageAdapter: StorageAdapter = {
  async getItem(key: string): Promise<string | null> {
    try {
      const stored = await OfficeRuntime.storage.getItem(key);
      return typeof stored === "string" ? stored : null;
    } catch {
      return null;
    }
  },
  async setItem(key: string, value: string): Promise<void> {
    await OfficeRuntime.storage.setItem(key, value);
  }
};

// Excel-specific HTTP adapter
const excelHttpAdapter: HttpAdapter = {
  async fetch(url: string, options: any): Promise<any> {
    return fetch(url, options);
  }
};

// ---------- Optional: allow setting BASE_ROOT from a formula ----------
/**
 * Stores the base API root URL for subsequent calls.
 * Usage: =SET_BASE_ROOT("https://api.example.com")
 * @customfunction
 * @param {string} url Base API root (no trailing slash required).
 * @returns {string} "OK" on success.
 */
export async function SET_BASE_ROOT(url: string): Promise<string> {
  return await setBaseRoot(url, excelStorageAdapter);
}

// ---------- Main custom function ----------
/**
 * Fetches rows from your Django API between two dates and spills them as a table.
 *
 * Anchor at top-left (e.g., C5). Returned 2‑D array spills down/right.
 *
 * @customfunction
 * @param {string} table_name
 * @param {Date|string|number} [start_date]
 * @param {Date|string|number} [end_date]
 * @param {boolean|string} [great_or_equal]
 * @param {boolean|string} [less_or_equal]
 * @param {string[][]} [columns]
 * @param {string[][]} [unique_identifier_list]
 * @returns {any[][]}
 */
export async function GET_DATA_BETWEEN_DATES_FROM_API(
  table_name: string,
  start_date?: any,
  end_date?: any,
  great_or_equal?: any,
  less_or_equal?: any,
  columns?: (string | number | boolean)[][],
  unique_identifier_list?: (string | number | boolean)[][]
): Promise<(string | number | boolean)[][]> {
  return await getDataBetweenDatesFromApi(
    table_name,
    start_date,
    end_date,
    great_or_equal,
    less_or_equal,
    columns,
    unique_identifier_list,
    excelStorageAdapter,
    excelHttpAdapter,
    // default base root from .env (overridden by SET_BASE_ROOT)
    config.baseRoot
  );
}
