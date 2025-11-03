/**
 * Shared API functionality for Excel and Google Sheets add-ons
 */
import {
  stripTrailingSlash,
  normalizeBoolean,
  flatten2DTo1D,
  parseToIsoDateString,
  inferColumns,
  recordsTo2D,
  ensureArray
} from './utils';
import { ApiPayload, StorageAdapter, HttpAdapter } from './types';

/**
 * Resolve BASE_ROOT from storage, or fall back to a provided default.
 */
async function resolveBaseRoot(
  storageAdapter: StorageAdapter,
  defaultBaseRoot?: string
): Promise<string> {
  const fromStorage = await storageAdapter.getItem("BASE_ROOT");
  const base = fromStorage && fromStorage.trim()
    ? fromStorage
    : (defaultBaseRoot || "");
  return stripTrailingSlash(base);
}

/**
 * Core API function used by both Excel and Google Sheets
 */
export async function getDataBetweenDatesFromApi(
  tableName: string,
  startDate?: any,
  endDate?: any,
  greatOrEqual?: any,
  lessOrEqual?: any,
  columns?: any[][],
  uniqueIdentifierList?: any[][],
  storageAdapter?: StorageAdapter,
  httpAdapter?: HttpAdapter,
  defaultBaseRoot?: string
): Promise<(string | number | boolean)[][]> {
  if (!storageAdapter || !httpAdapter) {
    throw new Error("Platform adapters not provided.");
  }

  const baseRoot = await resolveBaseRoot(storageAdapter, defaultBaseRoot);
  if (!baseRoot) throw new Error('BASE_ROOT not configured. Call SET_BASE_ROOT or set it in .env.');

  const tableId = String(tableName || "").trim();
  if (!tableId) throw new Error("table_name (table_id) is required.");

  const url = `${baseRoot}/${encodeURIComponent(tableId)}/get_data_between_dates_from_remote/`;

  const payload: ApiPayload = {
    start_date: parseToIsoDateString(startDate),
    end_date: parseToIsoDateString(endDate),
    great_or_equal: normalizeBoolean(greatOrEqual),
    less_or_equal: normalizeBoolean(lessOrEqual),
    columns: flatten2DTo1D(columns),
    unique_identifier_list: flatten2DTo1D(uniqueIdentifierList)
  };

  const response = await httpAdapter.fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });

  let data: any;
  let status: number;

  // Google Apps Script UrlFetchApp response has these methods:
  if (response && typeof response.getResponseCode === 'function') {
    status = response.getResponseCode();
    if (status < 200 || status >= 300) {
      throw new Error(`API ${status}: ${response.getContentText()}`);
    }
    data = JSON.parse(response.getContentText());
  } else {
    // Standard fetch Response
    status = response.status;
    if (!response.ok) {
      const text = await response.text().catch(() => "");
      throw new Error(`API ${status} ${response.statusText}: ${text || "Request failed."}`);
    }
    data = await response.json();
  }

  if (!Array.isArray(data)) {
    throw new Error("API returned a non-list payload. Expected a JSON array of records.");
  }

  const cols = ensureArray(payload.columns).length ? ensureArray(payload.columns) : inferColumns(data);
  return recordsTo2D(data as Record<string, any>[], cols, true);
}

/**
 * Sets the base root URL using the provided storage adapter
 */
export async function setBaseRoot(url: string, storageAdapter: StorageAdapter): Promise<string> {
  const u = stripTrailingSlash(String(url || "").trim());
  if (!/^https?:\/\//i.test(u)) {
    throw new Error("SET_BASE_ROOT: must be an http(s) URL.");
  }
  await storageAdapter.setItem("BASE_ROOT", u);
  return "OK";
}
