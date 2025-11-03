import {
  getDataBetweenDatesFromApi,
  setBaseRoot,
  StorageAdapter,
  HttpAdapter
} from '@excel-plugin/shared';

// Injected at build time from root .env
declare const process: any;
const DEFAULT_BASE_ROOT: string = (process && process.env && process.env.BASE_ROOT) || "";

// ---- Platform adapters (Apps Script) ----
const gsStorageAdapter: StorageAdapter = {
  getItem(key: string): string | null {
    try { return PropertiesService.getScriptProperties().getProperty(key); }
    catch { return null; }
  },
  setItem(key: string, value: string): void {
    PropertiesService.getScriptProperties().setProperty(key, value);
  }
};

const gsHttpAdapter: HttpAdapter = {
  fetch(url: string, options: any): GoogleAppsScript.URL_Fetch.HTTPResponse {
    return UrlFetchApp.fetch(url, {
      method: options?.method || 'get',
      contentType: options?.headers?.['Content-Type'] || 'application/json',
      payload: options?.body,
      muteHttpExceptions: true,
      followRedirects: true
    });
  }
};

// ---- Exported functions consumed by footer (NOT visible to Sheets directly) ----
export function __SET_BASE_ROOT(url: string): string {
  return setBaseRoot(url, gsStorageAdapter) as unknown as string;
}

export function __GET_DATA_BETWEEN_DATES_FROM_API(
  table_name: string,
  start_date?: any,
  end_date?: any,
  great_or_equal?: any,
  less_or_equal?: any,
  columns?: any[][],
  unique_identifier_list?: any[][]
): (string | number | boolean)[][] {
  return getDataBetweenDatesFromApi(
    table_name,
    start_date,
    end_date,
    great_or_equal,
    less_or_equal,
    columns,
    unique_identifier_list,
    gsStorageAdapter,
    gsHttpAdapter,
    DEFAULT_BASE_ROOT
  ) as unknown as (string | number | boolean)[][];
}

// ---- Optional debug/menu helpers ----
export function __onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('Plugin Debug')
    .addItem('Log BASE_ROOT', 'DEBUG_LOG_BASEROOT')
    .addItem('Ping API (HEAD)', 'DEBUG_PING')
    .addToUi();
}

export function __DEBUG_LOG_BASEROOT(): void {
  const prop = PropertiesService.getScriptProperties().getProperty('BASE_ROOT');
  const effective = prop || DEFAULT_BASE_ROOT || '(not set)';
  console.log('BASE_ROOT (script prop):', prop || '(none)');
  console.log('BASE_ROOT (effective):', effective);
}

export function __DEBUG_PING(): void {
  const prop = PropertiesService.getScriptProperties().getProperty('BASE_ROOT');
  const effective = prop || DEFAULT_BASE_ROOT;
  if (!effective) throw new Error('No BASE_ROOT. Use =SET_BASE_ROOT("https://...") or set .env');
  const url = effective.replace(/\/+$/,'') + '/';
  const resp = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, followRedirects: true });
  console.log('PING status:', resp.getResponseCode());
  console.log('PING headers:', JSON.stringify(resp.getAllHeaders()));
}
