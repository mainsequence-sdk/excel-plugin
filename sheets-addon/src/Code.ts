// Apps Script (TypeScript). If you prefer .gs, remove types & compile with clasp or paste as JS.

// Note: Since Google Apps Script doesn't support ES6 imports, we copy the shared functions here
// In a production setup, you could use a build process to bundle the shared library

// Shared utility functions (copied from shared library)
function stripTrailingSlash(u: string): string {
  return u ? u.replace(/\/+$/g, "") : u;
}

function normalizeBoolean(v: any): boolean | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "boolean") return v;
  const s = String(v).trim().toLowerCase();
  if (["true", "t", "1", "yes", "y"].includes(s)) return true;
  if (["false", "f", "0", "no", "n"].includes(s)) return false;
  return null;
}

function flatten2DTo1D(arg: any): string[] | null {
  if (arg === null || arg === undefined || arg === "") return null;
  if (Array.isArray(arg)) {
    const out: string[] = [];
    for (const row of arg) {
      if (Array.isArray(row)) {
        for (const cell of row) {
          if (cell !== null && cell !== undefined && String(cell).trim() !== "") {
            out.push(String(cell));
          }
        }
      } else if (row !== null && row !== undefined) {
        out.push(String(row));
      }
    }
    return out;
  }
  return [String(arg)];
}

function parseToIsoDateString(arg: any): string | null {
  if (arg === null || arg === undefined || arg === "") return null;

  // Handle Date objects
  if (Object.prototype.toString.call(arg) === "[object Date]") {
    const d = arg as Date;
    if (!isNaN(d.getTime())) return d.toISOString();
    return null;
  }

  // Handle string dates
  const s = String(arg).trim();
  if (!s) return null;
  const d = new Date(s);
  if (isNaN(d.getTime())) return null;
  return d.toISOString();
}

function inferColumns(records: Record<string, any>[]): string[] {
  const seen = new Set<string>();
  const cols: string[] = [];
  for (const r of records) {
    for (const k of Object.keys(r || {})) {
      if (!seen.has(k)) {
        seen.add(k);
        cols.push(k);
      }
    }
  }
  return cols;
}

function coerceCellValue(v: any): string | number | boolean | "" {
  if (v == null) return "";
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;
  try { 
    return JSON.stringify(v); 
  } catch { 
    return String(v); 
  }
}

function recordsTo2D(
  records: Record<string, any>[], 
  columns: string[], 
  includeHeader = true
): (string | number | boolean)[][] {
  const out: (string | number | boolean)[][] = [];
  if (includeHeader) out.push(columns.map(c => c));
  for (const rec of records) {
    out.push(columns.map(c => coerceCellValue(rec?.[c])));
  }
  return out;
}

function ensureArray<T>(arr: T[] | null | undefined): T[] {
  return Array.isArray(arr) ? arr : [];
}

/**
 * Set the base API root for the add-on.
 * Usage: =SET_BASE_ROOT("https://api.example.com")
 * @param {string} url
 * @return {string} "OK"
 */
function SET_BASE_ROOT(url: string): string {
  const u = stripTrailingSlash(String(url || "").trim());
  if (!/^https?:\/\//i.test(u)) throw new Error("SET_BASE_ROOT: must be an http(s) URL.");
  PROPS.setProperty("BASE_ROOT", u);
  return "OK";
}

/**
 * Fetches rows from your Django API between two dates and returns a 2-D array
 * that fills from the formula cell.
 *
 * @param {string} table_name
 * @param {Date|string|number} [start_date]
 * @param {Date|string|number} [end_date]
 * @param {boolean|string} [great_or_equal]
 * @param {boolean|string} [less_or_equal]
 * @param {string[][]} [columns]
 * @param {string[][]} [unique_identifier_list]
 * @return {any[][]}
 * @customfunction
 */
function GET_DATA_BETWEEN_DATES_FROM_API(
  table_name: string,
  start_date?: any,
  end_date?: any,
  great_or_equal?: any,
  less_or_equal?: any,
  columns?: any[][],
  unique_identifier_list?: any[][]
): (string|number|boolean)[][] {
  const baseRoot = getBaseRoot();
  if (!baseRoot) throw new Error('BASE_ROOT not configured. Call =SET_BASE_ROOT("https://your-api-root") once.');

  const tableId = String(table_name || "").trim();
  if (!tableId) throw new Error("table_name (table_id) is required.");

  const url = `${baseRoot}/${encodeURIComponent(tableId)}/get_data_between_dates_from_remote/`;

  const payload = {
    start_date: parseToIsoDateString(start_date),
    end_date: parseToIsoDateString(end_date),
    great_or_equal: normalizeBoolean(great_or_equal),
    less_or_equal: normalizeBoolean(less_or_equal),
    columns: flatten2DTo1D(columns),
    unique_identifier_list: flatten2DTo1D(unique_identifier_list)
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    followRedirects: true
  });

  const status = resp.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error(`API ${status}: ${resp.getContentText()}`);
  }

  const data = JSON.parse(resp.getContentText());
  if (!Array.isArray(data)) throw new Error("API returned a non-list payload.");

  const cols = (payload.columns && payload.columns.length) ? payload.columns : inferColumns(data as Record<string,any>[]);
  return recordsTo2D(data as Record<string, any>[], cols, true);
}
