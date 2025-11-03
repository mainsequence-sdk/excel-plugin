/**
 * Shared utility functions for Excel and Google Sheets add-ons
 */

export function stripTrailingSlash(u: string): string {
  return u ? u.replace(/\/+$/g, "") : u;
}

export function normalizeBoolean(v: any): boolean | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "boolean") return v;
  const s = String(v).trim().toLowerCase();
  if (["true", "t", "1", "yes", "y"].includes(s)) return true;
  if (["false", "f", "0", "no", "n"].includes(s)) return false;
  return null;
}

export function flatten2DTo1D(arg: any): string[] | null {
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

export function parseToIsoDateString(arg: any): string | null {
  if (arg === null || arg === undefined || arg === "") return null;

  // Date object
  if (arg instanceof Date) {
    if (!isNaN(arg.getTime())) return arg.toISOString();
    return null;
  }

  // Excel serial (only Excel side uses; harmless elsewhere)
  if (typeof arg === "number" && isFinite(arg)) {
    const ms = Math.round((arg - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (!isNaN(d.getTime())) return d.toISOString();
  }

  // String
  const s = String(arg).trim();
  if (!s) return null;
  const d = new Date(s);
  if (isNaN(d.getTime())) return null;
  return d.toISOString();
}

export function inferColumns(records: Record<string, any>[]): string[] {
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

export function coerceCellValue(v: any): string | number | boolean | "" {
  if (v == null) return "";
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;
  try { return JSON.stringify(v); } catch { return String(v); }
}

export function recordsTo2D(
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

export function ensureArray<T>(arr: T[] | null | undefined): T[] {
  return Array.isArray(arr) ? arr : [];
}
