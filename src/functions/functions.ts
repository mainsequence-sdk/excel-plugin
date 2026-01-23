import { getApiBaseUrl } from "../shared/apiConfig";

/**
 * Helper function to convert date to Unix seconds
 */
// const toUnixSeconds = (d) => (d ? Math.floor(new Date(d).getTime() / 1000) : null);

// Works for ALL Excel serial dates up to year 9999
// Works on Windows + Mac (1900 date system only)
// Does NOT limit by year (no more 2065 cutoff)

const toUnixSeconds = (
  value: number | string | Date | null | undefined
): number | null => {

  if (value === null || value === undefined || value === "") return null;

  // 1) Already a JS Date object
  if (value instanceof Date) {
    const ms = value.getTime();
    return isNaN(ms) ? null : Math.floor(ms / 1000);
  }

  // 2) Excel Serial Number (no range limits!)
  if (typeof value === "number") {
    // Excel epoch: 1899-12-30
    const excelEpochMs = Date.UTC(1899, 11, 30);
    const jsDateMs = excelEpochMs + value * 86400000;
    return Math.floor(jsDateMs / 1000);
  }

  // 3) String
  if (typeof value === "string") {
    const trimmed = value.trim();

    // 3a) If string contains only digits → treat as Excel serial
    if (/^\d+$/.test(trimmed)) {
      const serial = Number(trimmed);
      const excelEpochMs = Date.UTC(1899, 11, 30);
      const jsDateMs = excelEpochMs + serial * 86400000;
      return Math.floor(jsDateMs / 1000);
    }

    // 3b) Try treat as ISO / normal date
    const parsed = new Date(trimmed);
    const ms = parsed.getTime();
    return isNaN(ms) ? null : Math.floor(ms / 1000);
  }

  return null;
};

const coerce1DStringList = (value: unknown): string[] => {
  const normalized: string[] = [];

  const normalizeItem = (item: unknown): string | null => {
    if (item === null || item === undefined) return null;
    const trimmed = String(item).trim();
    if (!trimmed || trimmed.toLowerCase() === "null") return null;
    return trimmed;
  };

  const walk = (item: unknown): void => {
    if (Array.isArray(item)) {
      item.forEach(walk);
      return;
    }
    const norm = normalizeItem(item);
    if (norm !== null) normalized.push(norm);
  };

  walk(value);
  return normalized;
};

const isCustomFunctionsError = (value: unknown): value is { code?: string; message?: string } => {
  return typeof value === "object" && value !== null && "code" in value && "message" in value;
};

const findFirstCustomFunctionsError = (value: unknown): { code?: string; message?: string } | null => {
  if (isCustomFunctionsError(value)) return value;
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findFirstCustomFunctionsError(item);
      if (found) return found;
    }
  }
  return null;
};




/**
 * Refresh the access token using refresh token
 */
const refresh = async (refreshToken) => {
  if (!refreshToken) throw new Error("No refresh token available");
  const apiBaseUrl = await getApiBaseUrl();

  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");

  const raw = JSON.stringify({ refresh: refreshToken });

  const response = await fetch(
    `${apiBaseUrl}/auth/jwt-token/token/refresh/`,
    {
      method: "POST",
      headers: myHeaders,
      body: raw,
      redirect: "follow",
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Token refresh failed: ${response.status}: ${errorText}`);
  }
  return response.json(); // { access: "...", refresh: "..." }
};

/**
 * Fetch data from your backend with pagination support
 */
export const FetchData = async (
  start_date,
  end_date,
  unique_identifier_list,
  node_identifier,
  limit,          // Dynamically passed
  columns
) => {
  const accessToken = await OfficeRuntime.storage.getItem("token");
  const refreshToken = await OfficeRuntime.storage.getItem("refresh_token");
  const apiBaseUrl = await getApiBaseUrl();

  if (!accessToken) {
    throw new Error("Missing access token. Please sign in again.");
  }

  const makeRequest = async (token) => {
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");
    myHeaders.append("Authorization", `Bearer ${token}`);

    const normalizedUniqueIdentifiers = coerce1DStringList(unique_identifier_list);
    const normalizedColumns = coerce1DStringList(columns);

    const raw = JSON.stringify({
      start_date: toUnixSeconds(start_date),
      end_date: toUnixSeconds(end_date),
      unique_identifier_list: normalizedUniqueIdentifiers,
      columns: normalizedColumns.length > 0 ? normalizedColumns : null,
      limit: limit,   // Use the passed limit
      node_identifier: node_identifier
    });

    console.log(" Request Payload:", raw);

    const response = await fetch(
      `${apiBaseUrl}/orm/api/ts_manager/dynamic_table/get_data_between_dates_from_node_identifier/`,
      // `${APIURI}/orm/api/ts_manager/dynamic_table/714/get_data_between_dates_from_remote/`,
      {
        method: "POST",
        headers: myHeaders,
        body: raw,
        redirect: "follow",
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTPS ${response.status}: ${errorText}`);
    }

    return response.json();
  };

  try {
    const result = await makeRequest(accessToken);
    console.log(" Data fetched successfully:", result);
    return result;
  } catch (error) {
    console.warn(" Error fetching data:", error.message);

    // If token expired or invalid — try refresh
    if (error.message.includes("token_not_valid") || error.message.includes("401")) {
      console.log(" Refreshing token...");
      const newTokens = await refresh(refreshToken);

      await OfficeRuntime.storage.setItem("token", newTokens.access);
      if (newTokens.refresh) {
        await OfficeRuntime.storage.setItem("refresh_token", newTokens.refresh);
      }

      console.log("Token refreshed, retrying request...");
      return makeRequest(newTokens.access);
    }

    throw error;
  }
};

/**
 * Excel Custom Function: Fetch and show table in Excel with pagination hint
 * @customfunction
 * @param {string} start_date Start date (e.g., "2022-01-01")
 * @param {string} end_date End date (e.g., "2022-01-31")
 * @param {string[][]} unique_identifier_list List/range of unique identifiers (row or column range is flattened to 1D)
 * @param {string} node_identifier node_identifier cell reference
 * @param {number} [pageSize=1000] The number of rows to fetch per page.
 * @param {string[][]} [columns] List/range of columns to return (row or column range is flattened to 1D)
 * @returns {Promise<string[][]>} A 2D array of data including headers and potentially a "More data available" message.
 */

async function GET_DATA(start_date, end_date, unique_identifier_list, node_identifier, pageSize = 1000, columns = null) {
  try {
    const inputError = findFirstCustomFunctionsError([
      start_date,
      end_date,
      unique_identifier_list,
      node_identifier,
      pageSize,
      columns,
    ]);
    if (inputError) {
      const msg = inputError.message || inputError.code || "Invalid input";
      return [["Error", msg]];
    }

    const flat_unique_identifier_list = coerce1DStringList(unique_identifier_list);
    const flat_columns = coerce1DStringList(columns);

    const normalizeParam = (value: unknown) => {
      if (value === null || value === undefined) return null;
      const trimmed = String(value).trim();
      if (!trimmed || trimmed.toLowerCase() === "null") return null;
      return trimmed;
    };

    const normalizedNodeIdentifier = normalizeParam(node_identifier);

    if (!normalizedNodeIdentifier) {
      return [["Error", "Provide node_identifier."]];
    }

    const dataResponse = await FetchData(
      start_date,
      end_date,
      flat_unique_identifier_list,
      normalizedNodeIdentifier,
      pageSize,
      flat_columns
    );

    const results = Array.isArray(dataResponse) ? dataResponse : dataResponse.results;

    if (!results || results.length === 0) {
      console.log("No records found.");
      return [["Info", "No data found for given filters"]];
    }

    // Dynamically get all unique keys from results
    const allKeys = new Set<string>();
    results.forEach((item: Record<string, any>) => {
      Object.keys(item).forEach(key => allKeys.add(key));
    });
    const headers = Array.from(allKeys) as string[];

    // helper: make values Excel-safe (primitive + length limit)
    const safeString = (val: any) => {
      if (val === undefined || val === null) return "";
      const s = String(val);
      const MAX_EXCEL_CHARS = 32760; // leave small margin under 32,767
      return s.length > MAX_EXCEL_CHARS ? s.slice(0, MAX_EXCEL_CHARS) + `... [truncated ${s.length}]` : s;
    };

    // Dynamically map results to rows based on headers
    const dataRows = results.map((item: Record<string, any>) =>
      headers.map((key: string) => safeString(item[key]))
    );

    const finalOutput: string[][] = [];

    // Always add headers for first (and only) page
    finalOutput.push(headers);
    finalOutput.push(...dataRows);

    // If there's more data, add a special row to indicate this
    if (typeof dataResponse.next_offset === "number" && dataResponse.returned_count === pageSize) {
      const msg = `More data available. Increase page size or use API pagination. Page size: ${pageSize}`;
      finalOutput.push([msg, ...Array(headers.length - 1).fill("")]);
    }

    console.log("Returning cleaned data:", finalOutput.length, "rows (including headers/message)");
    return finalOutput;
  } catch (error) {
    console.error(" Error in GET_DATA:", error);
    return [["Error", (error && (error as any).message) || "Unknown error"]];
  }
}

CustomFunctions.associate("GET_DATA", GET_DATA);






/**
 * Fetch asset details from backend by unique identifier
 */
export const FetchAsset = async (unique_identifier: string) => {
  const accessToken = await OfficeRuntime.storage.getItem("token");
  const refreshToken = await OfficeRuntime.storage.getItem("refresh_token");
  const apiBaseUrl = await getApiBaseUrl();

  const makeRequest = async (token: string) => {
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");
    myHeaders.append("Authorization", `Bearer ${token}`);

    const response = await fetch(
      `${apiBaseUrl}/orm/api/assets/asset/?unique_identifier=${encodeURIComponent(unique_identifier)}`,
      {
        method: "GET",
        headers: myHeaders,
        redirect: "follow",
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP ${response.status}: ${errorText}`);
    }

    return response.json();
  };

  try {
    const result = await makeRequest(accessToken);
    console.log("Asset fetched successfully:", result);
    return result;
  } catch (error) {
    console.warn("Error fetching asset:", error.message);

    // If token expired or invalid — try refresh
    if (error.message.includes("token_not_valid") || error.message.includes("401")) {
      console.log("Refreshing token...");
      const newTokens = await refresh(refreshToken);

      await OfficeRuntime.storage.setItem("token", newTokens.access);
      if (newTokens.refresh) {
        await OfficeRuntime.storage.setItem("refresh_token", newTokens.refresh);
      }

      console.log("Token refreshed, retrying request...");
      return makeRequest(newTokens.access);
    }

    throw error;
  }
};




/**
 * Excel Custom Function: Get asset details by unique identifier
 * @customfunction
 * @param {string} unique_identifier The unique identifier of the asset (e.g., "90_CBPF_48")
 * @param {boolean} [spill_rows=true] If TRUE, return rows of field/value pairs; if FALSE, return a single JSON cell.
 * @returns {Promise<any[][]>} A 2D array suitable for Excel.
 */
async function GET_ASSET(unique_identifier: string, spill_rows: boolean = true): Promise<any[][]> {
  try {
    if (!unique_identifier || String(unique_identifier).trim() === "") {
      return [["Error", "unique_identifier is required"]];
    }

    const normalizedSpill = typeof spill_rows === "boolean"
      ? spill_rows
      : String(spill_rows).toLowerCase() === "true";

    const response = await FetchAsset(String(unique_identifier).trim());
    const asset = response && response.results && response.results.length > 0 ? response.results[0] : null;

    if (!asset) {
      return [["Error", "Asset not found."]];
    }

    const MAX_EXCEL_CHARS = 32760;
    const safeString = (val: string) => {
      const s = val ?? "";
      return s.length > MAX_EXCEL_CHARS ? s.slice(0, MAX_EXCEL_CHARS) + " ...[truncated]" : s;
    };

    const jsonReplacer = (_key: string, value: any) => {
      if (value instanceof Date) return value.toISOString();
      return value;
    };

    const serializeValue = (val: any): any => {
      if (val === undefined || val === null) return "";
      if (val instanceof Date) return val;
      if (Array.isArray(val) || typeof val === "object") {
        return safeString(JSON.stringify(val, jsonReplacer));
      }
      if (typeof val === "string") return safeString(val);
      return val;
    };

    if (normalizedSpill) {
      const rows: any[][] = [["field", "value"]];
      Object.keys(asset).forEach((key) => {
        const val = (asset as any)[key];
        rows.push([key, serializeValue(val)]);
      });
      return rows;
    }

    const payload = safeString(JSON.stringify(asset, jsonReplacer));
    return [[payload]];
  } catch (err) {
    console.error("Error in GET_ASSET:", err);
    return [["Error", (err as any).message || "Unknown error"]];
  }
}

CustomFunctions.associate("GET_ASSET", GET_ASSET);

/**
 * Excel Custom Function: Get a specific asset field by dot path
 * @customfunction
 * @param {string} unique_identifier The unique identifier of the asset (e.g., "90_CBPF_48")
 * @param {string} field_path Dot-separated field path (e.g., "current_snapshot.ticker")
 * @returns {Promise<any[][]>} A single-cell 2D array containing the field value.
 */
async function GET_ASSET_FIELD(unique_identifier: string, field_path: string): Promise<any[][]> {
  try {
    if (!unique_identifier || String(unique_identifier).trim() === "") {
      return [["Error", "unique_identifier is required"]];
    }
    if (!field_path || String(field_path).trim() === "") {
      return [["Error", "field_path is required"]];
    }

    const response = await FetchAsset(String(unique_identifier).trim());
    const asset = response && response.results && response.results.length > 0 ? response.results[0] : null;

    if (!asset) {
      return [["Error", "Asset not found."]];
    }

    const MAX_EXCEL_CHARS = 32760;
    const safeString = (val: string) => {
      const s = val ?? "";
      return s.length > MAX_EXCEL_CHARS ? s.slice(0, MAX_EXCEL_CHARS) + " ...[truncated]" : s;
    };

    const jsonReplacer = (_key: string, value: any) => {
      if (value instanceof Date) return value.toISOString();
      return value;
    };

    const serializeValue = (val: any): any => {
      if (val === undefined || val === null) return "";
      if (val instanceof Date) return val;
      if (Array.isArray(val) || typeof val === "object") {
        return safeString(JSON.stringify(val, jsonReplacer));
      }
      if (typeof val === "string") return safeString(val);
      return val;
    };

    let current: any = asset;
    const parts = String(field_path).split(".");
    for (const part of parts) {
      if (current === null || current === undefined) {
        throw new Error(`Unable to resolve field '${field_path}': null/undefined`);
      }
      if (Array.isArray(current) && /^\d+$/.test(part)) {
        const idx = Number(part);
        if (idx >= current.length) {
          throw new Error(`Unable to resolve field '${field_path}': index out of range`);
        }
        current = current[idx];
        continue;
      }
      if (typeof current !== "object" || !(part in current)) {
        throw new Error(`Unable to resolve field '${field_path}': missing '${part}'`);
      }
      current = (current as any)[part];
    }

    return [[serializeValue(current)]];
  } catch (err) {
    console.error("Error in GET_ASSET_FIELD:", err);
    return [["Error", (err as any).message || "Unknown error"]];
  }
}

CustomFunctions.associate("GET_ASSET_FIELD", GET_ASSET_FIELD);


function safeParse(json: string): any | null {
  try {
    return JSON.parse(json);
  } catch (e) {
    return null;
  }
}

/**
 * Excel Custom Function
 * GET_JSON_VALUE(jsonText, keyPath)
 * 
 * @customfunction
 * @param {string} jsonText The JSON text stored in a cell
 * @param {string} keyPath Dot-notation key path (e.g., "key2.s1")
 * @returns {string}
 */
function GET_JSON_VALUE(jsonText: string, keyPath: string): string {
  if (!jsonText) return "Invalid JSON";

  const obj = safeParse(jsonText);
  if (!obj) return "Invalid JSON";

  // Direct key like "key1"
  // Nested key like "key2.s1"
  const keys = keyPath.split(".");

  let current: any = obj;

  for (const key of keys) {
    if (current === null || current === undefined) return "Not found";

    if (typeof current !== "object") return "Not found";

    if (!(key in current)) return "Not found";

    current = current[key];
  }

  // If result is object, return JSON string
  if (typeof current === "object") {
    return JSON.stringify(current);
  }

  // Otherwise return value as string
  return String(current);
}

CustomFunctions.associate("GET_JSON_VALUE", GET_JSON_VALUE);
