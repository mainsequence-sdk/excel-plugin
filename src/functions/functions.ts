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




const API = {
  // loacluri: "https://ts-orm-cloud-run-445866662347.europe-west1.run.app", //// for now get_data_between_dates_from_remote api not working on this
  loacluri: "https://dev-tsorm.ngrok.app",
  liveuri: "https://dev-tsorm.ngrok.app"
}

let APIURI: string;
if (process.env.NODE_ENV == "development") {
  APIURI = API.loacluri;
} else {
  APIURI = API.liveuri;
}

/**
 * Refresh the access token using refresh token
 */
const refresh = async (refreshToken) => {
  if (!refreshToken) throw new Error("No refresh token available");

  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");

  const raw = JSON.stringify({ refresh: refreshToken });

  const response = await fetch(
    `${APIURI}/auth/jwt-token/token/refresh/`,
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
  great_or_equal, // This will now be a boolean
  less_or_equal,  // This will now be a boolean
  limit,          // Dynamically passed
  offset,       // Dynamically passed
  update_hash, //null for now
) => {
  console.log(update_hash);
  const accessToken = await localStorage.getItem("token");
  const refreshToken = await localStorage.getItem("refresh_token");

  const makeRequest = async (token) => {
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");
    myHeaders.append("Authorization", `Bearer ${token}`);

    const raw = JSON.stringify({
      start_date: toUnixSeconds(start_date),
      end_date: toUnixSeconds(end_date),
      great_or_equal: great_or_equal, // Ensure this is a boolean
      less_or_equal: less_or_equal,   // Ensure this is a boolean
      unique_identifier_list: Array.isArray(unique_identifier_list)
        ? unique_identifier_list
        : [],
      columns: null,
      limit: limit,   // Use the passed limit
      offset: offset, // Use the passed offset
      node_identifier: node_identifier,
      update_hash: null
    });

    console.log(" Request Payload:", raw);

    const response = await fetch(
      `${APIURI}/orm/api/ts_manager/dynamic_table/get_data_between_dates_from_node_identifier/`,
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
 * @param {string[][]} unique_identifier_list List of unique identifiers (e.g., [["BBG000C1S2X2"], ["BBG000QH56C1"]])
 * @param {string} node_identifier node_identifier cell reference
 * @param {boolean} great_or_equal Include data greater than or equal to start date (should be TRUE/FALSE)
 * @param {boolean} less_or_equal Include data less than or equal to end date (should be TRUE/FALSE)
 * @param {number} [pageSize=1000] The number of rows to fetch per page.
 * @param {number} [offset=0] The starting offset for pagination.
 * @param {string} update_hash The starting offset for pagination.
 * @returns {Promise<string[][]>} A 2D array of data including headers and potentially a "More data available" message.
 */

async function GET_DATA(start_date, end_date, unique_identifier_list, node_identifier, great_or_equal, less_or_equal, pageSize = 1000, offset = 0, update_hash = null) {
  try {
    // Flatten unique_identifier_list if it's a 2D array from Excel input
    const flat_unique_identifier_list = unique_identifier_list ? unique_identifier_list.flat().filter(item => item !== "") : [];

    // Ensure great_or_equal and less_or_equal are actual booleans
    const isGreatOrEqual = typeof great_or_equal === 'boolean' ? great_or_equal : String(great_or_equal).toLowerCase() === 'true';
    const isLessOrEqual = typeof less_or_equal === 'boolean' ? less_or_equal : String(less_or_equal).toLowerCase() === 'true';

    const dataResponse = await FetchData(
      start_date,
      end_date,
      flat_unique_identifier_list,
      node_identifier,
      isGreatOrEqual,
      isLessOrEqual,
      pageSize,
      offset,
      update_hash
    );

    const results = Array.isArray(dataResponse) ? dataResponse : dataResponse.results;
    const nextOffset = dataResponse.next_offset;

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

    // Only add headers if it's the first page (offset 0)
    if (offset === 0) {
      finalOutput.push(headers);
    }
    finalOutput.push(...dataRows);

    // If there's more data, add a special row to indicate this
    if (nextOffset !== null && dataResponse.returned_count === pageSize) {
      const msg = `More data available. Next offset: ${nextOffset}. Page size: ${pageSize}`;
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

  const makeRequest = async (token: string) => {
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");
    myHeaders.append("Authorization", `Bearer ${token}`);

    const response = await fetch(
      `${APIURI}/orm/api/assets/asset/?unique_identifier=${encodeURIComponent(unique_identifier)}`,
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
 * @returns {Promise<string[][]>} A 2D array with keys in column 1, values in column 2
 */
async function GET_ASSET(unique_identifier: string): Promise<string[][]> {
  try {
    if (!unique_identifier || String(unique_identifier).trim() === "") {
      return [["Error", "unique_identifier is required"]];
    }

    const response = await FetchAsset(String(unique_identifier).trim());
    const asset = response && response.results && response.results.length > 0 ? response.results[0] : null;

    if (!asset) {
      return [["Error", "No asset found"]];
    }

    const MAX_EXCEL_CHARS = 32760;
    const safeString = (val: any) => {
      if (val === undefined || val === null) return "";
      const s = String(val);
      return s.length > MAX_EXCEL_CHARS ? s.slice(0, MAX_EXCEL_CHARS) + " ...[truncated]" : s;
    };

    const rows: string[][] = [];

    // Iterate through all keys in results[0] and create key/value rows
    Object.keys(asset).forEach((key) => {
      const val = (asset as any)[key];
      // If value is nested object/array, store as JSON string; otherwise as string
      const valueStr = (typeof val === "object" && val !== null)
        ? safeString(JSON.stringify(val))
        : safeString(val);
      rows.push([key, valueStr]);
    });

    return rows;
  } catch (err) {
    console.error("Error in GET_ASSET:", err);
    return [["Error", (err as any).message || "Unknown error"]];
  }
}

CustomFunctions.associate("GET_ASSET", GET_ASSET);