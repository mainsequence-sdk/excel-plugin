/**
 * Helper function to convert date to Unix seconds
 */
// const toUnixSeconds = (d) => (d ? Math.floor(new Date(d).getTime() / 1000) : null);

const toUnixSeconds = (value: number | string | Date | null | undefined): number | null => {
  if (value === null || value === undefined || value === "") return null;

  console.log("toUnixSeconds input:", value, typeof value);

  // 1) If it's already a Date object
  if (value instanceof Date) {
    const ms = value.getTime();
    if (isNaN(ms)) return null;
    return Math.floor(ms / 1000);
  }

  // 2) If it's a number (Excel serial date)
  if (typeof value === "number") {
    // Basic sanity check for Excel serials (common range)
    if (value < 60 || value > 60000) {
      console.warn("Number value outside likely Excel serial range:", value);
      return null;
    }
    const excelEpochMs = Date.UTC(1899, 11, 30); // 1899-12-30 UTC in ms
    const jsDateMs = excelEpochMs + value * 86400000; // days -> ms
    return Math.floor(jsDateMs / 1000);
  }

  // 3) If it's a string:
  if (typeof value === "string") {
    const trimmed = value.trim();

    // 3a) Numeric-only string: treat as Excel serial if in range
    if (/^\d+$/.test(trimmed)) {
      const n = Number(trimmed);
      if (!isNaN(n) && n >= 60 && n <= 60000) {
        const excelEpochMs = Date.UTC(1899, 11, 30);
        const jsDateMs = excelEpochMs + n * 86400000;
        return Math.floor(jsDateMs / 1000);
      } else {
        console.warn("Numeric string but outside Excel serial range or invalid:", trimmed);
        return null;
      }
    }

    // 3b) Otherwise try parse as ISO / normal date string
    const parsed = new Date(trimmed);
    const ms = parsed.getTime();
    if (!isNaN(ms)) {
      return Math.floor(ms / 1000);
    }

    // Not a valid date string
    console.warn("String value is not a valid date:", trimmed);
    return null;
  }

  // Not recognisable type
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
  great_or_equal, // This will now be a boolean
  less_or_equal,  // This will now be a boolean
  limit,          // Dynamically passed
  offset,       // Dynamically passed
) => {
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
    });

    console.log(" Request Payload:", raw);

    const response = await fetch(
      `${APIURI}/orm/api/ts_manager/dynamic_table/714/get_data_between_dates_from_remote/`,
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
 * @param {boolean} great_or_equal Include data greater than or equal to start date (should be TRUE/FALSE)
 * @param {boolean} less_or_equal Include data less than or equal to end date (should be TRUE/FALSE)
 * @param {number} [pageSize=1000] The number of rows to fetch per page.
 * @param {number} [offset=0] The starting offset for pagination.
 * @returns {string[][]} A 2D array of data including headers and potentially a "More data available" message.
 */
async function GET_DATA(start_date, end_date, unique_identifier_list, great_or_equal, less_or_equal, pageSize = 1000, offset = 0) {
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
      isGreatOrEqual, // Pass the converted boolean
      isLessOrEqual,  // Pass the converted boolean
      pageSize,
      offset
    );

    const results = Array.isArray(dataResponse) ? dataResponse : dataResponse.results;
    const nextOffset = dataResponse.next_offset;

    if (!results || results.length === 0) {
      console.log("⚠️ No records found.");
      return [["Info", "No data found for given filters"]];
    }

    // const headers = [
    //   "Ticker",
    //   "Name",
    //   "ISIN",
    //   "Open",
    //   "Close",
    //   "High",
    //   "Low",
    //   "Volume",
    //   "VWAP",
    //   "Trade Date",
    // ];


    const headers = [
      "Ticker",
      "Name",
      "ISIN",
      "Open",
      "Open Size",
      "High",
      "High Time",
      "Low",
      "Low Time",
      "Close",
      "Close Size",
      "Volume",
      "Listed Market Hours Volume",
      "Listed Market Hours Trades",
      "Listed Total Trades",
      "FINRA Market Hours Volume",
      "FINRA Market Hours Trades",
      "FINRA Total Volume",
      "FINRA Total Trades",
      "VWAP",
      "Daily VWAP",
      "Primary Exchange",
      "Unique Identifier",
      "SECID",
      "Trade Date",
      "Time Index",
      "Adjustment Factor",
      "Adjustment Reason",
      "Cumulative Price Factor",
      "Cumulative Volume Factor",
      "Open Time"
    ];



    // const dataRows = results.map((item) => [
    //   String(item.ticker ?? ""),
    //   String(item.name ?? ""),
    //   String(item.isin ?? ""),
    //   Number(item.open ?? 0),
    //   Number(item.close ?? 0),
    //   Number(item.high ?? 0),
    //   Number(item.low ?? 0),
    //   Number(item.volume ?? 0),
    //   Number(item.vwap ?? 0),
    //   String(item.tradedate ?? ""),
    // ]);

    const dataRows = results.map((item) => [
      String(item.ticker || ""),
      String(item.name || ""),
      String(item.isin || ""),
      String(item.open ?? ""),
      String(item.opensize ?? ""),
      String(item.high ?? ""),
      String(item.hightime || ""),
      String(item.low ?? ""),
      String(item.lowtime || ""),
      String(item.close ?? ""),
      String(item.closesize ?? ""),
      String(item.volume ?? ""),
      String(item.listedmarkethoursvolume ?? ""),
      String(item.listedmarkethourstrades ?? ""),
      String(item.listedtotaltrades ?? ""),
      String(item.finramarkethoursvolume ?? ""),
      String(item.finramarkethourstrades ?? ""),
      String(item.finratotalvolume ?? ""),
      String(item.finratotaltrades ?? ""),
      String(item.vwap ?? ""),
      String(item.dailyvwap ?? ""),
      String(item.primaryexchange || ""),
      String(item.unique_identifier || ""),
      String(item.secid ?? ""),
      String(item.tradedate ?? ""),
      String(item.time_index || ""),
      String(item.adjustmentfactor ?? ""),
      String(item.adjustmentreason ?? ""),
      String(item.cumulativepricefactor ?? ""),
      String(item.cumulativevolumefactor ?? ""),
      String(item.open_time ?? "")
    ]);


    let finalOutput = [];


    // Only add headers if it's the first page (offset 0)
    // and if there's no existing data, or if you explicitly want to refresh headers
    if (offset === 0) {
      finalOutput.push(headers);
    }
    finalOutput.push(...dataRows);
    // If there's more data, add a special row to indicate this.
    // The user would then call GET_DATA again with the new offset.
    // if (nextOffset !== null && dataResponse.returned_count === pageSize) {
    //     finalOutput.push(["", "", "", "", "", "", "", "", "", `More data available. Next offset: ${nextOffset}. Current page size: ${pageSize}`]);
    // }


    if (nextOffset !== null && dataResponse.returned_count === pageSize) {
      const msg = `More data available. Next offset: ${nextOffset}. Page size: ${pageSize}`;
      finalOutput.push([msg, ...Array(headers.length - 1).fill("")]);
    }


    console.log("📊 Returning cleaned data:", finalOutput.length, "rows (including headers/message)");
    return finalOutput;
  } catch (error) {
    console.error(" Error in GET_DATA:", error);
    // Return a single cell error message to Excel
    return [["Error", error.message || "Unknown error"]];
  }
}

CustomFunctions.associate("GET_DATA", GET_DATA);













