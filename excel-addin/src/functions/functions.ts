/**
 * Helper function to convert date to Unix seconds
 */
const toUnixSeconds = (d) => (d ? Math.floor(new Date(d).getTime() / 1000) : null);

/**
 * Refresh the access token using refresh token
 */
const refresh = async (refreshToken) => {
  if (!refreshToken) throw new Error("No refresh token available");

  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");

  const raw = JSON.stringify({ refresh: refreshToken });

  const response = await fetch(
    "https://dev-tsorm.ngrok.app/auth/jwt-token/token/refresh/",
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
  offset          // Dynamically passed
) => {
  const accessToken = await (OfficeRuntime.storage.getItem("token")?OfficeRuntime.storage.getItem("token"):localStorage.getItem("token"));
  const refreshToken =await (OfficeRuntime.storage.getItem("refresh_token")?OfficeRuntime.storage.getItem("refresh_token"):localStorage.getItem("refresh_token"));

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
      "https://dev-tsorm.ngrok.app/orm/api/ts_manager/dynamic_table/714/get_data_between_dates_from_remote/",
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
 * Excel Get Data
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
async function GETDATA(start_date, end_date, unique_identifier_list, great_or_equal, less_or_equal, pageSize = 1000, offset = 0) {
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

    const headers = [
      "Ticker",
      "Name",
      "ISIN",
      "Open",
      "Close",
      "High",
      "Low",
      "Volume",
      "VWAP",
      "Trade Date",
    ];

    const dataRows = results.map((item) => [
      String(item.ticker ?? ""),
      String(item.name ?? ""),
      String(item.isin ?? ""),
      Number(item.open ?? 0),
      Number(item.close ?? 0),
      Number(item.high ?? 0),
      Number(item.low ?? 0),
      Number(item.volume ?? 0),
      Number(item.vwap ?? 0),
      String(item.tradedate ?? ""),
    ]);

    
    let finalOutput = [];
    
    
    // Only add headers if it's the first page (offset 0)
    // and if there's no existing data, or if you explicitly want to refresh headers
    if (offset === 0) {
      finalOutput.push(headers);
    }
    finalOutput.push(...dataRows);
    // If there's more data, add a special row to indicate this.
    // The user would then call GETDATA again with the new offset.
    if (nextOffset !== null && dataResponse.returned_count === pageSize) {
        finalOutput.push(["", "", "", "", "", "", "", "", "", `More data available. Next offset: ${nextOffset}. Current page size: ${pageSize}`]);
    }


    console.log("📊 Returning cleaned data:", finalOutput.length, "rows (including headers/message)");
    return finalOutput;
  } catch (error) {
    console.error(" Error in GETDATA:", error);
    // Return a single cell error message to Excel
    return [["Error", error.message || "Unknown error"]];
  }
}

CustomFunctions.associate("GETDATA", GETDATA);













