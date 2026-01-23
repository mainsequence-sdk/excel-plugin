import { getApiBaseUrl } from "../../../shared/apiConfig";

export const FetchData = async (
  start_date,
  end_date,
  unique_identifier_list,
  great_or_equal,
  less_or_equal,
  offset,
  // callback
) => {
  const apiBaseUrl = await getApiBaseUrl();
  const token = localStorage.getItem('token');
  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");
  myHeaders.append("Authorization", `Bearer ${token}`);

  const toUnixSeconds = (d) => {
    if (!d) return null;
    return Math.floor(new Date(d).getTime() / 1000);
  };

  const raw = JSON.stringify({
    start_date: toUnixSeconds(start_date),
    end_date: toUnixSeconds(end_date),
    great_or_equal: great_or_equal ?? true,
    less_or_equal: less_or_equal ?? true,
    unique_identifier_list: unique_identifier_list ?? [],
    columns: null,
    offset: offset,
  });

  const requestOptions: any = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow",
  };

  fetch(`${apiBaseUrl}/orm/api/ts_manager/dynamic_table/714/get_data_between_dates_from_remote/`, requestOptions)
    .then((response) => response.json())
    .then((result) => {
      console.log("Data fetched:", result);
      if (result.code == "token_not_valid") {
        console.log("Token expired. Refreshing token...");
      }
      // callback(result, null);
    })
    .catch((error) => {
      console.error("Error:", error);
      // callback(null, error);
    });
};
