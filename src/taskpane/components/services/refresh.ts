import { getApiBaseUrl } from "../../../shared/apiConfig";

const refresh = async (token, callback) => {
  const apiBaseUrl = await getApiBaseUrl();
  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");

  const raw = JSON.stringify({
    refresh: token,
  });

  const requestOptions: any = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow"
  };

  fetch(`${apiBaseUrl}/auth/jwt-token/token/refresh/`, requestOptions)
    .then((response) => response.text())
    .then((result) => callback(result, null))
    .catch((error) => callback(null, error));
};

export default refresh;
