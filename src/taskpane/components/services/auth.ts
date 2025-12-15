import { API } from "../utlis/Configs";

let APIURI: string;
if (process.env.NODE_ENV == "development") {
  APIURI = API.loacluri;
} else {
  APIURI = API.liveuri;
}

const auth = (email, password, callback) => {
  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");

  const raw = JSON.stringify({ email, password });

  const requestOptions: any = {
    method: "POST",
    headers: myHeaders,
    body: raw
  };

  fetch(`${APIURI}/auth/jwt-token/token/`, requestOptions)
    .then(async (response) => {
      const data = await response.json().catch(() => null); // avoid crash on empty body

      // Handle 401 Unauthorized separately
      if (response.status === 401) {
        return callback(null, {
          type: "unauthorized",
          message: data?.detail || "Invalid credentials",
          status: 401
        });
      }

      // Handle all other non-200 errors
      if (!response.ok) {
        return callback(null, {
          type: "error",
          message: data?.detail || "Request failed",
          status: response.status
        });
      }

      // Success
      return callback(data, null);
    })
    .catch((error) => {
      // Handle network errors (no internet, server unreachable)
      return callback(null, {
        type: "network",
        message: error.message || "Network error",
        status: null
      });
    });
};


export default auth