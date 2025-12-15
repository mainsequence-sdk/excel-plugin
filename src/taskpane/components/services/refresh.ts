import { API } from "../utlis/Configs";

let APIURI: string;
if (process.env.NODE_ENV == "development") {
  APIURI = API.loacluri;
} else {
  APIURI = API.liveuri;
}

const refresh = (token, callback) => {
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

  fetch(`${APIURI}/auth/jwt-token/token/refresh/`, requestOptions)
    .then((response) => response.text())
    .then((result) => callback(result, null))
    .catch((error) => callback(null, error));
};

export default refresh;
