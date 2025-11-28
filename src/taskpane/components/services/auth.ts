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

  const raw = JSON.stringify({
    "email": email,
    "password": password
  });

  const requestOptions: any = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow"
  };

  fetch(`${APIURI}/auth/jwt-token/token/`, requestOptions)
    .then((response) => response.json())
    .then((result) => { callback(result, null), console.log(result) })
    .catch((error) => callback(null, error));
}

export default auth