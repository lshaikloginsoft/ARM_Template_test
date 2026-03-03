import fetch from "node-fetch";
import form from "form-urlencoded";
import jwt from "jsonwebtoken";
import { JwksClient } from "jwks-rsa";


const DISCOVERY_KEYS_ENDPOINT =
  `https://login.microsoftonline.com/${process.env.TENANT_ID}/discovery/v2.0/keys`;

export async function getAccessToken(authorization) {
  if (!authorization) {
    let error = new Error("No Authorization header was found.");
    return Promise.reject(error);
  } else {
    const [, /* schema */ assertion] = authorization.split(" ");


    const tokenScopes = jwt.decode(assertion).scp.split(" ");
    const accessAsUserScope = tokenScopes.find((scope) => scope === "access_as_user");
    if (!accessAsUserScope) {
      throw new Error("Missing access_as_user");
    }

    const formParams = {
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: assertion,
      requested_token_use: "on_behalf_of",
      resource: "https://graph.microsoft.com"
    };

    const stsDomain = "https://login.microsoftonline.com";
    const tenant = process.env.TENANT_ID;
    const tokenURLSegment = "oauth2/token";
    const encodedForm = form(formParams);

    const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
      method: "POST",
      body: encodedForm,
      headers: {
        Accept: "application/json",
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });
    const json = await tokenResponse.json();
    return json;
  }
}

export function validateJwt(req, res, next) {
  const authHeader = req.headers.authorization;
  if (authHeader) {
    const token = authHeader.split(" ")[1];
    const decoded = jwt.decode(token);
    const validationOptions = {
      audience: [
        `api://outlook-web-app.azurewebsites.net/${process.env.CLIENT_ID}`
      ],
    };

    jwt.verify(token, getSigningKeys, validationOptions, (err) => {
      if (err) {
        console.log(err);
        return res.sendStatus(403);
      }

      next();
    });
  }
}

function getSigningKeys(header, callback) {
  var client = new JwksClient({
    jwksUri: DISCOVERY_KEYS_ENDPOINT,
  });

  client.getSigningKey(header.kid, function (err, key) {
    callback(null, key.getPublicKey());
  });
}
