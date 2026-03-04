import form from "form-urlencoded";
import jwt from "jsonwebtoken";
import { JwksClient } from "jwks-rsa";


const DISCOVERY_KEYS_ENDPOINT =
  `https://login.microsoftonline.com/${process.env.TENANT_ID}/discovery/v2.0/keys`;

const jwksClient = new JwksClient({
  jwksUri: DISCOVERY_KEYS_ENDPOINT,
  cache: true,
  rateLimit: true,
});

export async function getAccessToken(authorization) {
  if (!authorization) {
    throw new Error("Missing Authorization header");
  } else {
    const [, /* schema */ assertion] = authorization.split(" ");

    const decoded = jwt.decode(assertion);
    if (!decoded || !decoded.scp) {
      throw new Error("Invalid token scope");
    }
    const tokenScopes = decoded.scp.split(" ");
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

    const encodedForm = form(formParams);

    const tokenResponse = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/token`, {
      method: "POST",
      body: encodedForm,
      headers: {
        Accept: "application/json",
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });
    if (!tokenResponse.ok) {
      throw new Error("Failed to obtain Graph access token");
    }
    const json = await tokenResponse.json();
    return json;
  }
}

export function validateJwt(req, res, next) {
  const authHeader = req.headers.authorization;
  if (!authHeader) {
    return res.sendStatus(401);
  }
  const token = authHeader.split(" ")[1];
  const validationOptions = {
    audience: [
      `api://${req.headers.host}/${process.env.CLIENT_ID}`,
    ]
    };

  jwt.verify(token, getSigningKeys, validationOptions, (err) => {
    if (err) {
      console.error("JWT validation failed: ",err.message);
      return res.sendStatus(403);
    }

    next();
  });
}

function getSigningKeys(header, callback) {
  jwksClient.getSigningKey(header.kid, function (err, key) {
    if (err) {
      return callback(err);
    }
    callback(null, key.getPublicKey());
  });
}
