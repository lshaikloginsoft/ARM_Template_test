import * as https from "https";
import { getAccessToken } from "./ssoauth-helper";
import CONFIG from "../config";
import { mapGraphErrorToUiMessage } from "./error-mapper";

const domain = "graph.microsoft.com";
const version = "v1.0";



async function withRetry(operation, description) {
  const maxAttempts = CONFIG.RETRY.MAX_ATTEMPTS;
  const baseDelay = CONFIG.RETRY.BASE_SLEEP_MS;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await operation();
    } catch (err) {

      const retryable =
        err.code === 408 ||
        err.code === 429 ||
        (typeof err.code === "number" && err.code >= 500);

      if (!retryable || attempt === maxAttempts) {
        console.error(`${description} failed after ${attempt} attempts`);
        throw err;
      }

      const delay = baseDelay * Math.pow(2, attempt - 1);
      console.warn(` RETRY ${attempt} for ${description} in ${delay}ms`);
      await new Promise(res => setTimeout(res, delay));
    }
  }
}

export async function forwardMail(req, res) {

  try {
    const recipient = CONFIG.SECURITY.RECIPIENT;

    if (!recipient || recipient.trim() === "") {
      return res.send({
        forward: { success: false, error: CONFIG.UI.ERROR_NO_RECIPIENT },
        move: null
      });
    }

    if (!isValidRecipient(recipient)) {
      return res.send({
        forward: { success: false, error: CONFIG.UI.INVALID_RECIPIENT_TEXT },
        move: null
      });
    }
    const graphToken = await resolveGraphToken(req);
    const messageId = extractMessageId(req);

    const mimeBuffer = await fetchMime(graphToken, messageId);

    await forwardMessage(graphToken, mimeBuffer);

    const moveResult = await moveOriginalMessage(graphToken, messageId);

    return res.send({
      forward: { success: true, error: null },
      move: moveResult
    });

  } catch (err) {
    console.error("forwardMail ERROR:", err);

    const friendlyMessage = mapGraphErrorToUiMessage(err);

    // Only return HTTP errors for auth problems
    if (err.code === 401 || err.code === 403) {
      return res.status(err.code).send({
        forward: { success: false, error: friendlyMessage },
        move: null
      });
    }

    return res.send({
      forward: { success: false, error: friendlyMessage },
      move: null
    });
  }
}

async function resolveGraphToken(req) {
  const authorization = req.get("Authorization");
  if (!authorization) {
    const err = new Error("Missing Authorization header");
    err.code = 401;
    throw err;
  }
  const graphTokenResponse = await getAccessToken(authorization);

  if (graphTokenResponse.error) {
    const err = new Error(graphTokenResponse.error);
    err.code = 401;
    throw err;
  }

  return graphTokenResponse.access_token;
}

function extractMessageId(req) {
  const messageId = req.body.messageId;

  if (!messageId) {
    const err = new Error("Missing messageId");
    err.code = 400;
    throw err;
  }

  return messageId;
}

async function fetchMime(graphToken, messageId) {

  const cleanId = encodeURIComponent(messageId);

  const mimeBuffer = await withRetry(
    () => makeGraphRawCall(graphToken, `/me/messages/${cleanId}/$value`),
    "Fetch MIME"
  );
  if (!mimeBuffer || mimeBuffer.length === 0) {
    const err = new Error("Fetch MIME failed");
    err.code = 500;
    err.type = "FETCH_FAILED";
    throw err;
  }

  return mimeBuffer;
}

async function forwardMessage(graphToken, mimeBuffer) {
  const recipient = CONFIG.SECURITY.RECIPIENT;
  const base64Mime = mimeBuffer.toString("base64");

  await withRetry(
    () => makeGraphApiCall(
      graphToken,
      "/me/sendMail",
      "POST",
      {
        message: {
          subject: CONFIG.EMAIL.SUBJECT,
          toRecipients: [
            { emailAddress: { address: recipient } }
          ],
          attachments: [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: CONFIG.EMAIL.EML_FILENAME,
            contentType: CONFIG.EMAIL.EML_MIME,
            contentBytes: base64Mime
          }]
        },
        saveToSentItems: false
      }
    ),
    "Forward Mail"
  );
}

async function moveOriginalMessage(graphToken, messageId) {

  try {
    await withRetry(
      () => ensureFolderAndMoveMessage(
        graphToken,
        CONFIG.EMAIL.MOVE_TO_FOLDER,
        messageId
      ),
      "Move Message"
    );

    return { success: true, error: null };

  } catch (err) {
    console.error("Move failed:", err.message);
    return { success: false, error: mapGraphErrorToUiMessage(err)};
  }
}


function makeGraphRawCall(accessToken, apiUrl) {
  return new Promise((resolve, reject) => {
    const options = {
      host: domain,
      path: `/${version}${apiUrl}`,
      method: "GET",
      headers: {
        Authorization: "Bearer " + accessToken
      }
    };

    const request = https.request(options, (response) => {

      const chunks = [];

      response.on("data", chunk => chunks.push(chunk));

      response.on("end", () => {

        if (response.statusCode >= 200 && response.statusCode < 300) {
          resolve(Buffer.concat(chunks));
        } else {

          const body = Buffer.concat(chunks).toString();

          let message = "Failed to fetch MIME";
          let graphCode = null;

          try {
            const parsed = JSON.parse(body);
            if (parsed.error) {
              message = parsed.error.message || message;
              graphCode = parsed.error.code;
            }
          } catch {}

          const error = new Error(message);
          error.code = response.statusCode;
          error.graphCode = graphCode;

          reject(error);
        }
      });
    });

    request.on("error", reject);
    request.end();
  });
}

function makeGraphApiCall(accessToken, apiUrl, method, body = null) {
  return new Promise((resolve, reject) => {

    const options = {
      host: domain,
      path: `/${version}${apiUrl}`,
      method,
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + accessToken
      }
    };

    const request = https.request(options, (response) => {

      let responseBody = "";

      response.on("data", chunk => responseBody += chunk);

      response.on("end", () => {

        if (response.statusCode >= 200 && response.statusCode < 300) {
          resolve(responseBody ? JSON.parse(responseBody) : {});
        } else {

          let message = "Graph API error";
          let graphCode = null;

          try {
            const parsed = JSON.parse(responseBody);
            if (parsed.error) {
              message = parsed.error.message || message;
              graphCode = parsed.error.code;
            }
          } catch {}

          const error = new Error(message);
          error.code = response.statusCode;
          error.graphCode = graphCode;

          reject(error);
        }
      });
    });

    request.on("error", reject);

    if (body) request.write(JSON.stringify(body));
    request.end();
  });
}

async function ensureFolderAndMoveMessage(graphToken, folderName, messageId) {

  let folder = await getFolderByName(graphToken, folderName);

  if (!folder) {
    folder = await createFolder(graphToken, folderName);
  }

  return await moveMessageToFolder(graphToken, messageId, folder.id);
}

async function getFolderByName(graphToken, folderName) {

  const escaped = folderName.replace(/'/g, "''");
  const filter = encodeURIComponent(`displayName eq '${escaped}'`);

  const folderList = await makeGraphApiCall(
    graphToken,
    `/me/mailFolders?$filter=${filter}`,
    "GET"
  );

  if (folderList.value && folderList.value.length > 0) {
    return folderList.value[0];
  }

  return null;
}

async function createFolder(graphToken, folderName) {
  try {
    return await makeGraphApiCall(
      graphToken,
      "/me/mailFolders",
      "POST",
      { displayName: folderName }
    );
  } catch (err) {

    // If folder already exists 
    if (err.code === 409 || err.graphCode === "ErrorFolderExists") {
      console.warn("Folder already exists. Re-fetching...");
      return await getFolderByName(graphToken, folderName);
    }

    throw err;
  }
}

async function moveMessageToFolder(graphToken, messageId, folderId) {

  const cleanId = encodeURIComponent(messageId);

  return await makeGraphApiCall(
    graphToken,
    `/me/messages/${cleanId}/move`,
    "POST",
    { destinationId: folderId }
  );
}

function isValidRecipient(email) {
  const e = String(email).trim().toLowerCase();

  const re = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
  if (!re.test(e)) return false;

  const parts = e.split("@");
  if (parts.length !== 2) return false;

  const domain = parts[1];
  const base = CONFIG.SECURITY.ALLOWED_RECIPIENT_BASE_DOMAIN.toLowerCase();

  if (domain === base) return true;

  if (domain.endsWith("." + base)) {
    const sub = domain.slice(0, domain.length - base.length - 1);
    if (!sub) return false;
    if (sub.includes("..")) return false;
    return true;
  }

  return false;
}