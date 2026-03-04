import { dialogFallback } from "./fallbackauthdialog.js";
import { forwardMailToMiddleTier } from "./middle-tier-calls";
import { handleClientSideErrors } from "./error-handler";

let retryGetMiddletierToken = 0;

export async function forwardMail(messageId, callback, onTokenAcquired) {
  try {
    let middletierToken;

    try {
      middletierToken = await Office.auth.getAccessToken({
        allowSignInPrompt: false
      });
    } catch {
      console.warn("Silent SSO failed, retrying with signin prompt", err);
      middletierToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true
      });
    }

    if (onTokenAcquired) {
      onTokenAcquired();
    }

    let response = await forwardMailToMiddleTier(middletierToken, messageId);

    if (!response) {
      throw new Error("No response received from server.");
    }

    if (response.claims) {
      const mfaToken = await Office.auth.getAccessToken({
        authChallenge: response.claims,
      });
      response = await forwardMailToMiddleTier(mfaToken, messageId);
    }

    if (response.error) {
      await handleAADErrors(response, callback, messageId, onTokenAcquired);
      return;
    }

    retryGetMiddletierToken = 0;
    callback(response);

  } catch (exception) {
    if (exception.code && handleClientSideErrors(exception)) {
      dialogFallback(callback, messageId, onTokenAcquired);
      return;
    }

    callback({
      forward: { success: false, error: "Unable to authenticate your session. Please try again." },
      move: null
    });
  }
}

async function handleAADErrors(response, callback, messageId, onTokenAcquired) {
  if (
    response.error_description &&
    response.error_description.includes("AADSTS500133") &&
    retryGetMiddletierToken === 0
  ) {
    retryGetMiddletierToken++;

    try {
      const newToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true
      });

      const retryResponse = await forwardMailToMiddleTier(newToken, messageId);
      callback(retryResponse);
      return;
    } catch {
    }
  }

  retryGetMiddletierToken = 0;
  dialogFallback(callback, messageId, onTokenAcquired);
}