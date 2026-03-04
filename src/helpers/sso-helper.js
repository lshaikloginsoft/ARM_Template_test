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
    } catch(err) {
      console.warn("Silent SSO failed, retrying with signin prompt", err);
      try {
        middletierToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true
        });
      } catch (err2) {
        console.warn("SSO failed, using dialog fallback", err2);
        dialogFallback(callback, messageId, onTokenAcquired);
        return;
      }
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
    console.warn("SSO flow failed, falling back to dialog auth", exception);
    dialogFallback(callback, messageId, onTokenAcquired);
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
        allowSignInPrompt: true,
      });

      const retryResponse = await forwardMailToMiddleTier(newToken, messageId);
      callback(retryResponse);
      return;
    } catch (exception) {
        console.error("Authentication failed", exception);
        dialogFallback(callback, messageId, onTokenAcquired);
    }
  }

  retryGetMiddletierToken = 0;
  dialogFallback(callback, messageId, onTokenAcquired);
}