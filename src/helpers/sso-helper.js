import { dialogFallback } from "./fallbackauthdialog.js";
import { forwardMailToMiddleTier  } from "./middle-tier-calls";
import { handleClientSideErrors } from "./error-handler";


let retryGetMiddletierToken = 0;

export async function forwardMail(messageId, callback, onTokenAcquired) {
  try {
    let middletierToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      
    });
    if (onTokenAcquired) {
      onTokenAcquired();
    }
    let response = await forwardMailToMiddleTier(middletierToken, messageId);
    if (!response) {
      throw new Error("Middle tier didn't respond");
    } else if (response.claims) {
      // Microsoft Graph requires an additional form of authentication. Have the Office host
      // get a new token using the Claims string, which tells AAD to prompt the user for all
      // required forms of authentication.
      let mfaMiddletierToken = await Office.auth.getAccessToken({
        authChallenge: response.claims,
      });
      response = await forwardMailToMiddleTier(mfaMiddletierToken, messageId);
    }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    if (response.error) {
      await handleAADErrors(response, callback, middletierToken, messageId, onTokenAcquired);
    } else {
      retryGetMiddletierToken = 0;
      callback(response);
    }
  } catch (exception) {
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        dialogFallback(callback, messageId, onTokenAcquired);
      }
    } else {
      console.error("EXCEPTION: " + JSON.stringify(exception));
      throw exception;
    }
  }
}

async function handleAADErrors(response, callback, middletierToken, messageId, onTokenAcquired) {
  // On rare occasions the middle tier token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired middle tier token.

  if (response.error_description && response.error_description.indexOf("AADSTS500133") !== -1 && retryGetMiddletierToken === 0) {
    retryGetMiddletierToken++;
    const newToken = await Office.auth.getAccessToken({
      allowSignInPrompt: false,
      
    });
  const retryResponse = await forwardMailToMiddleTier(newToken, messageId);
  callback(retryResponse);
  } else {
    retryGetMiddletierToken = 0;
    dialogFallback(callback, messageId, onTokenAcquired);
  }
}
