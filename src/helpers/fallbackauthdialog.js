import { LogLevel, PublicClientApplication } from "@azure/msal-browser";
import { forwardMailToMiddleTier } from "./middle-tier-calls";

const loginRequest = {
  scopes: []
};

async function loadConfig() {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = `${window.location.origin}/runtime-config.js`;
    script.onload = () => resolve(window.APP_CONFIG);
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

let publicClientApp = null;

async function initializeMsal() {

  const cfg = await loadConfig();

  const clientId = cfg.clientId;
  const tenantId = cfg.tenantId;
  const domain = window.location.hostname;

  const accessScope =
    `api://${domain}/${clientId}/access_as_user`;

  loginRequest.scopes = [accessScope];

  const msalConfig = {
    auth: {
      clientId: clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri: `https://${domain}/fallbackauthdialog.html`,
      navigateToLoginRequestUrl: false
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: true
    },
    system: {
      loggerOptions: {
        loggerCallback: (level, message, containsPii) => {
          if (containsPii) return;

          switch (level) {
            case LogLevel.Error:
              console.error(message);
              break;
            case LogLevel.Info:
              console.info(message);
              break;
            case LogLevel.Verbose:
              console.debug(message);
              break;
            case LogLevel.Warning:
              console.warn(message);
              break;
          }
        }
      }
    }
  };

  publicClientApp = new PublicClientApplication(msalConfig);
}

let loginDialog = null;
let homeAccountId = null;
let callbackFunction = null;
let storedMessageId = null;
let tokenAcquiredHandler = null;

Office.onReady(async () => {
  await initializeMsal();
  if (Office.context.ui.messageParent) {
    publicClientApp
      .handleRedirectPromise()
      .then(handleResponse)
      .catch((error) => {
        console.log(error);
        Office.context.ui.messageParent(JSON.stringify({ status: "failure", result: error }));
      });

    if (localStorage.getItem("loggedIn") === "yes") {
      publicClientApp.acquireTokenRedirect(loginRequest);
    } else {
      publicClientApp.loginRedirect(loginRequest);
    }
  }
});

function handleResponse(response) {
  if (response.tokenType === "id_token") {
    console.log("LoggedIn");
    localStorage.setItem("loggedIn", "yes");
  } else {
    console.log("token type is:" + response.tokenType);
    Office.context.ui.messageParent(
      JSON.stringify({
        status: "success",
        result: response.accessToken,
        accountId: response.account.homeAccountId,
      })
    );
  }
}

export async function dialogFallback(callback, messageId, onTokenAcquired) {
  // Attempt to acquire token silently if user is already signed in.
  storedMessageId = messageId;
  callbackFunction = callback;
  tokenAcquiredHandler = onTokenAcquired;
  if (homeAccountId !== null) {
    try {
      const result = await publicClientApp.acquireTokenSilent(loginRequest);

      if (result && result.accessToken) {

        if (tokenAcquiredHandler) {
          tokenAcquiredHandler();
        }

        const response = await forwardMailToMiddleTier(result.accessToken, messageId);
        callbackFunction(response);
        return;
      }

    } catch (silentError) {
      console.warn("Silent token failed. Falling back to dialog.", silentError);
      // Continue to popup fallback
    }
  }
  const url = "/fallbackauthdialog.html";
  showLoginPopup(url);
}

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
async function processMessage(arg) {
  // Uncomment to view message content in debugger, but don't deploy this way since it will expose the token.
  //console.log("Message received in processMessage: " + JSON.stringify(arg));

  let messageFromDialog = JSON.parse(arg.message);

  if (messageFromDialog.status === "success") {
    // We now have a valid access token.
    loginDialog.close();

    // Configure MSAL to use the signed-in account as the active account for future requests.
    const homeAccount = publicClientApp.getAccountByHomeId(messageFromDialog.accountId);
    if (homeAccount) {
      homeAccountId = messageFromDialog.accountId; // Track the account id for future silent token requests.
      publicClientApp.setActiveAccount(homeAccount);
    }
    if (tokenAcquiredHandler) {
      tokenAcquiredHandler();
    }
    const response = await forwardMailToMiddleTier(messageFromDialog.result, storedMessageId);
    
    callbackFunction(response);
  } else if (
    messageFromDialog.error === undefined &&
    messageFromDialog.result.errorCode === undefined
  ) {
    // Need to pick the user to use to auth
  } else {
    // Something went wrong with authentication or the authorization of the web application.
    loginDialog.close();
    if (messageFromDialog.error) {
      console.error(JSON.stringify(messageFromDialog.error.toString()));
    } else if (messageFromDialog.result) {
      console.error(JSON.stringify(messageFromDialog.result.errorMessage.toString()));
    }
  }
}

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
function showLoginPopup(url) {
  var fullUrl =
    location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") + url;

  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(fullUrl, { height: 60, width: 30 }, function (result) {
    console.log("Dialog has initialized. Wiring up events");
    loginDialog = result.value;
    loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
  });
}
