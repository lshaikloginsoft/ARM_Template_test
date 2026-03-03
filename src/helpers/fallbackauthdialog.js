import { LogLevel, PublicClientApplication } from "@azure/msal-browser";
import { forwardMailToMiddleTier } from "./middle-tier-calls";

const clientId = "147b90e4-536b-4d3a-9a5a-058abb17e506"; //Replace with your client ID
const accessScope = 
     `api://outlook-addin-vmray.azurewebsites.net/147b90e4-536b-4d3a-9a5a-058abb17e506/access_as_user`; //Replace with Scope value
const loginRequest = {
  scopes: [accessScope]
};

const msalConfig = {
  auth: {
    clientId: clientId,
    authority: "https://login.microsoftonline.com/df620235-50d7-4400-bb7e-3b112e9b1ff4",// Replace with your tenant Id
    redirectUri: "https://outlook-addin-vmray.azurewebsites.net/fallbackauthdialog.html", //replace with your fallback redirect URI
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "localStorage", // Needed to avoid "User login is required" error.
    storeAuthStateInCookie: true, // Recommended to avoid certain IE/Edge issues.
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
    },
  },
};

const publicClientApp = new PublicClientApplication(msalConfig);

let loginDialog = null;
let homeAccountId = null;
let callbackFunction = null;
let storedMessageId = null;
let tokenAcquiredHandler = null;

Office.onReady(() => {
  if (Office.context.ui.messageParent) {
    publicClientApp
      .handleRedirectPromise()
      .then(handleResponse)
      .catch((error) => {
        console.log(error);
        Office.context.ui.messageParent(JSON.stringify({ status: "failure", result: error }));
      });

    // The very first time the add-in runs on a developer's computer, msal.js hasn't yet
    // stored login data in localStorage. So a direct call of acquireTokenRedirect
    // causes the error "User login is required". Once the user is logged in successfully
    // the first time, msal data in localStorage will prevent this error from ever hap-
    // pening again; but the error must be blocked here, so that the user can login
    // successfully the first time. To do that, call loginRedirect first instead of
    // acquireTokenRedirect.
    if (localStorage.getItem("loggedIn") === "yes") {
      publicClientApp.acquireTokenRedirect(loginRequest);
    } else {
      // This will login the user and then the (response.tokenType === "id_token")
      // path in authCallback below will run, which sets localStorage.loggedIn to "yes"
      // and then the dialog is redirected back to this script, so the
      // acquireTokenRedirect above runs.
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
