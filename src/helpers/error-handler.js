import { showMessage } from "./message-helper";

export function handleClientSideErrors(error) {
  let invokeFallBackDialog = false;

  switch (error.code) {

    case 13001:
      // No one signed into Office
      invokeFallBackDialog = true;
      return invokeFallBackDialog;

    case 13002:
      // User cancelled consent
      showMessage("❌ Permission request was cancelled. Please try again.");
      return invokeFallBackDialog;

    case 13006:
      // Office on Web issue
      showMessage("❌ Office on the Web is experiencing an issue. Please refresh and try again.");
      return invokeFallBackDialog;

    case 13008:
      // Operation still in progress
      showMessage("⏳ Office is completing a previous operation. Please try again in a moment.");
      return invokeFallBackDialog;

    case 13010:
      showMessage("❌ Browser configuration issue detected. Please try again or contact IT support.");
      return invokeFallBackDialog;

    default:
      // For all other SSO-related errors → fallback to dialog login
      invokeFallBackDialog = true;
      return invokeFallBackDialog;
  }
}