import CONFIG from "../config";

export function mapGraphErrorToUiMessage(err) {

  if (!err) {
    return CONFIG.UI.ERROR_SERVER;
  }

  // Config Errors

  if (err.type === "CONFIG_RECIPIENT_MISSING") {
    return CONFIG.UI.ERROR_NO_RECIPIENT;
  }

  if (err.type === "CONFIG_INVALID_RECIPIENT") {
    return CONFIG.UI.INVALID_RECIPIENT_TEXT;
  }

  // Empty/Missing

  if (err.message === "Missing messageId") {
    return CONFIG.UI.ERROR_INVALID_MESSAGE;
  }

  if (err.message === "Empty message content") {
    return CONFIG.UI.ERROR_SERVER;
  }

  // Graph Erros

  switch (err.graphCode) {

    case "InvalidAuthenticationToken":
      return CONFIG.UI.ERROR_AUTH_EXPIRED;

    case "ErrorAccessDenied":
      return CONFIG.UI.ERROR_ACCESS_DENIED;

    case "ErrorItemNotFound":
      return CONFIG.UI.ERROR_ITEM_NOT_FOUND;

    case "ErrorInvalidRecipients":
    case "ErrorRecipientNotFound":
      return CONFIG.UI.INVALID_RECIPIENT_TEXT;

    case "RequestEntityTooLarge":
    case "ErrorMessageSizeExceeded":
      return CONFIG.UI.ERROR_SIZE_EXCEEDED;
  }

  // Size Exceed
  if (
    typeof err.message === "string" &&
    err.message.toLowerCase().includes("exceed")
  ) {
    return CONFIG.UI.ERROR_SIZE_EXCEEDED;
  }

  // Http Status

  switch (err.code) {
    case 400:
        return CONFIG.UI.ERROR_BAD_REQUEST;

    case 401:
      return CONFIG.UI.ERROR_AUTH_FAILED;

    case 403:
      return CONFIG.UI.ERROR_ACCESS_DENIED;

    case 404:
      return CONFIG.UI.ERROR_ITEM_NOT_FOUND;

    case 408:
      return CONFIG.UI.ERROR_TIMEOUT;

    case 413:
      return CONFIG.UI.ERROR_SIZE_EXCEEDED;

    case 429:
      return CONFIG.UI.ERROR_THROTTLED;
  }

  if (typeof err.code === "number" && err.code >= 500) {
    return CONFIG.UI.ERROR_SERVER;
  }

  return CONFIG.UI.ERROR_SERVER;
}