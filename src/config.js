const CONFIG = {
  EMAIL: {
    SUBJECT: "Phishing Report - Outlook Add-in",
    EML_FILENAME: "reported_email.eml",
    EML_MIME: "message/rfc822",
    MOVE_TO_FOLDER: "Phishing Reports"
  },

  RETRY: {
    MAX_ATTEMPTS: 3,
    BASE_SLEEP_MS: 250
  },

  UI: {
    CONFIRM_DIALOG:
      "Are you sure you want to report this email as phishing?",

    SUCCESS_MESSAGE:
      "✅ Email successfully reported to Security Team.",

    MOVE_FAILED_TEXT:
      "⚠️ Report sent, but moving the email failed.",

    FAILED_TO_REPORT:
        "❌ Failed to report email.",

    INVALID_RECIPIENT_TEXT:
      "❌ Invalid recipient configuration. Recipient must be under allowed domain.",

    REQUEST_PERMISSIONS:
      "Requesting permissions, a popup may appear...",

    PROCESSING:
      "Do not close this window until reporting is complete...",

    ERROR_NO_RECIPIENT:
      "❌ Recipient email not configured",

    ERROR_BAD_REQUEST:
        "❌ Unable to process this email.",

    ERROR_SERVER:
      "❌ Temporary server issue. Please try again later.",

    ERROR_SIZE_EXCEEDED:
      "❌ Email too large to process",
     ERROR_AUTH_EXPIRED:
        "❌ Authentication expired. Please sign in again.",

    ERROR_AUTH_FAILED:
        "❌ Authentication failed. Please sign in again.",

    ERROR_ACCESS_DENIED:
        "❌ Access denied.",

    ERROR_ITEM_NOT_FOUND:
        "❌ The selected email no longer exists.",

    ERROR_TIMEOUT:
        "❌ Request timed out. Please try again.",

    ERROR_THROTTLED:
        "❌ Service is busy. Please try again shortly.",

    ERROR_INVALID_MESSAGE:
        "❌ Unable to identify the selected email.",
  },

};

export default CONFIG;