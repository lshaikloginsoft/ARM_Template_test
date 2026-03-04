import { forwardMail } from "../helpers/sso-helper";
import CONFIG from "../config";

let reportBtn, cancelBtn, statusEl, progressContainer, progressMsg;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {

    // Cache DOM elements
    reportBtn = document.getElementById("reportBtn");
    cancelBtn = document.getElementById("cancelBtn");
    statusEl = document.getElementById("status");
    progressContainer = document.getElementById("progressContainer");
    progressMsg = document.getElementById("progressMsg");

    const confirmEl = document.getElementById("confirmText");
    if (confirmEl) {
      confirmEl.innerText = CONFIG.UI.CONFIRM_DIALOG;
    }

    // Attach event handlers
    reportBtn?.addEventListener("click", run);
    cancelBtn?.addEventListener("click", closeTaskPane);
  }
});

export function run() {
  if (reportBtn.disabled) return;

  reportBtn.disabled = true;
  reportBtn.classList.add("btn-loading");

  showProgress(CONFIG.UI.REQUEST_PERMISSIONS);

  // Allow UI to render before starting heavy work
  setTimeout(startReportProcess, 0);
}

async function startReportProcess() {
  try {
    const item = Office.context.mailbox.item;

    const messageId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    showProgress(CONFIG.UI.PROCESSING);

    await forwardMail(
      messageId,
      handleResponse
    );

  } catch (err) {
    console.error(err);
    hideProgress();
    show(CONFIG.UI.FAILED_TO_REPORT);
    resetReportButton();
  }
}

function handleResponse(response) {
  console.log("Response received:", response);

  hideProgress();
  resetReportButton();

  let message = CONFIG.UI.FAILED_TO_REPORT;

  if (response.forward?.success && response.move?.success) {
    message = CONFIG.UI.SUCCESS_MESSAGE;
  }
  else if (response.forward?.success && !response.move?.success) {
    message = CONFIG.UI.MOVE_FAILED_TEXT;
  }
  else if (response.forward?.error) {
    message = `${CONFIG.UI.FAILED_TO_REPORT} ${response.forward.error.replace(/^❌\s*/, "")}`;
  }

  show(message);

  // Keep success message visible long enough
  setTimeout(() => {
    closeTaskPane();
  }, 5500);
}


// UI Helpers
function resetReportButton() {
  reportBtn.disabled = false;
  reportBtn.classList.remove("btn-loading");
}

function show(msg) {
  statusEl.innerText = msg;
}

function showProgress(msg) {
  progressContainer.style.display = "block";
  progressMsg.innerText = msg;
}

function hideProgress() {
  progressContainer.style.display = "none";
}

function closeTaskPane() {
  Office.context.ui.closeContainer();
}