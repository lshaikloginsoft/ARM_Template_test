import { forwardMail } from "../helpers/sso-helper";
import CONFIG from "../config";

let reportBtn, cancelBtn, statusEl, progressContainer, progressMsg;

document.addEventListener("DOMContentLoaded", () => {
  const confirmEl = document.getElementById("confirmText");
  if (confirmEl) {
    confirmEl.innerText = CONFIG.UI.CONFIRM_DIALOG;
  }
});

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Cache DOM elements once
    reportBtn = document.getElementById("reportBtn");
    cancelBtn = document.getElementById("cancelBtn");
    statusEl = document.getElementById("status");
    progressContainer = document.getElementById("progressContainer");
    progressMsg = document.getElementById("progressMsg");

    // Assign event handlers
    reportBtn?.addEventListener("click", run);
    cancelBtn?.addEventListener("click", closeTaskPane);
  }
});

export async function run() {
  reportBtn.disabled = true;
  reportBtn.classList.add("btn-loading");
  showProgress(CONFIG.UI.REQUEST_PERMISSIONS);

  await new Promise(resolve => setTimeout(resolve, 50));

  try {
    const item = Office.context.mailbox.item;

    const messageId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    await forwardMail(
      messageId,
      handleResponse,
      () => {
        showProgress(CONFIG.UI.PROCESSING);
      }
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

  if (response.forward?.success && response.move?.success) {
    show(CONFIG.UI.SUCCESS_MESSAGE);
  }
  else if (response.forward?.success && !response.move?.success) {
    show(CONFIG.UI.MOVE_FAILED_TEXT);
  }
  else {
    const err = response.forward?.error;
    if (err) {
      show(`${CONFIG.UI.FAILED_TO_REPORT} ${err.replace(/^❌\s*/, "")}`);
    } else {
      show(CONFIG.UI.FAILED_TO_REPORT);
    }
  }

  setTimeout(() => {
    closeTaskPane();
  }, 3500);
}

// UI helpers
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