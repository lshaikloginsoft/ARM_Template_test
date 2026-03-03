import * as $ from "jquery";

export async function forwardMailToMiddleTier(middletierToken,  messageId) {
  try {
    const response = await $.ajax({
      type: "POST",
      url: `/forwardMail`,
      headers: { Authorization: "Bearer " + middletierToken },
      data: JSON.stringify({
        messageId: messageId
      }),
      contentType: "application/json",
      cache: false,
    });
    return response;
  } catch (err) {
     // If backend returned structured JSON, return it
    if (err.responseJSON) {
      return err.responseJSON;
    }

    return {
      forward: { success: false, error: "❌ Server error, please try again later" },
      move: null
    };
  }
}
