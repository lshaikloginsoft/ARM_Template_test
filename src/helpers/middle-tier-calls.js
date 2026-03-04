export async function forwardMailToMiddleTier(middletierToken,  messageId) {
  try {
    const response = await fetch("/forwardMail", {
      method: "POST",
      headers: { 
        "Authorization": "Bearer " + middletierToken,
        "Content-Type": "application/json" 
      },
      body: JSON.stringify({
        messageId: messageId
      }),
    });
    const data = await response.json();
    return data;
  } catch (err) {
     // If backend returned structured JSON, return it
    console.error("forwardMailToMiddleTier error:", err);
     if (err.responseJSON) {
      return err.responseJSON;
    }
    return {
      forward: { success: false, error: "❌ Server error, please try again later" },
      move: null
    };
  }
}
