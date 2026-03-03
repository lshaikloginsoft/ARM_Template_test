export function showMessage(text) {
  const el = document.getElementById("status");
  if (el) el.innerText = text;
}