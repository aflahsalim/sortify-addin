Office.onReady(() => {
  const sideload = document.getElementById("sideload-msg");
  if (sideload) sideload.style.display = "none";

  const appBody = document.getElementById("app-body");
  if (appBody) appBody.style.display = "block";

  const item = Office.context.mailbox.item;
  if (!item) {
    setStatus("No email item available.");
    return;
  }

  setStatus("Reading email...");
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Email body:", result.value);
      const emailText = result.value || "";
      if (!emailText.trim()) {
        setStatus("Email has no readable body text.");
        return;
      }
      classifyEmail(emailText);
    } else {
      console.error("getAsync error:", result.error);
      setStatus("Failed to read email body.");
    }
  });
});

function classifyEmail(emailText) {
  setStatus("Classifying email...");

  fetch("https://sortify-y7ru.onrender.com/classify", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text: emailText })
  })
    .then(async (res) => {
      if (!res.ok) {
        const body = await res.text();
        throw new Error(`Backend ${res.status}: ${body}`);
      }
      return res.json();
    })
    .then(showResult)
    .catch(err => {
      console.error("Fetch error:", err);
      setStatus("Error contacting backend.");
    });
}

function showResult(data) {
  if (!data || typeof data.score === "undefined" || !data.label) {
    setStatus("Invalid response from backend.");
    return;
  }

  document.getElementById("result").innerHTML =
    `<h3>Phishing Score: ${data.score}</h3><p>${data.label}</p>`;
}

function setStatus(msg) {
  const el = document.getElementById("result");
  if (el) el.innerText = msg;
}
