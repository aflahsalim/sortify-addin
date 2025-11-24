/* global Office, document */

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

  const scorePercent = Math.round(data.score * 100);

  // Animate gauge
  const circle = document.querySelector('.progress-ring__circle');
  const radius = circle.r.baseVal.value;
  const circumference = radius * 2 * Math.PI;
  circle.style.strokeDasharray = `${circumference} ${circumference}`;
  circle.style.strokeDashoffset = circumference;
  const offset = circumference - (scorePercent / 100) * circumference;
  circle.style.strokeDashoffset = offset;

  // Update score text
  document.querySelector('.score-value').innerText = `${scorePercent}%`;

  // Update badge
  const badge = document.querySelector('.status-badge');
  if (data.label.toLowerCase() === "phishing" || data.label.toLowerCase() === "spam") {
    badge.innerText = "SPAM DETECTED";
    badge.classList.add("status-spam");
    badge.classList.remove("status-safe");
  } else {
    badge.innerText = "SAFE";
    badge.classList.add("status-safe");
    badge.classList.remove("status-spam");
  }

  // Update analysis details (mocked for now — replace with backend fields if available)
  document.getElementById("sender").innerText = data.sender || "Low / Unverified";
  document.getElementById("links").innerText = data.links || "Suspicious Redirects";
  document.getElementById("keywords").innerText = data.content || "Urgency Patterns";
}

function setStatus(msg) {
  const badge = document.querySelector('.status-badge');
  if (badge) badge.innerText = msg;
}

// Optional manual test button
function testBackend() {
  classifyEmail("This is a test email with suspicious links and urgency patterns.");
}
