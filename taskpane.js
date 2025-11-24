/* global Office, document */

Office.onReady(() => {
  const item = Office.context.mailbox.item;
  if (!item) {
    setStatus("No email item available.");
    return;
  }

  setStatus("Reading email...");
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailText = result.value || "";
      if (!emailText.trim()) {
        setStatus("Email has no readable body text.");
        return;
      }
      classifyEmail(emailText);
    } else {
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

  const rawScore = data.score;
  const label = data.label.toLowerCase();
  const isSpam = label === "phishing" || label === "spam";
  const scorePercent = isSpam ? Math.round(rawScore * 100) : Math.round((1 - rawScore) * 100);

  // Animate gauge
  const circle = document.querySelector('.progress-ring__circle');
  const radius = circle.r.baseVal.value;
  const circumference = radius * 2 * Math.PI;
  circle.style.strokeDasharray = `${circumference} ${circumference}`;
  const offset = circumference - (scorePercent / 100) * circumference;
  circle.style.strokeDashoffset = offset;

  // Update score text
  document.querySelector('.score-value').innerText = `${scorePercent}%`;

  // Update badge
  const badge = document.querySelector('.status-badge');
  badge.classList.remove("status-safe", "status-spam", "status-loading");
  if (isSpam) {
    badge.innerText = "SPAM DETECTED";
    badge.classList.add("status-spam");
  } else {
    badge.innerText = "SAFE";
    badge.classList.add("status-safe");
  }

  // Update analysis details
  document.getElementById("sender").innerText = data.sender || "--";
  document.getElementById("links").innerText = data.links || "--";
  document.getElementById("keywords").innerText = data.content || "--";
}

function setStatus(message) {
  const badge = document.querySelector('.status-badge');
  if (badge) {
    badge.innerText = message;
    badge.classList.remove("status-safe", "status-spam", "status-loading");
    badge.classList.add("status-loading");
  }
}
