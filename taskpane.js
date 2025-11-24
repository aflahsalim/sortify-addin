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
      const hasAttachment =
        Array.isArray(item.attachments) && item.attachments.length > 0;

      if (!emailText.trim()) {
        // Even if body text is empty, we still show attachment info in the UI
        setStatus("Email has no readable body text.");
        showResult({
          score: 0,
          label: "safe",
          sender: "--",
          links: "--",
          content: "No content",
          attachment: hasAttachment ? "Yes" : "No",
        });
        return;
      }

      classifyEmail(emailText, hasAttachment);
    } else {
      setStatus("Failed to read email body.");
    }
  });
});

function classifyEmail(emailText, hasAttachment) {
  setStatus("Classifying email...");

  fetch("https://sortify-y7ru.onrender.com/classify", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text: emailText, attachment: hasAttachment ? "Yes" : "No" })
  })
    .then(async (res) => {
      if (!res.ok) {
        const body = await res.text();
        throw new Error(`Backend ${res.status}: ${body}`);
      }
      return res.json();
    })
    .then((data) => {
      // Ensure attachment shows even if backend doesn’t return it
      if (typeof data.attachment === "undefined") {
        data.attachment = hasAttachment ? "Yes" : "No";
      }
      showResult(data);
    })
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

  const rawScore = Number(data.score) || 0;
  const label = String(data.label).toLowerCase();
  const isSpam = label === "phishing" || label === "spam";
  const scorePercent = isSpam ? Math.round(rawScore * 100) : Math.round((1 - rawScore) * 100);

  // Animate gauge safely
  const circle = document.querySelector('.progress-ring__circle');
  if (circle && circle.r && circle.r.baseVal) {
    const radius = circle.r.baseVal.value;
    const circumference = radius * 2 * Math.PI;
    circle.style.strokeDasharray = `${circumference} ${circumference}`;
    const offset = circumference - (scorePercent / 100) * circumference;
    circle.style.strokeDashoffset = offset;
  }

  // Update score text
  const scoreEl = document.querySelector('.score-value');
  if (scoreEl) scoreEl.innerText = `${scorePercent}%`;

  // Update badge
  const badge = document.querySelector('.status-badge');
  if (badge) {
    badge.classList.remove("status-safe", "status-spam", "status-loading");
    if (isSpam) {
      badge.innerText = "SPAM DETECTED";
      badge.classList.add("status-spam");
    } else {
      badge.innerText = "SAFE";
      badge.classList.add("status-safe");
    }
  }

  // Update analysis details
  const senderEl = document.getElementById("sender");
  const linksEl = document.getElementById("links");
  const contentEl = document.getElementById("keywords");
  const attachmentEl = document.getElementById("attachment");

  if (senderEl) senderEl.innerText = data.sender || "--";
  if (linksEl) linksEl.innerText = data.links || "--";
  if (contentEl) contentEl.innerText = data.content || "--";
  if (attachmentEl) attachmentEl.innerText = data.attachment || "--";
}

function setStatus(message) {
  const badge = document.querySelector('.status-badge');
  if (badge) {
    badge.innerText = message;
    badge.classList.remove("status-safe", "status-spam", "status-loading");
    badge.classList.add("status-loading");
  }
}
