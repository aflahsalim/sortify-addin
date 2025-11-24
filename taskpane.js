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

      // Even if body text is empty, still show attachment info and a Low risk by default
      if (!emailText.trim()) {
        setStatus("Email has no readable body text.");
        showResult({
          score: 0, // treat empty as low risk
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
    body: JSON.stringify({
      text: emailText,
      attachment: hasAttachment ? "Yes" : "No",
    }),
  })
    .then(async (res) => {
      if (!res.ok) {
        const body = await res.text();
        throw new Error(`Backend ${res.status}: ${body}`);
      }
      return res.json();
    })
    .then((data) => {
      // Ensure attachment is present in UI even if backend doesn't return it
      if (typeof data.attachment === "undefined") {
        data.attachment = hasAttachment ? "Yes" : "No";
      }
      showResult(data);
    })
    .catch((err) => {
      console.error("Fetch error:", err);
      setStatus("Error contacting backend.");
    });
}

function showResult(data) {
  if (!data || typeof data.score === "undefined" || !data.label) {
    setStatus("Invalid response from backend.");
    return;
  }

  // Determine risk level from score (0–1). No percentages shown.
  const rawScore = Number(data.score) || 0;
  let riskLevel = "Low";
  let gaugeColor = "#00FF94"; // green
  let needleAngle = -90; // left (Low risk)

  if (rawScore >= 0.7) {
    riskLevel = "High";
    gaugeColor = "#FF4B4B"; // red
    needleAngle = 90; // right
  } else if (rawScore >= 0.4) {
    riskLevel = "Medium";
    gaugeColor = "#FFD700"; // yellow
    needleAngle = 0; // center
  }

  // Update needle rotation
  const needle = document.getElementById("needle");
  if (needle) {
    needle.setAttribute("transform", `rotate(${needleAngle} 100 100)`);
  }

  // Update arc color
  const arc = document.getElementById("risk-arc");
  if (arc) {
    arc.setAttribute("stroke", gaugeColor);
  }

  // Update risk label (no percentage, just Low/Medium/High)
  const scoreEl = document.querySelector(".score-value");
  if (scoreEl) {
    scoreEl.innerText = riskLevel;
  }

  // Update badge to match risk
  const badge = document.querySelector(".status-badge");
  if (badge) {
    badge.classList.remove("status-safe", "status-spam", "status-loading", "status-medium");
    if (riskLevel === "High") {
      badge.innerText = "RISK DETECTED";
      badge.classList.add("status-spam");
    } else if (riskLevel === "Medium") {
      badge.innerText = "POTENTIAL RISK";
      badge.classList.add("status-medium");
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
  const badge = document.querySelector(".status-badge");
  if (badge) {
    badge.innerText = message;
    badge.classList.remove("status-safe", "status-spam", "status-medium");
    badge.classList.add("status-loading");
  }
}
