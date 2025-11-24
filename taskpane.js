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
  let riskLevel = "Low";
  let gaugeColor = "#00FF94"; // green
  let needleAngle = -90; // left

  if (rawScore >= 0.7) {
    riskLevel = "High";
    gaugeColor = "#FF4B4B";
    needleAngle = 90;
  } else if (rawScore >= 0.4) {
    riskLevel = "Medium";
    gaugeColor = "#FFD700";
    needleAngle = 0;
  }

  // Animate needle
  const needle = document.getElementById("needle");
  if (needle) {
    needle.setAttribute("transform", `rotate(${needleAngle} 100 100)`);
  }

  // Update arc color
  const arc = document.getElementById("risk-arc");
  if (arc) {
    arc.setAttribute("stroke", gaugeColor);
  }

  // Update risk label
  const scoreEl = document.querySelector('.score-value');
  if (scoreEl) {
    scoreEl.innerText = riskLevel;
  }

  // Update badge
  const badge = document.querySelector('.status-badge');
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
  document.getElementById("sender").innerText = data.sender || "--";
  document.getElementById("links").innerText = data.links || "--";
  document.getElementById("keywords").innerText = data.content || "--";
  document.getElementById("attachment").innerText = data.attachment || "--";
}

function setStatus(message) {
  const badge = document.querySelector('.status-badge');
  if (badge) {
    badge.innerText = message;
    badge.classList.remove("status-safe", "status-spam", "status-medium");
    badge.classList.add("status-loading");
  }
}
