/* global Office, document */

Office.onReady(() => {
  const item = Office.context?.mailbox?.item;
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
          label: "ham",
          display: "Ham (Safe)",
          color: "green",
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
      if (typeof data.attachment === "undefined") {
        data.attachment = hasAttachment ? "Yes" : "No";
      }
      showResult(data);
      setStatus("Classification complete.");
    })
    .catch((err) => {
      console.error("Fetch error:", err);
      setStatus("Error contacting backend");
    });
}

function showResult(data) {
  const label = data.label || "unknown";
  const score = Number(data.score) || 0;

  // Fixed needle angles by category
  const angleMap = {
    ham: -90,
    support: -45,
    spam: 45,
    phishing: 90,
    unknown: -90,
  };
  const needleAngle = angleMap[label] ?? -90;

  // Color mapping from backend keywords to hex
  const colorMap = {
    green: "#28a745",
    orange: "#fd7e14",
    red: "#dc3545",
    blue: "#007bff",
  };
  const gaugeColor = colorMap[data.color] || "#00FF94";

  // Animate needle
  const needle = document.getElementById("needle");
  if (needle) {
    needle.setAttribute("transform", `rotate(${needleAngle} 100 100)`);
  }

  // Animate arc length by confidence + update color
  const arc = document.getElementById("risk-arc");
  if (arc) {
    arc.setAttribute("stroke", gaugeColor);
    const maxArc = 283; // half-circle path length used in SVG
    const offset = maxArc - (score * maxArc); // 0 = full, maxArc = empty
    arc.style.strokeDashoffset = offset;
  }

  // Update labels
  const scoreLabel = document.getElementById("score-label");
  const scoreValue = document.getElementById("score-value");
  if (scoreLabel) scoreLabel.textContent = data.display || label.toUpperCase();
  if (scoreValue) scoreValue.textContent = `${Math.round(score * 100)}%`;

  // Confidence text
  const confidenceEl = document.getElementById("confidence");
  if (confidenceEl) {
    confidenceEl.textContent = `Confidence: ${Math.round(score * 100)}%`;
  }

  // Update badge
  const badge = document.getElementById("status");
  if (badge) {
    badge.textContent = data.display || label.toUpperCase();
    badge.className = "status-badge"; // reset classes
    if (label === "phishing") badge.classList.add("status-spam");
    else if (label === "spam") badge.classList.add("status-medium");
    else if (label === "support") badge.classList.add("status-support");
    else badge.classList.add("status-safe");
  }

  // Update analysis details (placeholders until you compute these)
  setText("sender", data.sender || "--");
  setText("links", data.links || "--");
  setText("keywords", data.content || "--");
  setText("attachment", data.attachment || "--");
}

function setStatus(message) {
  const badge = document.getElementById("status");
  if (badge) {
    badge.textContent = message;
    badge.className = "status-badge status-loading";
  }
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}
