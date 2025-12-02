console.log("✅ Sortify gauge logic loaded");
/* global Office, document */

Office.onReady(() => {
  waitForGauge(() => {
    initializeGauge();

    // Test render (remove in production)
    showResult({
      label: "support",
      display: labelDisplay("support"),
      sender: "support@company.com",
      links: "No suspicious links",
      content: "No phishing keywords",
      attachment: "No"
    });

    startClassification();
  });
});

function waitForGauge(callback) {
  const arc = document.getElementById("risk-arc");
  const needle = document.getElementById("needle");
  if (arc && needle) callback();
  else requestAnimationFrame(() => waitForGauge(callback));
}

function initializeGauge() {
  const arc = document.getElementById("risk-arc");
  if (arc) {
    const len = arc.getTotalLength();
    arc.setAttribute("stroke-dasharray", len);
    arc.style.strokeDashoffset = len;
    arc.dataset.arcLength = String(len);
  }

  const needle = document.getElementById("needle");
  if (needle) {
    needle.setAttribute("transform", "rotate(-90 100 80)");
  }
}

function showResult(data) {
  const label = (data.label || "unknown").toLowerCase();
  const display = data.display || labelDisplay(label);
  const color = getRiskColor(label);
  const angle = getFixedAngle(label);
  const fillRatio = getFillRatio(label);

  // Rotate needle
  const needle = document.getElementById("needle");
  if (needle) {
    needle.setAttribute("transform", `rotate(${angle} 100 80)`);
  }

  // Fill arc up to needle
  const arc = document.getElementById("risk-arc");
  if (arc) {
    const len = parseFloat(arc.dataset.arcLength || arc.getTotalLength());
    arc.style.strokeDashoffset = len - fillRatio * len;
    arc.setAttribute("stroke", color);
  }

  // Label (classification type)
  const labelEl = document.getElementById("score-label");
  if (labelEl) {
    labelEl.textContent = label.toUpperCase(); // e.g., PHISHING
    labelEl.style.color = color;
  }

  // Button (risk level)
  const button = document.getElementById("result-button");
  if (button) {
    button.textContent = display.toUpperCase(); // e.g., HIGH RISK
    button.style.background = color;
  }

  // Details (with fallbacks)
  setText("sender", data.sender || "Unknown");
  setText("links", data.links || "No link data");
  setText("keywords", data.content || "No keyword data");
  setText("attachment", typeof data.attachment === "string" ? data.attachment : (data.attachment ? "Yes" : "No"));
}

function getRiskColor(label) {
  switch (label) {
    case "ham": return "#28a745";       // Green
    case "support": return "#00bfff";   // Blue
    case "spam": return "#fd7e14";      // Orange
    case "phishing": return "#dc3545";  // Red
    default: return "#6c757d";          // Gray
  }
}

function labelDisplay(label) {
  switch (label) {
    case "ham": return "Safe";
    case "support": return "Safe";
    case "spam": return "Risk";
    case "phishing": return "High Risk";
    default: return "Unknown";
  }
}

function getFixedAngle(label) {
  switch (label) {
    case "ham": return -90;
    case "support": return -45;
    case "spam": return 45;
    case "phishing": return 90;
    default: return 0;
  }
}

function getFillRatio(label) {
  switch (label) {
    case "ham": return 0.0;
    case "support": return 0.25;
    case "spam": return 0.75;
    case "phishing": return 1.0;
    default: return 0.5;
  }
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value || "--";
}

function setStatus(msg) {
  const badge = document.getElementById("status");
  if (badge) badge.textContent = msg;
}

function startClassification() {
  const item = Office.context?.mailbox?.item;
  if (!item) return setStatus("No email item available.");

  setStatus("Reading email...");
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded)
      return setStatus("Failed to read email body.");

    const emailText = result.value || "";
    const hasAttachment = Array.isArray(item.attachments) && item.attachments.length > 0;

    if (!emailText.trim()) return setStatus("Email has no readable body text.");

    classifyEmail(emailText, hasAttachment);
  });
}

function classifyEmail(emailText, hasAttachment) {
  setStatus("Classifying email...");

  fetch("https://sortify-y7ru.onrender.com/classify", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      text: emailText,
      attachment: hasAttachment ? "Yes" : "No"
    }),
  })
    .then(async (res) => {
      if (!res.ok) {
        const body = await res.text();
        throw new Error(body || `HTTP ${res.status}`);
      }
      return res.json();
    })
    .then((data) => {
      const label = String(data.label || "unknown").toLowerCase();

      showResult({
        label,
        display: labelDisplay(label),
        sender: data.sender,
        links: data.links,
        content: data.content,
        attachment: data.attachment
      });

      setStatus("Classification complete.");
    })
    .catch((err) => {
      console.error("Backend error:", err);
      setStatus("Error contacting backend.");
    });
}
