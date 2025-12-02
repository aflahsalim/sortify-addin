console.log("✅ Sortify JS restored");
/* global Office, document */

Office.onReady(() => {
  waitForGauge(() => {
    initializeGauge();

    // Demo render (remove in production)
    showResult({
      score: 0.6,
      label: "spam",
      display: labelDisplay("spam"),
      sender: "debug@example.com",
      links: "None",
      content: "Debug content",
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
  const score = resolveScore(data.score);
  const label = (data.label || "unknown").toLowerCase();
  const display = data.display || labelDisplay(label);
  const color = getRiskColor(label);

  // Rotate needle proportional to score
  const needle = document.getElementById("needle");
  if (needle) {
    const angle = -90 + score * 180;
    needle.setAttribute("transform", `rotate(${angle} 100 80)`);
  }

  // Arc stroke
  const arc = document.getElementById("risk-arc");
  if (arc) {
    const len = parseFloat(arc.dataset.arcLength || arc.getTotalLength());
    arc.style.strokeDashoffset = len - score * len;
    arc.setAttribute("stroke", color);
  }

  // Label
  const labelEl = document.getElementById("score-label");
  if (labelEl) {
    labelEl.textContent = display;
    labelEl.style.color = color;
  }

  // Button
  const button = document.getElementById("result-button");
  if (button) {
    button.textContent = display.toUpperCase();
    button.style.background = color;
  }

  // Details
  setText("sender", data.sender);
  setText("links", data.links);
  setText("keywords", data.content);
  setText("attachment", typeof data.attachment === "string" ? data.attachment : (data.attachment ? "Yes" : "No"));
}

function getRiskColor(label) {
  switch ((label || "").toLowerCase()) {
    case "ham": return "#28a745";       // Safe
    case "spam": return "#fd7e14";      // Risk
    case "phishing": return "#dc3545";  // High Risk
    case "support": return "#00bfff";   // Safe (Support)
    default: return "#6c757d";          // Unknown
  }
}

function labelDisplay(label) {
  switch ((label || "").toLowerCase()) {
    case "ham": return "Safe";
    case "spam": return "Risk";
    case "phishing": return "High Risk";
    case "support": return "Safe";
    default: return "Unknown";
  }
}

function resolveScore(raw) {
  let s = typeof raw === "number" ? raw : parseFloat(raw);
  if (Number.isNaN(s)) return 0.5;
  if (s > 1 && s <= 100) s = s / 100;
  return Math.max(0, Math.min(s, 1));
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = (value === null || value === undefined || value === "") ? "--" : value;
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
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      return setStatus("Failed to read email body.");
    }

    const emailText = result.value || "";
    const hasAttachment = Array.isArray(item.attachments) && item.attachments.length > 0;

    if (!emailText.trim()) {
      return setStatus("Email has no readable body text.");
    }

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
      const score = resolveScore(data.score);

      showResult({
        score,
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
