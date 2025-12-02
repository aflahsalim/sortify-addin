Office.onReady(() => {
  waitForGauge(() => {
    initializeGauge();
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
  if (needle) needle.setAttribute("transform", "rotate(-90 100 90)");
}

function showResult(data) {
  const label = (data.label || "unknown").toLowerCase();
  const color = getRiskColor(label);
  const angle = getFixedAngle(label);
  const fillRatio = getFillRatio(label);

  // Needle + arc update
  const needle = document.getElementById("needle");
  if (needle) needle.setAttribute("transform", `rotate(${angle} 100 90)`);

  const arc = document.getElementById("risk-arc");
  if (arc) {
    const len = parseFloat(arc.dataset.arcLength || arc.getTotalLength());
    arc.style.strokeDashoffset = len - fillRatio * len;
    arc.setAttribute("stroke", color);
  }

  // Gauge label
  const labelEl = document.getElementById("score-label");
  if (labelEl) {
    labelEl.textContent = label.toUpperCase();
    labelEl.style.color = color;
  }

  // Badge
  const badge = document.getElementById("result-button");
  if (badge) {
    badge.textContent = getBadgeText(label);
    badge.style.background = bubbleColor(label);
    badge.style.color = "#000";
  }

  // Analysis details
  setText("sender", formatOrigin(data.sender));
  setText("links", formatPresence(data.links));
  setText("attachment", formatPresence(data.attachment));
  setText("urgency", formatUrgency(label));
}

/* --- Professional formatting helpers --- */

function formatOrigin(value) {
  const val = String(value || "").toLowerCase();
  if (val.includes("trusted")) return "Trusted";
  if (val.includes("suspicious")) return "Suspicious";
  return "Unknown";
}

function formatPresence(value) {
  const val = String(value || "").toLowerCase();
  if (val.includes("yes") || val.includes("link") || val.includes("present")) {
    return "Detected";
  }
  return "None";
}

function formatUrgency(label) {
  switch (label) {
    case "phishing": return "Critical";
    case "spam": return "Elevated";
    case "ham":
    case "support": return "Normal";
    default: return "Unknown";
  }
}

/* --- Badge text logic --- */
function getBadgeText(label) {
  switch (label) {
    case "ham":
    case "support":
      return "SAFE";
    case "spam":
      return "RISK";
    case "phishing":
      return "HIGH RISK";
    default:
      return "UNKNOWN";
  }
}

function getRiskColor(label) {
  switch (label) {
    case "ham": return "#28a745";
    case "support": return "#00bfff";
    case "spam": return "#fd7e14";
    case "phishing": return "#dc3545";
    default: return "#6c757d";
  }
}

function bubbleColor(label) {
  switch (label) {
    case "ham": return "#9ff08c";
    case "support": return "#8fd5ff";
    case "spam": return "#ffdd57";
    case "phishing": return "#ff9aa2";
    default: return "#d0d3d8";
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
  if (!el) return;
  el.textContent = value || "--";
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

    const linkRegex = /(https?:\/\/[^\s]+)/gi;
    const hasLinks = linkRegex.test(emailText);

    classifyEmail(emailText, hasAttachment, hasLinks, item);
  });
}

function classifyEmail(emailText, hasAttachment, hasLinks, item) {
  setStatus("Classifying email...");

  const senderEmail = item?.from?.emailAddress?.address || "";
  const senderDomain = senderEmail.split("@")[1] || "";
  const isFreeDomain = ["gmail.com", "yahoo.com", "outlook.com", "hotmail.com"].includes(senderDomain?.toLowerCase());
  const senderReputation = senderEmail
    ? (isFreeDomain ? "Suspicious" : "Trusted")
    : "Unknown";

  fetch("https://sortify-y7ru.onrender.com/classify", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      text: emailText || "",
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
        sender: senderReputation,
        links: hasLinks ? "Present" : "Absent",
        attachment: hasAttachment ? "Present" : "Absent"
      });

      setStatus("Classification complete.");
    })
    .catch((err) => {
      console.error("Backend error:", err);
      setStatus("Error contacting backend.");
    });
}
