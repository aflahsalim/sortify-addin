/* global Office, document */

Office.onReady(() => {
  waitForGauge(() => {
    initializeGauge();

    // Optional: test rendering
    showResult({
      score: 0.2,
      label: "ham",
      display: "Ham (Safe)",
      sender: "debug@example.com",
      links: "None",
      content: "Debug content",
      attachment: "Yes"
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
    arc.dataset.arcLength = len;
  }

  const needle = document.getElementById("needle");
  if (needle) {
    needle.setAttribute("transform", "rotate(-90 100 80)");
  }
}

function showResult(data) {
  const score = resolveScore(data.score);
  const color = getRiskColor(score);

  // Rotate needle
  const needle = document.getElementById("needle");
  if (needle) {
    const angle = -90 + score * 180;
    needle.setAttribute("transform", `rotate(${angle} 100 80)`);
  }

  // Animate arc
  const arc = document.getElementById("risk-arc");
  if (arc) {
    const len = parseFloat(arc.dataset.arcLength);
    arc.style.strokeDashoffset = len - score * len;
    arc.setAttribute("stroke", color);
  }

  // Update label
  const label = document.getElementById("score-label");
  if (label) {
    label.textContent = data.display || labelDisplay(data.label);
    label.style.color = color;
  }

  // Update analysis details
  setText("sender", data.sender);
  setText("links", data.links);
  setText("keywords", data.content);
  setText("attachment", data.attachment);
}

function getRiskColor(score) {
  if (score >= 0.75) return "#dc3545"; // red
  if (score >= 0.67) return "#00bfff"; // blue
  if (score >= 0.34) return "#fd7e14"; // orange
  return "#28a745"; // green
}

function resolveScore(raw) {
  let s = typeof raw === "number" ? raw : parseFloat(raw);
  if (isNaN(s)) return 0.5;
  if (s > 1 && s <= 100) s /= 100;
  return Math.max(0, Math.min(s, 1));
}

function labelDisplay(label) {
  switch ((label || "").toLowerCase()) {
    case "ham": return "Ham (Safe)";
    case "support": return "Support Ticket";
    case "spam": return "Spam";
    case "phishing": return "Phishing";
    default: return "Unknown";
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
    .then(res => res.ok ? res.json() : res.text().then(body => { throw new Error(body); }))
    .then(data => {
      showResult({
        score: resolveScore(data.score),
        label: data.label,
        display: data.display || labelDisplay(data.label),
        sender: data.sender,
        links: data.links,
        content: data.content,
        attachment: data.attachment ? "Yes" : "No"
      });
      setStatus("Classification complete.");
    })
    .catch(err => {
      console.error("Error:", err);
      setStatus("Error contacting backend.");
    });
}
