/* global Office, document */

Office.onReady(() => {
  waitForGauge(() => {
    initializeGaugeVisuals();

    // Self-check: test rendering
    showResult({
      score: 0.75,
      label: "spam",
      display: "Spam (Test)",
      sender: "debug@example.com",
      links: "2 test links",
      content: "Debug content",
      attachment: "No"
    });

    startClassification();
  });
});

function waitForGauge(callback) {
  const arc = document.getElementById("risk-arc");
  const needle = document.getElementById("needle");
  if (arc && needle) {
    callback();
  } else {
    requestAnimationFrame(() => waitForGauge(callback));
  }
}

function initializeGaugeVisuals() {
  const arc = document.getElementById("risk-arc");
  if (arc) {
    const arcLength = arc.getTotalLength();
    arc.setAttribute("stroke-dasharray", arcLength);
    arc.style.strokeDashoffset = arcLength;
    arc.dataset.arcLength = arcLength;
  }

  const needle = document.getElementById("needle");
  if (needle) {
    needle.setAttribute("transform", "rotate(-90 100 80)");
  }
}

function startClassification() {
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
        return;
      }

      classifyEmail(emailText, hasAttachment);
    } else {
      setStatus("Failed to read email body.");
    }
  });
}

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
      const label = String(data.label || "unknown").toLowerCase();
      const score = resolveScore(data.score);

      showResult({
        ...data,
        label,
        score,
        display: data.display || labelDisplay(label),
        attachment:
          typeof data.attachment === "string"
            ? data.attachment
            : data.attachment ? "Yes" : "No",
      });

      setStatus("Classification complete.");
    })
    .catch((err) => {
      console.error("Fetch error:", err);
      setStatus("Error contacting backend");
    });
}

function showResult(data) {
  const label = data.label || "unknown";
  const score = resolveScore(data.score);

  // Needle rotation
  const needle = document.getElementById("needle");
  if (needle) {
    const angle = -90 + score * 180;
    needle.setAttribute("transform", `rotate(${angle} 100 80)`);
  }

  // Arc update
  const arc = document.getElementById("risk-arc");
  if (arc) {
    const arcLength = parseFloat(arc.dataset.arcLength) || arc.getTotalLength();
    arc.style.strokeDashoffset = arcLength - score * arcLength;

    // Pick solid color based on score
    let color = "#28a745"; // green
    if (score >= 0.75) {
      color = "#dc3545"; // red
    } else if (score >= 0.67) {
      color = "#00bfff"; // blue
    } else if (score >= 0.34) {
      color = "#fd7e14"; // orange
    }
    arc.setAttribute("stroke", color);
  }

  // Update labels
  setText("score-label", data.display || labelDisplay(label));
  setText("sender", data.sender || "--");
  setText("links", data.links || "--");
  setText("keywords", data.content || "--");
  setText("attachment", data.attachment || "--");
}

function resolveScore(raw) {
  let s = typeof raw === "number" ? raw : parseFloat(raw);
  if (Number.isNaN(s)) return 0.5;
  if (s > 1 && s <= 100) s = s / 100;
  return Math.max(0, Math.min(s, 1));
}

function labelDisplay(label) {
  switch (label) {
    case "ham": return "Ham (Safe)";
    case "support": return "Support Ticket";
    case "spam": return "Spam";
    case "phishing": return "Phishing";
    default: return "Unknown";
  }
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
