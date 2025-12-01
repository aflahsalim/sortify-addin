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
  const score = Math.max(0, Math.min(Number(data.score) || 0, 1)); // clamp between 0–1

  // Needle angle by category
  const angleMap = {
    ham: -90,
    support: -45,
    spam: 45,
    phishing: 90,
    unknown: -90,
  };
  const needleAngle = angleMap[label] ?? -90;

  const needle = document.getElementById("needle");
  if (needle) {
    needle.setAttribute("transform", `rotate(${needleAngle} 100 100)`);
  }

  // Gradient color stops
  const palette = {
    green: "#28a745",
    orange: "#fd7e14",
    red: "#dc3545",
    blue: "#007bff",
    gray: "#6c757d",
  };

  let g1 = palette.green, g2 = palette.orange, g3 = palette.red;
  if (label === "ham") {
    g1 = palette.green; g2 = "#4bd07e"; g3 = "#7be0a3";
  } else if (label === "support") {
    g1 = palette.blue; g2 = "#2fa8ff"; g3 = palette.green;
  } else if (label === "spam") {
    g1 = palette.orange; g2 = "#ff9a3b"; g3 = palette.red;
  } else if (label === "phishing") {
    g1 = "#ff6b6b"; g2 = palette.red; g3 = "#b00020";
  } else {
    g1 = palette.gray; g2 = "#8a8f94"; g3 = "#b0b5bb";
  }

  ["grad-stop-1", "grad-stop-2", "grad-stop-3"].forEach((id, i) => {
    const stop = document.getElementById(id);
    if (stop) {
      stop.setAttribute("stop-color", [g1, g2, g3][i]);
    }
  });

  // Arc fill animation
  const arc = document.getElementById("risk-arc");
  if (arc) {
    const maxArc = 283;
    arc.style.strokeDashoffset = maxArc - (score * maxArc);
  }

  // Update labels
  const confidencePercent = `${Math.round(score * 100)}%`;
  setText("score-label", data.display || label.toUpperCase());
  setText("score-value", confidencePercent);
  setText("confidence", `Confidence: ${confidencePercent}`);

  // Status badge
  const badge = document.getElementById("status");
  if (badge) {
    badge.textContent = data.display || label.toUpperCase();
    badge.className = "status-badge";
    if (label === "phishing") badge.classList.add("status-spam");
    else if (label === "spam") badge.classList.add("status-medium");
    else if (label === "support") badge.classList.add("status-support");
    else badge.classList.add("status-safe");
  }

  // Analysis details
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
