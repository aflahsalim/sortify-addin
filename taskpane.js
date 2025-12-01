/* global Office, document */

Office.onReady(() => {
  waitForGauge(() => {
    initializeGaugeVisuals();

    // 🔧 Self-check mode: force a visible arc immediately
    console.log("🔧 Self-check: forcing arc to 75% fill");
    showResult({
      score: 0.75, // hardcoded test value
      label: "spam",
      display: "Spam (Test)",
      sender: "debug@example.com",
      links: "2 test links",
      content: "Debug content",
      attachment: "No"
    });

    // After test, continue with real classification
    startClassification();
  });
});

function waitForGauge(callback) {
  const arc = document.getElementById("risk-arc");
  const needle = document.getElementById("needle");
  const stops = [
    document.getElementById("grad-stop-1"),
    document.getElementById("grad-stop-2"),
    document.getElementById("grad-stop-3"),
  ];
  if (arc && needle && stops.every(Boolean)) {
    callback();
  } else {
    requestAnimationFrame(() => waitForGauge(callback));
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
      console.log("✅ Backend response:", data);
      const label = String(data.label || "unknown").toLowerCase();
      const score = resolveScore(data.score);
      console.log("✅ Resolved score:", score);

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
  const percent = `${Math.round(score * 100)}%`;

  console.log("🔧 showResult called with:", { label, score, percent });

  const needle = document.getElementById("needle");
  if (needle) {
    needle.style.transition = "transform 0.9s cubic-bezier(0.22, 1, 0.36, 1)";
    needle.setAttribute("transform", `rotate(${angleFor(label)} 100 100)`);
  }

  const arc = document.getElementById("risk-arc");
  if (arc) {
    const maxArc = 283;
    arc.setAttribute("stroke", "url(#arcGradient)");
    arc.style.transition =
      "stroke-dashoffset 0.9s cubic-bezier(0.22, 1, 0.36, 1), stroke 0.5s ease-in-out";
    arc.style.strokeDashoffset = `${maxArc - score * maxArc}`;
    console.log("🔧 Arc updated:", arc.style.strokeDashoffset);
  }

  setText("score-label", data.display || labelDisplay(label));
  setText("score-value", percent);
}

function initializeGaugeVisuals() {
  const arc = document.getElementById("risk-arc");
  const needle = document.getElementById("needle");
  if (arc) {
    arc.setAttribute("stroke", "url(#arcGradient)");
    arc.style.transition = "none";
    arc.style.strokeDashoffset = "283";
  }
  if (needle) {
    needle.style.transition = "none";
    needle.setAttribute("transform", "rotate(-90 100 100)");
  }
}

function resolveScore(raw) {
  let s = typeof raw === "number" ? raw : parseFloat(raw);
  if (Number.isNaN(s)) return 0.5;
  if (s > 1 && s <= 100) s = s / 100;
  return Math.max(0, Math.min(s, 1));
}

function angleFor(label) {
  const angleMap = {
    ham: -90,
    support: -45,
    spam: 45,
    phishing: 90,
    unknown: -90,
  };
  return angleMap[label] ?? -90;
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

function gradientFor(label, palette) {
  switch (label) {
    case "ham":
      return { g1: palette.green, g2: "#4bd07e", g3: "#7be0a3", fallback: palette.green };
    case "support":
      return { g1: palette.blue, g2: "#2fa8ff", g3: palette.green, fallback: palette.blue };
    case "spam":
      return { g1: palette.orange, g2: "#ff9a3b", g3: palette.red, fallback: palette.orange };
    case "phishing":
      return { g1: "#ff6b6b", g2: palette.red, g3: "#b00020", fallback: palette.red };
    default:
      return { g1: palette.gray, g2: "#8a8f94", g3: "#b0b5bb", fallback: palette.gray };
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
