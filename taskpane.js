/* global Office, document */

Office.onReady(() => {
  // Verify DOM is present before doing anything
  const ready = ensureDom();
  if (!ready) {
    setStatus("UI not ready");
    return;
  }

  // Prepare visuals so first update animates
  initializeGaugeVisuals();

  const item = Office.context?.mailbox?.item;
  if (!item) {
    console.warn("[Sortify] No mailbox item found.");
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
        console.warn("[Sortify] Empty email body.");
        setStatus("Email has no readable body text.");
        showResult({
          score: 0.5, // default to mid so arc is visible
          label: "support",
          display: "Support Ticket",
          sender: "--",
          links: "--",
          content: "No content",
          attachment: hasAttachment ? "Yes" : "No",
        });
        return;
      }

      classifyEmail(emailText, hasAttachment);
    } else {
      console.error("[Sortify] Failed to read email body.", result.error);
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
      console.log("[Sortify] Backend response:", data);

      // Normalize data
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
      console.error("[Sortify] Fetch error:", err);
      setStatus("Error contacting backend");
      // Show a visible default so UI isn't empty
      showResult({
        score: 0.6,
        label: "spam",
        display: "Spam",
        sender: "--",
        links: "--",
        content: "--",
        attachment: "No",
      });
    });
}

function showResult(data) {
  const label = data.label || "unknown";
  const score = resolveScore(data.score);
  const percent = `${Math.round(score * 100)}%`;

  // Needle animation
  const needle = document.getElementById("needle");
  if (needle) {
    needle.style.transition = "transform 0.9s cubic-bezier(0.22, 1, 0.36, 1)";
    needle.setAttribute("transform", `rotate(${angleFor(label)} 100 100)`);
  } else {
    console.warn("[Sortify] Needle not found.");
  }

  // Gradient colors per classification
  const palette = {
    green: "#28a745",
    orange: "#fd7e14",
    red: "#dc3545",
    blue: "#007bff",
    gray: "#6c757d",
  };
  const { g1, g2, g3, fallback } = gradientFor(label, palette);

  const s1 = document.getElementById("grad-stop-1");
  const s2 = document.getElementById("grad-stop-2");
  const s3 = document.getElementById("grad-stop-3");
  if (s1 && s2 && s3) {
    s1.setAttribute("stop-color", g1);
    s2.setAttribute("stop-color", g2);
    s3.setAttribute("stop-color", g3);
  } else {
    console.warn("[Sortify] Gradient stops not found, using solid fallback.");
  }

  // Arc animation + gradient application
  const arc = document.getElementById("risk-arc");
  if (arc) {
    // Always reapply gradient so the browser refreshes the paint
    arc.setAttribute("stroke", "url(#arcGradient)");

    // If gradient stops aren’t present, force a visible solid color
    if (!(s1 && s2 && s3)) {
      arc.setAttribute("stroke", fallback);
    }

    // Ensure transitions exist (in case CSS didn't load yet)
    arc.style.transition =
      "stroke 0.5s ease-in-out, stroke-dashoffset 0.9s cubic-bezier(0.22, 1, 0.36, 1)";

    // Force an animation by starting hidden, then animating to target
    const maxArc = 283;
    const target = maxArc - score * maxArc;

    // If target equals current, nudge by 0.1 to force a visible change
    const current = parseFloat(getComputedStyle(arc).strokeDashoffset) || maxArc;
    const needsNudge = Math.abs(current - target) < 0.1;

    if (needsNudge) {
      arc.style.strokeDashoffset = `${Math.max(0, Math.min(maxArc, target + 0.1))}`;
      requestAnimationFrame(() => {
        arc.style.strokeDashoffset = `${target}`;
      });
    } else {
      arc.style.strokeDashoffset = `${target}`;
    }
  } else {
    console.error("[Sortify] risk-arc not found.");
  }

  // Labels
  setText("score-label", data.display || labelDisplay(label));
  setText("score-value", percent);

  // Status badge
  const badge = document.getElementById("status");
  if (badge) {
    badge.textContent = data.display || labelDisplay(label);
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

// Ensure initial visual state so first update animates
function initializeGaugeVisuals() {
  const arc = document.getElementById("risk-arc");
  const needle = document.getElementById("needle");

  if (arc) {
    arc.setAttribute("stroke", "url(#arcGradient)");
    arc.style.transition = "none";
    arc.style.strokeDashoffset = "283"; // start hidden
    // Force paint
    void arc.getBoundingClientRect();
  }

  if (needle) {
    needle.style.transition = "none";
    needle.setAttribute("transform", "rotate(-90 100 100)");
    // Force paint
    void needle.getBoundingClientRect();
  }
}

// Verify DOM is present so JS updates actually hit elements
function ensureDom() {
  const svg = document.getElementById("gauge-svg");
  const arc = document.getElementById("risk-arc");
  const needle = document.getElementById("needle");
  const s1 = document.getElementById("grad-stop-1");
  const s2 = document.getElementById("grad-stop-2");
  const s3 = document.getElementById("grad-stop-3");

  const ok = !!(svg && arc && needle && s1 && s2 && s3);
  if (!ok) {
    console.error("[Sortify] Missing UI elements:", {
      svg: !!svg,
      arc: !!arc,
      needle: !!needle,
      stop1: !!s1,
      stop2: !!s2,
      stop3: !!s3,
    });
  }
  return ok;
}

// Helpers

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

function resolveScore(raw) {
  // Accept 0–1; if 0–100, convert; default to 0.5 so arc is visible
  let s = typeof raw === "number" ? raw : parseFloat(raw);
  if (Number.isNaN(s)) return 0.5;
  if (s > 1 && s <= 100) s = s / 100;
  return Math.max(0, Math.min(s, 1));
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
