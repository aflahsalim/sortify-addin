// ── Config ────────────────────────────────────────────────────────
const BACKEND = "https://sortify-backend-hwf9d0exgqdub9cn.canadacentral-01.azurewebsites.net";

// Get your free key at console.cloud.google.com → Enable "Safe Browsing API"
const SAFE_BROWSING_KEY = "YOUR_GOOGLE_SAFE_BROWSING_API_KEY";

// Attachment extensions considered risky
const RISKY_EXT = ["exe","bat","vbs","js","msi","ps1","cmd","scr","zip","rar","docm","xlsm"];

// Phrases that indicate urgency / social engineering
const URGENCY_PHRASES = [
  "urgent","immediately","suspended","verify","confirm your",
  "click now","24 hours","48 hours","act now","account locked",
  "unusual activity","security alert","update your password"
];

// Holds current email data for the report button
let currentScanData = null;

// ── Entry point ───────────────────────────────────────────────────
Office.onReady(() => {
  waitForDOM(() => {

    // Try immediately if email already selected
    if (Office.context.mailbox.item) {
      startClassification();
    } else {
      setStatus("Click an email to scan", "");
    }

    // Re-run every time user clicks a different email
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => {
        if (Office.context.mailbox.item) {
          resetUI();
          startClassification();
        }
      }
    );

  });
});

// Wait until DOM elements exist before doing anything
function waitForDOM(callback) {
  if (document.getElementById("risk-arc")) callback();
  else requestAnimationFrame(() => waitForDOM(callback));
}

// Reset all fields back to dashes when switching emails
function resetUI() {
  setStatus("Scanning...", "");
  ["sender","links","attachment","urgency"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = "—";
  });
  ["sender-dot","links-dot","attachment-dot","urgency-dot"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.className = "dot";
  });
  // Reset gauge
  const arc = document.getElementById("risk-arc");
  if (arc) arc.style.strokeDashoffset = "251";
  const needle = document.getElementById("needle");
  if (needle) needle.style.transform = "rotate(-90deg)";
  const label = document.getElementById("score-label");
  if (label) label.textContent = "—";
  const badge = document.getElementById("result-button");
  if (badge) { badge.textContent = "—"; badge.className = "risk-badge"; }
  // Reset report button
  const btn = document.getElementById("report-btn");
  if (btn) { btn.disabled = false; btn.textContent = "Send to Sortify team"; }
  currentScanData = null;
}

// ── Main classification flow ──────────────────────────────────────
function startClassification() {
  const item = Office.context?.mailbox?.item;
  if (!item) return setStatus("No email selected", "");

  setStatus("Scanning...", "");

  // Read plain-text body of the email
  item.body.getAsync(Office.CoercionType.Text, async (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded)
      return setStatus("Could not read email", "warn");

    const body = result.value || "";

    // ── 1. Sender reputation ──────────────────────────────────────
    let senderEmail = "";
    try { senderEmail = item?.from?.emailAddress?.address || ""; } catch(e) {}
    const senderDomain = senderEmail.split("@")[1] || "";

    // Free/personal domains are suspicious for official-looking emails
    const freeDomains = ["gmail.com","yahoo.com","hotmail.com","outlook.com"];
    const senderIsFree = freeDomains.includes(senderDomain.toLowerCase());
    const senderLabel  = senderIsFree ? "Free domain" : (senderDomain || "Unknown");
    const senderRisk   = senderIsFree ? "warn" : "safe";

    // ── 2. Auth headers — SPF / DKIM ─────────────────────────────
    let authLabel = "Unavailable";
    let authRisk  = "warn";
    try {
      await new Promise(resolve => {
        item.internetHeaders?.getAsync(["Authentication-Results"], (r) => {
          if (r.status === Office.AsyncResultStatus.Succeeded) {
            const h = (r.value?.["Authentication-Results"] || "").toLowerCase();
            const spf  = h.includes("spf=pass");
            const dkim = h.includes("dkim=pass");
            const passed = [spf, dkim].filter(Boolean).length;
            if (passed === 2)      { authLabel = "SPF + DKIM pass"; authRisk = "safe"; }
            else if (passed === 1) { authLabel = "Partial pass";    authRisk = "warn"; }
            else                   { authLabel = "Auth failed";     authRisk = "danger"; }
          }
          resolve();
        });
      });
    } catch(e) { /* not available in all Outlook versions */ }

    // ── 3. File attachment check ──────────────────────────────────
    const attachments = Array.isArray(item.attachments) ? item.attachments : [];
    let filesLabel = "None";
    let filesRisk  = "safe";
    if (attachments.length > 0) {
      const risky = attachments.find(a => {
        const ext = (a.name || "").split(".").pop().toLowerCase();
        return RISKY_EXT.includes(ext);
      });
      if (risky) {
        const ext  = risky.name.split(".").pop().toUpperCase();
        filesLabel = `${ext} file — risky`;
        filesRisk  = "danger";
      } else {
        filesLabel = `${attachments.length} safe file(s)`;
        filesRisk  = "safe";
      }
    }

    // ── 4. Urgency phrase scan ────────────────────────────────────
    const bodyLower = body.toLowerCase();
    const matched   = URGENCY_PHRASES.filter(p => bodyLower.includes(p));
    let urgencyLabel = "None detected";
    let urgencyRisk  = "safe";
    if (matched.length >= 3)      { urgencyLabel = `${matched.length} signals — High`; urgencyRisk = "danger"; }
    else if (matched.length >= 1) { urgencyLabel = `"${matched[0]}"`;                  urgencyRisk = "warn"; }

    // ── Update the 4 detail rows ──────────────────────────────────
    setField("sender",     senderLabel,  "sender-dot",     senderRisk);
    setField("links",      authLabel,    "links-dot",      authRisk);
    setField("attachment", filesLabel,   "attachment-dot", filesRisk);
    setField("urgency",    urgencyLabel, "urgency-dot",    urgencyRisk);

    // Save for report button
    currentScanData = {
      sender:           senderEmail,
      subject:          item.subject || "",
      sender_risk:      senderRisk,
      auth_result:      authLabel,
      files_result:     filesLabel,
      urgency_result:   urgencyLabel,
      attachment_count: attachments.length,
      body_preview:     body.substring(0, 300)
    };

    // ── Call ML backend ───────────────────────────────────────────
    await classifyWithBackend(body, attachments.length > 0);
  });
}

// ── Backend ML call ───────────────────────────────────────────────
async function classifyWithBackend(bodyText, hasAttachment) {
  // Abort if backend takes more than 10 seconds
  const controller = new AbortController();
  const timeout    = setTimeout(() => controller.abort(), 10000);

  try {
    const res = await fetch(`${BACKEND}/classify`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: bodyText, attachment: hasAttachment ? "Yes" : "No" }),
      signal: controller.signal
    });
    clearTimeout(timeout);
    const data  = await res.json();
    const label = (data.label || "unknown").toLowerCase();
    renderGauge(label);
    logScan(label);

  } catch(e) {
    // Backend unreachable — fall back to local rules
    clearTimeout(timeout);
    setStatus("Local analysis", "warn");
    const bodyLower   = bodyText.toLowerCase();
    const urgentCount = URGENCY_PHRASES.filter(p => bodyLower.includes(p)).length;
    const hasLinks    = /https?:\/\//i.test(bodyText);
    let label = "ham";
    if (hasLinks && urgentCount >= 3) label = "phishing";
    else if (urgentCount >= 2)        label = "spam";
    else if (hasLinks && urgentCount) label = "spam";
    renderGauge(label);
    logScan(label);
  }
}

// ── Gauge renderer ────────────────────────────────────────────────
function renderGauge(label) {
  const map = {
    ham:      { angle: -90, fill: 0.0,  text: "SAFE",     cls: "safe",   status: "All clear",  sCls: "done"   },
    support:  { angle: -45, fill: 0.25, text: "SUPPORT",  cls: "info",   status: "Low risk",   sCls: "done"   },
    spam:     { angle:  45, fill: 0.75, text: "SPAM",     cls: "warn",   status: "Suspicious", sCls: "warn"   },
    phishing: { angle:  90, fill: 1.0,  text: "PHISHING", cls: "danger", status: "High risk",  sCls: "danger" },
  };
  const c = map[label] || { angle: 0, fill: 0.5, text: "UNKNOWN", cls: "", status: "Scanned", sCls: "done" };

  // Rotate needle
  const needle = document.getElementById("needle");
  if (needle) needle.style.transform = `rotate(${c.angle}deg)`;

  // Fill arc (total arc path length = 251)
  const arc = document.getElementById("risk-arc");
  if (arc) arc.style.strokeDashoffset = 251 - c.fill * 251;

  // Update text labels
  const scoreLabel = document.getElementById("score-label");
  if (scoreLabel) scoreLabel.textContent = c.text;

  const badge = document.getElementById("result-button");
  if (badge) { badge.textContent = c.text; badge.className = `risk-badge ${c.cls}`; }

  setStatus(c.status, c.sCls);

  if (currentScanData) currentScanData.label = label;
}

// ── UI helpers ────────────────────────────────────────────────────

function setField(valueId, value, dotId, risk) {
  const el = document.getElementById(valueId);
  if (el) el.textContent = value || "—";
  const dot = document.getElementById(dotId);
  if (dot) dot.className = `dot ${risk}`;
}

function setStatus(msg, cls) {
  const pill = document.getElementById("status");
  if (!pill) return;
  pill.textContent = msg;
  pill.className   = `status-pill ${cls}`;
}

// ── Report button ─────────────────────────────────────────────────

function reportEmail() {
  if (!currentScanData) return;
  document.getElementById("confirm-overlay").classList.remove("hidden");
}

function closeConfirm() {
  document.getElementById("confirm-overlay").classList.add("hidden");
}

async function confirmReport() {
  closeConfirm();
  const btn    = document.getElementById("report-btn");
  btn.disabled = true;
  btn.textContent = "Sending...";
  try {
    await fetch(`${BACKEND}/report`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...currentScanData, reported: true })
    });
    btn.textContent = "✓ Reported";
  } catch(e) {
    btn.textContent = "Failed — try again";
    btn.disabled    = false;
  }
}

// ── Silent background logging ─────────────────────────────────────

// Every scan is silently logged to the admin dashboard
async function logScan(label) {
  if (!currentScanData) return;
  try {
    await fetch(`${BACKEND}/log-scan`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...currentScanData, label })
    });
  } catch(e) { /* fail silently */ }
}
