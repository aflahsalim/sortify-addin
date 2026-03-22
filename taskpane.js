// ── Config ────────────────────────────────────────────────────────
const BACKEND = "https://sortify-backend-hwf9d0exgqdub9cn.canadacentral-01.azurewebsites.net";
const RISKY_EXT = ["exe","bat","vbs","js","msi","ps1","cmd","scr","zip","rar","docm","xlsm"];
const URGENCY_PHRASES = [
  "urgent","immediately","suspended","verify","confirm your",
  "click now","24 hours","48 hours","act now","account locked",
  "unusual activity","security alert","update your password"
];
let currentScanData = null;

// ── Entry point ───────────────────────────────────────────────────
Office.onReady(() => {
  waitForDOM(() => {
    if (Office.context?.mailbox?.item) {
      startClassification();
    } else {
      setStatus("Click an email to scan", "");
    }
    // Re-scan when user clicks a different email
    Office.context?.mailbox?.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => { resetUI(); startClassification(); }
    );
  });
});

function waitForDOM(callback) {
  if (document.getElementById("risk-arc")) callback();
  else requestAnimationFrame(() => waitForDOM(callback));
}

// ── Reset UI between emails ───────────────────────────────────────
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
  const arc = document.getElementById("risk-arc");
  if (arc) arc.style.strokeDashoffset = "251";
  const needle = document.getElementById("needle");
  if (needle) needle.style.transform = "rotate(-90deg)";
  const lbl = document.getElementById("score-label");
  if (lbl) lbl.textContent = "—";
  const badge = document.getElementById("result-button");
  if (badge) { badge.textContent = "—"; badge.className = "risk-badge"; }
  const btn = document.getElementById("report-btn");
  if (btn) { btn.disabled = false; btn.textContent = "Send to Sortify team"; }
  currentScanData = null;
}

// ── Step 1: read email body ───────────────────────────────────────
function startClassification() {
  const item = Office.context?.mailbox?.item;
  if (!item) return setStatus("No email selected", "");
  setStatus("Scanning...", "");

  // NOT async — plain callback to avoid Office.js async issues
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded)
      return setStatus("Could not read email", "warn");
    processEmail(item, result.value || "");
  });
}

// ── Step 2: run all checks then call backend ──────────────────────
function processEmail(item, body) {

  // 1. Sender
  let senderEmail = "";
  try { senderEmail = item?.from?.emailAddress?.address || ""; } catch(e) {}
  const domain = senderEmail.split("@")[1] || "";
  const freeDomains = ["gmail.com","yahoo.com","hotmail.com","outlook.com"];
  const isFree = freeDomains.includes(domain.toLowerCase());
  const senderLabel = isFree ? "Free domain" : (domain || "Unknown");
  const senderRisk  = isFree ? "warn" : "safe";

  // 2. Attachments (sync — no async needed)
  const attachments = Array.isArray(item.attachments) ? item.attachments : [];
  let filesLabel = "None";
  let filesRisk  = "safe";
  if (attachments.length > 0) {
    const risky = attachments.find(a =>
      RISKY_EXT.includes((a.name || "").split(".").pop().toLowerCase())
    );
    if (risky) {
      filesLabel = risky.name.split(".").pop().toUpperCase() + " — risky";
      filesRisk  = "danger";
    } else {
      filesLabel = attachments.length + " safe file(s)";
      filesRisk  = "safe";
    }
  }

  // 3. Urgency phrases (sync)
  const bodyLower = body.toLowerCase();
  const matched   = URGENCY_PHRASES.filter(p => bodyLower.includes(p));
  let urgencyLabel = "None detected";
  let urgencyRisk  = "safe";
  if (matched.length >= 3)      { urgencyLabel = matched.length + " signals — High"; urgencyRisk = "danger"; }
  else if (matched.length >= 1) { urgencyLabel = '"' + matched[0] + '"';             urgencyRisk = "warn"; }

  // 4. Auth headers (async but non-blocking — update UI after)
  let authLabel = "Unavailable";
  let authRisk  = "warn";
  try {
    item.internetHeaders?.getAsync(["Authentication-Results"], (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        const h = (r.value?.["Authentication-Results"] || "").toLowerCase();
        const spf  = h.includes("spf=pass");
        const dkim = h.includes("dkim=pass");
        const n    = [spf, dkim].filter(Boolean).length;
        if (n === 2)      { authLabel = "SPF + DKIM pass"; authRisk = "safe"; }
        else if (n === 1) { authLabel = "Partial pass";    authRisk = "warn"; }
        else              { authLabel = "Auth failed";     authRisk = "danger"; }
      }
      // Update auth row after it comes back
      setField("links", authLabel, "links-dot", authRisk);
      if (currentScanData) currentScanData.auth_result = authLabel;
    });
  } catch(e) {}

  // Update UI immediately with what we have (auth updates async above)
  setField("sender",     senderLabel,  "sender-dot",     senderRisk);
  setField("links",      "Checking…",  "links-dot",      "warn");
  setField("attachment", filesLabel,   "attachment-dot", filesRisk);
  setField("urgency",    urgencyLabel, "urgency-dot",    urgencyRisk);

  // Save for report button
  currentScanData = {
    sender:           senderEmail,
    subject:          item.subject || "",
    label:            "unknown",
    sender_risk:      senderRisk,
    auth_result:      authLabel,
    files_result:     filesLabel,
    urgency_result:   urgencyLabel,
    attachment_count: attachments.length,
    body_preview:     body.substring(0, 300)
  };

  // Call ML backend
  callBackend(body, attachments.length > 0);
}

// ── Step 3: ML backend call ───────────────────────────────────────
function callBackend(bodyText, hasAttachment) {
  const controller = new AbortController();
  const timeout    = setTimeout(() => controller.abort(), 10000);

  fetch(BACKEND + "/classify", {
    method:  "POST",
    headers: { "Content-Type": "application/json" },
    body:    JSON.stringify({ text: bodyText, attachment: hasAttachment ? "Yes" : "No" }),
    signal:  controller.signal
  })
  .then(res => res.json())
  .then(data => {
    clearTimeout(timeout);
    const label = (data.label || "unknown").toLowerCase();
    renderGauge(label);
    logScan(label);
  })
  .catch(() => {
    // Backend unreachable — local fallback
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
  });
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

  const needle = document.getElementById("needle");
  if (needle) needle.style.transform = "rotate(" + c.angle + "deg)";

  const arc = document.getElementById("risk-arc");
  if (arc) arc.style.strokeDashoffset = 251 - c.fill * 251;

  const scoreLabel = document.getElementById("score-label");
  if (scoreLabel) scoreLabel.textContent = c.text;

  const badge = document.getElementById("result-button");
  if (badge) { badge.textContent = c.text; badge.className = "risk-badge " + c.cls; }

  setStatus(c.status, c.sCls);
  if (currentScanData) currentScanData.label = label;
}

// ── UI helpers ────────────────────────────────────────────────────
function setField(valueId, value, dotId, risk) {
  const el = document.getElementById(valueId);
  if (el) el.textContent = value || "—";
  const dot = document.getElementById(dotId);
  if (dot) dot.className = "dot " + risk;
}

function setStatus(msg, cls) {
  const pill = document.getElementById("status");
  if (!pill) return;
  pill.textContent = msg;
  pill.className   = "status-pill " + cls;
}

// ── Report button ─────────────────────────────────────────────────
function reportEmail() {
  if (!currentScanData) return;
  document.getElementById("confirm-overlay").classList.remove("hidden");
}

function closeConfirm() {
  document.getElementById("confirm-overlay").classList.add("hidden");
}

function confirmReport() {
  closeConfirm();
  const btn = document.getElementById("report-btn");
  btn.disabled    = true;
  btn.textContent = "Sending...";
  fetch(BACKEND + "/report", {
    method:  "POST",
    headers: { "Content-Type": "application/json" },
    body:    JSON.stringify(Object.assign({}, currentScanData, { reported: true }))
  })
  .then(() => { btn.textContent = "✓ Reported"; })
  .catch(() => { btn.textContent = "Failed — try again"; btn.disabled = false; });
}

// ── Silent background log ─────────────────────────────────────────
function logScan(label) {
  if (!currentScanData) return;
  fetch(BACKEND + "/log-scan", {
    method:  "POST",
    headers: { "Content-Type": "application/json" },
    body:    JSON.stringify(Object.assign({}, currentScanData, { label }))
  }).catch(() => {});
}
