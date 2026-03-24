// ── Config ────────────────────────────────────────────────────────
const BACKEND = "https://sortify-backend-hwf9d0exgqdub9cn.canadacentral-01.azurewebsites.net";
const RISKY_EXT = ["exe","bat","vbs","js","msi","ps1","cmd","scr","zip","rar","docm","xlsm"];
const URGENCY_PHRASES = [
  "urgent","immediately","suspended","verify","confirm your",
  "click now","24 hours","48 hours","act now","account locked",
  "unusual activity","security alert","update your password",
  "will be terminated","will be suspended","limited time"
];

// Arc gradient colours per risk
const RISK_COLORS = {
  ham:      { start: "#00e87a", end: "#00b85f" },
  support:  { start: "#06b6d4", end: "#0891b2" },
  spam:     { start: "#f59e0b", end: "#d97706" },
  phishing: { start: "#ef4444", end: "#dc2626" },
};

let currentScanData = null;
let scanStartTime   = null;

// ── Entry point ───────────────────────────────────────────────────
Office.onReady(() => {
  waitForDOM(() => {
    if (Office.context?.mailbox?.item) {
      startClassification();
    } else {
      setStatus("Open an email to scan", "");
    }
    Office.context?.mailbox?.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => { resetUI(); startClassification(); }
    );
  });
});

function waitForDOM(cb) {
  if (document.getElementById("risk-arc")) cb();
  else requestAnimationFrame(() => waitForDOM(cb));
}

// ── Reset ─────────────────────────────────────────────────────────
function resetUI() {
  setStatus("Scanning...", "");
  setScanTime("");
  ["sender","links","attachment","urgency"].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.textContent = "—"; el.className = "check-value"; }
  });
  ["row-trust","row-links","row-urgency","row-files"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.className = "check-row";
  });
  const arc = document.getElementById("risk-arc");
  if (arc) arc.style.strokeDashoffset = "251";
  const needle = document.getElementById("needle");
  if (needle) needle.style.transform = "rotate(-90deg)";
  const lbl = document.getElementById("score-label");
  if (lbl) { lbl.textContent = "—"; lbl.style.color = ""; }
  const badge = document.getElementById("result-button");
  if (badge) { badge.textContent = "—"; badge.className = "risk-badge"; }
  const card = document.getElementById("gauge-card");
  if (card) card.style.borderColor = "";
  const btn = document.getElementById("report-btn");
  if (btn) { btn.disabled = false; btn.textContent = "Report to Sortify"; }
  setArcColor("ham");
  currentScanData = null;
}

// ── Step 1: read body ─────────────────────────────────────────────
function startClassification() {
  const item = Office.context?.mailbox?.item;
  if (!item) return setStatus("Open an email to scan", "");
  setStatus("Scanning...", "");
  scanStartTime = Date.now();
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded)
      return setStatus("Could not read email", "warn");
    processEmail(item, result.value || "");
  });
}

// ── Step 2: all checks ────────────────────────────────────────────
function processEmail(item, body) {

  // ── Sender domain
  let senderEmail = "";
  try { senderEmail = item?.from?.emailAddress?.address || ""; } catch(e) {}
  const domain     = senderEmail.split("@")[1] || "";
  const freeDomains = ["gmail.com","yahoo.com","hotmail.com","outlook.com","live.com"];
  const isFree     = freeDomains.includes(domain.toLowerCase());

  // ── Attachments
  const attachments = Array.isArray(item.attachments) ? item.attachments : [];
  let filesLabel = "No attachments", filesRisk = "safe";
  if (attachments.length > 0) {
    const risky = attachments.find(a =>
      RISKY_EXT.includes((a.name || "").split(".").pop().toLowerCase())
    );
    if (risky) {
      filesLabel = "⚠ " + risky.name.split(".").pop().toUpperCase() + " file — risky";
      filesRisk  = "danger";
    } else {
      filesLabel = attachments.length + " file(s) — looks safe";
      filesRisk  = "safe";
    }
  }

  // ── Urgency
  const bodyLower  = body.toLowerCase();
  const matched    = URGENCY_PHRASES.filter(p => bodyLower.includes(p));
  let urgencyLabel = "No pressure tactics", urgencyRisk = "safe";
  if (matched.length >= 3)      { urgencyLabel = "High pressure (" + matched.length + " signals)"; urgencyRisk = "danger"; }
  else if (matched.length >= 1) { urgencyLabel = "Mild pressure detected";                          urgencyRisk = "warn"; }

  // ── Links — count them and show simply
  const urls       = (body.match(/(https?:\/\/[^\s]+)/gi) || []);
  let linksLabel   = "No links found", linksRisk = "safe";
  if (urls.length > 0) {
    linksLabel = urls.length + " link(s) found";
    linksRisk  = "warn";  // backend will refine this
  }

  // Update rows immediately
  setCheckRow("sender",     "—",         "row-trust",   "");
  setCheckRow("links",      linksLabel,  "row-links",   linksRisk);
  setCheckRow("attachment", filesLabel,  "row-files",   filesRisk);
  setCheckRow("urgency",    urgencyLabel,"row-urgency",  urgencyRisk);

  currentScanData = {
    sender: senderEmail, subject: item.subject || "",
    label: "unknown",    sender_risk: isFree ? "warn" : "safe",
    auth_result: "pending", files_result: filesLabel,
    urgency_result: urgencyLabel, attachment_count: attachments.length,
    body_preview: body.substring(0, 300)
  };

  // Auth check — updates trust row when done
  checkAuth(item, senderEmail, isFree);

  // ML backend
  callBackend(body, attachments.length > 0);
}

// ── Auth check ────────────────────────────────────────────────────
function checkAuth(item, senderEmail, isFree) {
  const domain = senderEmail.split("@")[1] || "";

  if (typeof item.getAllInternetHeadersAsync === "function") {
    const timer = setTimeout(() => applyTrustFallback(isFree, domain), 4000);
    item.getAllInternetHeadersAsync((result) => {
      clearTimeout(timer);
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        parseTrust(result.value, isFree, domain);
      } else {
        applyTrustFallback(isFree, domain);
      }
    });
  } else if (item.internetHeaders?.getAsync) {
    const timer = setTimeout(() => applyTrustFallback(isFree, domain), 4000);
    item.internetHeaders.getAsync(["Authentication-Results"], (r) => {
      clearTimeout(timer);
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        parseTrust(r.value?.["Authentication-Results"] || "", isFree, domain);
      } else {
        applyTrustFallback(isFree, domain);
      }
    });
  } else {
    applyTrustFallback(isFree, domain);
  }
}

// Convert auth headers + sender info into a plain-English trust verdict
function parseTrust(rawHeaders, isFree, domain) {
  const h    = rawHeaders.toLowerCase();
  const spf  = h.includes("spf=pass");
  const dkim = h.includes("dkim=pass");
  const dmarc= h.includes("dmarc=pass");
  const n    = [spf, dkim, dmarc].filter(Boolean).length;

  let label, risk;
  if (n === 3 && !isFree)      { label = "Verified sender ✓";       risk = "safe"; }
  else if (n >= 2 && !isFree)  { label = "Mostly verified";          risk = "safe"; }
  else if (n >= 1)             { label = "Partially verified";        risk = "warn"; }
  else if (isFree)             { label = "Personal email address";    risk = "warn"; }
  else                         { label = "Could not verify sender";   risk = "warn"; }

  setCheckRow("sender", label, "row-trust", risk);
  if (currentScanData) currentScanData.auth_result = label;
}

// Fallback when auth APIs unavailable
function applyTrustFallback(isFree, domain) {
  let label, risk;
  if (!domain)       { label = "No sender info";          risk = "warn"; }
  else if (isFree)   { label = "Personal email address";  risk = "warn"; }
  else               { label = domain;                    risk = "safe"; }
  setCheckRow("sender", label, "row-trust", risk);
  if (currentScanData) currentScanData.auth_result = label;
}

// ── ML backend ────────────────────────────────────────────────────
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
    showScanTime();
    logScan(label);
  })
  .catch(() => {
    clearTimeout(timeout);
    // Local fallback
    const lower = bodyText.toLowerCase();
    const uc    = URGENCY_PHRASES.filter(p => lower.includes(p)).length;
    const hl    = /https?:\/\//i.test(bodyText);
    let label   = "ham";
    if (hl && uc >= 3) label = "phishing";
    else if (uc >= 2)  label = "spam";
    else if (hl && uc) label = "spam";
    renderGauge(label);
    showScanTime();
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

  setArcColor(label);

  const scoreLabel = document.getElementById("score-label");
  if (scoreLabel) {
    scoreLabel.textContent = c.text;
    const col = RISK_COLORS[label];
    if (col) scoreLabel.style.color = col.start;
  }

  const badge = document.getElementById("result-button");
  if (badge) { badge.textContent = c.text; badge.className = "risk-badge " + c.cls; }

  // Gauge card border glow
  const card = document.getElementById("gauge-card");
  if (card) {
    const col = RISK_COLORS[label];
    if (col) card.style.borderColor = col.start + "55";
  }

  setStatus(c.status, c.sCls);
  if (currentScanData) currentScanData.label = label;
}

function setArcColor(label) {
  const col = RISK_COLORS[label] || RISK_COLORS.ham;
  const s = document.getElementById("grad-start");
  const e = document.getElementById("grad-end");
  if (s) s.setAttribute("stop-color", col.start);
  if (e) e.setAttribute("stop-color", col.end);
}

// ── UI helpers ────────────────────────────────────────────────────

// Update a check row value + colour + row border
function setCheckRow(valueId, value, rowId, risk) {
  const el = document.getElementById(valueId);
  if (el) {
    el.textContent = value || "—";
    el.className   = "check-value " + (risk || "");
  }
  const row = document.getElementById(rowId);
  if (row) row.className = "check-row " + (risk ? risk + "-row" : "");
}

function setStatus(msg, cls) {
  const pill = document.getElementById("status");
  if (!pill) return;
  pill.textContent = msg;
  pill.className   = "status-pill " + (cls || "");
}

function setScanTime(msg) {
  const el = document.getElementById("scan-time");
  if (el) el.textContent = msg;
}

function showScanTime() {
  if (!scanStartTime) return;
  const ms = Date.now() - scanStartTime;
  setScanTime("Scanned in " + (ms / 1000).toFixed(1) + "s");
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
  btn.disabled = true; btn.textContent = "Sending...";
  fetch(BACKEND + "/report", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(Object.assign({}, currentScanData, { reported: true }))
  })
  .then(() => { btn.textContent = "✓ Reported"; })
  .catch(() => { btn.textContent = "Failed — try again"; btn.disabled = false; });
}

// ── Silent log ────────────────────────────────────────────────────
function logScan(label) {
  if (!currentScanData) return;
  fetch(BACKEND + "/log-scan", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(Object.assign({}, currentScanData, { label }))
  }).catch(() => {});
}
