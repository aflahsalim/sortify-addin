// ── Config ────────────────────────────────────────────────────────
const BACKEND = "https://sortify-backend-hwf9d0exgqdub9cn.canadacentral-01.azurewebsites.net";
const RISKY_EXT = ["exe","bat","vbs","js","msi","ps1","cmd","scr","zip","rar","docm","xlsm"];
const URGENCY_PHRASES = [
  "urgent","immediately","suspended","verify","confirm your",
  "click now","24 hours","48 hours","act now","account locked",
  "unusual activity","security alert","update your password"
];

// Gauge arc colours per risk level
const RISK_COLORS = {
  ham:      { start: "#00ff88", end: "#00cc66" },
  support:  { start: "#22d3ee", end: "#0891b2" },
  spam:     { start: "#fbbf24", end: "#d97706" },
  phishing: { start: "#f87171", end: "#dc2626" },
};

let currentScanData = null;

// ── Entry point ───────────────────────────────────────────────────
Office.onReady(() => {
  waitForDOM(() => {
    if (Office.context?.mailbox?.item) {
      startClassification();
    } else {
      setStatus("Click an email to scan", "");
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

// ── Reset between emails ──────────────────────────────────────────
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
  if (arc) arc.style.strokeDashoffset = "226";
  const needle = document.getElementById("needle");
  if (needle) needle.style.transform = "rotate(-90deg)";
  const lbl = document.getElementById("score-label");
  if (lbl) { lbl.textContent = "—"; lbl.style.color = ""; }
  const badge = document.getElementById("result-button");
  if (badge) { badge.textContent = "—"; badge.className = "risk-badge"; }
  const btn = document.getElementById("report-btn");
  if (btn) { btn.disabled = false; btn.textContent = "Send to Sortify team"; }
  setArcColor("ham");
  currentScanData = null;
}

// ── Step 1: read body ─────────────────────────────────────────────
function startClassification() {
  const item = Office.context?.mailbox?.item;
  if (!item) return setStatus("No email selected", "");
  setStatus("Scanning...", "");
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded)
      return setStatus("Could not read email", "warn");
    processEmail(item, result.value || "");
  });
}

// ── Step 2: all checks ────────────────────────────────────────────
function processEmail(item, body) {

  // 1. Sender
  let senderEmail = "";
  try { senderEmail = item?.from?.emailAddress?.address || ""; } catch(e) {}
  const domain    = senderEmail.split("@")[1] || "";
  const freeDomains = ["gmail.com","yahoo.com","hotmail.com","outlook.com"];
  const isFree    = freeDomains.includes(domain.toLowerCase());
  const senderLabel = isFree ? "Free domain" : (domain || "Unknown");
  const senderRisk  = isFree ? "warn" : "safe";

  // 2. Attachments (sync)
  const attachments = Array.isArray(item.attachments) ? item.attachments : [];
  let filesLabel = "None", filesRisk = "safe";
  if (attachments.length > 0) {
    const risky = attachments.find(a =>
      RISKY_EXT.includes((a.name || "").split(".").pop().toLowerCase())
    );
    if (risky) {
      filesLabel = risky.name.split(".").pop().toUpperCase() + " — risky";
      filesRisk  = "danger";
    } else {
      filesLabel = attachments.length + " safe file(s)";
    }
  }

  // 3. Urgency (sync)
  const bodyLower = body.toLowerCase();
  const matched   = URGENCY_PHRASES.filter(p => bodyLower.includes(p));
  let urgencyLabel = "None detected", urgencyRisk = "safe";
  if (matched.length >= 3)      { urgencyLabel = matched.length + " signals — High"; urgencyRisk = "danger"; }
  else if (matched.length >= 1) { urgencyLabel = '"' + matched[0] + '"';             urgencyRisk = "warn"; }

  // Update rows immediately with what we have
  setField("sender",     senderLabel,  "sender-dot",     senderRisk);
  setField("links",      "Checking…",  "links-dot",      "warn");
  setField("attachment", filesLabel,   "attachment-dot", filesRisk);
  setField("urgency",    urgencyLabel, "urgency-dot",    urgencyRisk);

  currentScanData = {
    sender: senderEmail, subject: item.subject || "",
    label: "unknown",    sender_risk: senderRisk,
    auth_result: "Checking", files_result: filesLabel,
    urgency_result: urgencyLabel, attachment_count: attachments.length,
    body_preview: body.substring(0, 300)
  };

  // 4. Auth — try getAllInternetHeadersAsync first (better desktop support)
  //    fall back to internetHeaders if not available
  //    fall back to heuristic if neither works
  checkAuth(item, senderEmail);

  // Call ML backend in parallel
  callBackend(body, attachments.length > 0);
}

// ── Auth check — tries 3 methods ─────────────────────────────────
function checkAuth(item, senderEmail) {

  // Method 1: getAllInternetHeadersAsync (best support in Outlook Desktop)
  if (typeof item.getAllInternetHeadersAsync === "function") {
    const timer = setTimeout(() => {
      // If it takes more than 4s, fall back to heuristic
      applyHeuristicAuth(senderEmail);
    }, 4000);

    item.getAllInternetHeadersAsync((result) => {
      clearTimeout(timer);
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        parseAuthHeaders(result.value);
      } else {
        applyHeuristicAuth(senderEmail);
      }
    });

  // Method 2: internetHeaders.getAsync (Outlook Web)
  } else if (item.internetHeaders && typeof item.internetHeaders.getAsync === "function") {
    const timer = setTimeout(() => applyHeuristicAuth(senderEmail), 4000);
    item.internetHeaders.getAsync(["Authentication-Results"], (r) => {
      clearTimeout(timer);
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        const h = r.value?.["Authentication-Results"] || "";
        parseAuthHeaders(h);
      } else {
        applyHeuristicAuth(senderEmail);
      }
    });

  // Method 3: neither available — use heuristic
  } else {
    applyHeuristicAuth(senderEmail);
  }
}

// Parse the raw Authentication-Results header string
function parseAuthHeaders(rawHeaders) {
  const h     = rawHeaders.toLowerCase();
  const spf   = h.includes("spf=pass");
  const dkim  = h.includes("dkim=pass");
  const dmarc = h.includes("dmarc=pass");
  const n     = [spf, dkim, dmarc].filter(Boolean).length;

  let authLabel, authRisk;
  if (n === 3)      { authLabel = "SPF · DKIM · DMARC ✓"; authRisk = "safe"; }
  else if (n === 2) { authLabel = "2/3 checks pass";       authRisk = "safe"; }
  else if (n === 1) { authLabel = "Partial — 1/3 pass";    authRisk = "warn"; }
  else if (h.includes("spf=") || h.includes("dkim=")) {
                      authLabel = "Auth failed";            authRisk = "danger"; }
  else              { authLabel = "No auth data";           authRisk = "warn"; }

  setField("links", authLabel, "links-dot", authRisk);
  if (currentScanData) currentScanData.auth_result = authLabel;
}

// Heuristic fallback — guess based on sender domain
// If no header APIs work, we at least show something meaningful
function applyHeuristicAuth(senderEmail) {
  const domain = senderEmail.split("@")[1] || "";
  const freeDomains = ["gmail.com","yahoo.com","hotmail.com","outlook.com"];

  let authLabel, authRisk;
  if (!domain) {
    authLabel = "No sender info";
    authRisk  = "warn";
  } else if (freeDomains.includes(domain.toLowerCase())) {
    // Free domains usually have SPF — flag as partial
    authLabel = "Free domain";
    authRisk  = "warn";
  } else {
    // Corporate domain — assume passing unless we know otherwise
    authLabel = "Headers inaccessible";
    authRisk  = "warn";
  }

  setField("links", authLabel, "links-dot", authRisk);
  if (currentScanData) currentScanData.auth_result = authLabel;
}

// ── Step 3: ML backend ────────────────────────────────────────────
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
    clearTimeout(timeout);
    setStatus("Local analysis", "warn");
    const lower = bodyText.toLowerCase();
    const uc    = URGENCY_PHRASES.filter(p => lower.includes(p)).length;
    const hl    = /https?:\/\//i.test(bodyText);
    let label   = "ham";
    if (hl && uc >= 3) label = "phishing";
    else if (uc >= 2)  label = "spam";
    else if (hl && uc) label = "spam";
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
  if (arc) arc.style.strokeDashoffset = 226 - c.fill * 226;

  setArcColor(label);

  const scoreLabel = document.getElementById("score-label");
  if (scoreLabel) {
    scoreLabel.textContent = c.text;
    const col = RISK_COLORS[label];
    if (col) scoreLabel.style.color = col.start;
  }

  const badge = document.getElementById("result-button");
  if (badge) { badge.textContent = c.text; badge.className = "risk-badge " + c.cls; }

  const card = document.getElementById("gauge-card");
  if (card) {
    const col = RISK_COLORS[label];
    if (col) card.style.borderColor = col.start + "44";
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
function setField(valueId, value, dotId, risk) {
  const el = document.getElementById(valueId);
  if (el) el.textContent = value || "—";
  const dot = document.getElementById(dotId);
  if (dot) dot.className = "dot " + (risk || "");
}

function setStatus(msg, cls) {
  const pill = document.getElementById("status");
  if (!pill) return;
  pill.textContent = msg;
  pill.className   = "status-pill " + (cls || "");
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
