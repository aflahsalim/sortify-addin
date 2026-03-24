// ── Config ────────────────────────────────────────────────────────
const BACKEND = "https://sortify-backend-hwf9d0exgqdub9cn.canadacentral-01.azurewebsites.net";
const RISKY_EXT = ["exe","bat","vbs","js","msi","ps1","cmd","scr","zip","rar","docm","xlsm"];
const URGENCY_PHRASES = [
  "urgent","immediately","suspended","verify","confirm your",
  "click now","24 hours","48 hours","act now","account locked",
  "unusual activity","security alert","update your password",
  "will be terminated","will be suspended","limited time"
];

// Risk score (0–100) per label — shown as the big number
const RISK_SCORES = { ham: 5, support: 20, spam: 72, phishing: 95 };

// Arc gradient colours per risk
const RISK_COLORS = {
  ham:      { start: "#22c55e", end: "#16a34a" },
  support:  { start: "#06b6d4", end: "#0891b2" },
  spam:     { start: "#f59e0b", end: "#d97706" },
  phishing: { start: "#ef4444", end: "#dc2626" },
};

// Colour bar dot position (%) per label
const BAR_POSITIONS = { ham: 4, support: 18, spam: 70, phishing: 95 };

// Verdict text shown under the bar
const VERDICTS = {
  ham:      { text: "Safe — Full Check Passed",  color: "#22c55e" },
  support:  { text: "Low Risk — Support Email",  color: "#06b6d4" },
  spam:     { text: "Suspicious — Possible Spam",color: "#f59e0b" },
  phishing: { text: "High Risk — Likely Phishing",color: "#ef4444" },
};

let currentScanData = null;

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

  // Reset gauge
  const arc = document.getElementById("risk-arc");
  if (arc) arc.style.strokeDashoffset = "251";
  const needle = document.getElementById("needle");
  if (needle) needle.style.transform = "rotate(-90deg)";
  const num = document.getElementById("score-number");
  if (num) { num.textContent = "—"; num.setAttribute("fill","#22c55e"); }
  setBarDot(0);
  const verdict = document.getElementById("verdict");
  if (verdict) { verdict.textContent = "Analysing..."; verdict.style.color = "var(--muted)"; }

  // Reset cards
  ["sender","links","attachment","urgency"].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.textContent = "—"; el.className = "card-result neutral"; }
  });
  ["sender-dot","links-dot","attachment-dot","urgency-dot"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.className = "card-dot";
  });
  ["card-sender","card-links","card-attach","card-urgency"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.className = "info-card";
  });
  setCardBody("sender-body",  "Checking sender...");
  setCardBody("links-body",   "Checking links...");
  setCardBody("attach-body",  "Checking files...");
  setCardBody("urgency-body", "Checking language...");

  // Reset score card border
  const sc = document.getElementById("score-card");
  if (sc) sc.style.borderColor = "";

  // Reset button
  const btn = document.getElementById("report-btn");
  if (btn) { btn.disabled = false; btn.innerHTML = "⚑ &nbsp;Mark as Suspicious"; }

  setArcColor("ham");
  currentScanData = null;
}

// ── Step 1: read body ─────────────────────────────────────────────
function startClassification() {
  const item = Office.context?.mailbox?.item;
  if (!item) return setStatus("Open an email to scan", "");
  setStatus("Scanning...", "");
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded)
      return setStatus("Could not read email", "warn");
    processEmail(item, result.value || "");
  });
}

// ── Step 2: all checks ────────────────────────────────────────────
function processEmail(item, body) {

  // Sender
  let senderEmail = "";
  try { senderEmail = item?.from?.emailAddress?.address || ""; } catch(e) {}
  const domain    = senderEmail.split("@")[1] || "";
  const freeDomains = ["gmail.com","yahoo.com","hotmail.com","outlook.com","live.com"];
  const isFree    = freeDomains.includes(domain.toLowerCase());

  // Attachments
  const attachments = Array.isArray(item.attachments) ? item.attachments : [];
  let filesResult, filesBody, filesRisk = "safe";
  if (attachments.length === 0) {
    filesResult = "No attachments";
    filesBody   = "No files attached";
  } else {
    const risky = attachments.find(a =>
      RISKY_EXT.includes((a.name || "").split(".").pop().toLowerCase())
    );
    if (risky) {
      filesResult = "Risky file type";
      filesBody   = risky.name + " — executable/macro file";
      filesRisk   = "danger";
    } else {
      filesResult = "No threats found";
      filesBody   = attachments.length + " attachment(s) — safe type";
    }
  }

  // Urgency / behaviour
  const bodyLower = body.toLowerCase();
  const matched   = URGENCY_PHRASES.filter(p => bodyLower.includes(p));
  let urgencyResult, urgencyBody, urgencyRisk = "safe";
  if (matched.length >= 3) {
    urgencyResult = "High pressure tactics";
    urgencyBody   = matched.length + " phishing signals detected";
    urgencyRisk   = "danger";
  } else if (matched.length >= 1) {
    urgencyResult = "Mild pressure language";
    urgencyBody   = '"' + matched[0] + '" and similar phrases';
    urgencyRisk   = "warn";
  } else {
    urgencyResult = "Zero phishing tactics";
    urgencyBody   = "No urgent or threatening language";
  }

  // Links
  const urls = (body.match(/(https?:\/\/[^\s]+)/gi) || []);
  let linksResult, linksBody, linksRisk = "safe";
  if (urls.length === 0) {
    linksResult = "No links found";
    linksBody   = "No external URLs in email";
  } else {
    linksResult = urls.length + " link(s) found";
    linksBody   = "Analysed against risk model";
    linksRisk   = "warn";
  }

  // Update cards immediately
  setCard("sender",     "—",           "sender-dot",     "",         "sender-body", "Checking sender...",    "card-sender",  "");
  setCard("links",      linksResult,   "links-dot",      linksRisk,  "links-body",  linksBody,               "card-links",   linksRisk);
  setCard("attachment", filesResult,   "attachment-dot", filesRisk,  "attach-body", filesBody,               "card-attach",  filesRisk);
  setCard("urgency",    urgencyResult, "urgency-dot",    urgencyRisk,"urgency-body",urgencyBody,              "card-urgency", urgencyRisk);

  currentScanData = {
    sender: senderEmail, subject: item.subject || "",
    label: "unknown",    sender_risk: isFree ? "warn" : "safe",
    auth_result: "pending", files_result: filesResult,
    urgency_result: urgencyResult, attachment_count: attachments.length,
    body_preview: body.substring(0, 300)
  };

  // Auth check updates sender card
  checkAuth(item, senderEmail, isFree, domain);

  // ML backend
  callBackend(body, attachments.length > 0);
}

// ── Auth check ────────────────────────────────────────────────────
function checkAuth(item, senderEmail, isFree, domain) {
  if (typeof item.getAllInternetHeadersAsync === "function") {
    const timer = setTimeout(() => applyTrustFallback(isFree, domain, senderEmail), 4000);
    item.getAllInternetHeadersAsync((result) => {
      clearTimeout(timer);
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        parseTrust(result.value, isFree, domain, senderEmail);
      } else {
        applyTrustFallback(isFree, domain, senderEmail);
      }
    });
  } else if (item.internetHeaders?.getAsync) {
    const timer = setTimeout(() => applyTrustFallback(isFree, domain, senderEmail), 4000);
    item.internetHeaders.getAsync(["Authentication-Results"], (r) => {
      clearTimeout(timer);
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        parseTrust(r.value?.["Authentication-Results"] || "", isFree, domain, senderEmail);
      } else {
        applyTrustFallback(isFree, domain, senderEmail);
      }
    });
  } else {
    applyTrustFallback(isFree, domain, senderEmail);
  }
}

function parseTrust(rawHeaders, isFree, domain, senderEmail) {
  const h    = rawHeaders.toLowerCase();
  const spf  = h.includes("spf=pass");
  const dkim = h.includes("dkim=pass");
  const dmarc= h.includes("dmarc=pass");
  const n    = [spf, dkim, dmarc].filter(Boolean).length;

  let result, body, risk;
  if (n === 3 && !isFree)     { result = "Verified ✓";         body = "Verified by DMARC/SPF/DKIM";   risk = "safe"; }
  else if (n >= 2 && !isFree) { result = "Mostly verified";     body = n + "/3 checks passed";          risk = "safe"; }
  else if (n >= 1)            { result = "Partially verified";  body = n + "/3 auth checks passed";     risk = "warn"; }
  else if (isFree)            { result = "Personal address";    body = "Free email — " + (domain || senderEmail); risk = "warn"; }
  else                        { result = "Unverified";          body = "Could not verify sender";        risk = "warn"; }

  setCard("sender", result, "sender-dot", risk, "sender-body", body, "card-sender", risk);
  if (currentScanData) currentScanData.auth_result = result;
}

function applyTrustFallback(isFree, domain, senderEmail) {
  let result, body, risk;
  if (!domain)     { result = "No sender info";     body = "Sender address missing";               risk = "warn"; }
  else if (isFree) { result = "Personal address";   body = "Free email — " + (domain || senderEmail); risk = "warn"; }
  else             { result = "Domain: " + domain;  body = "Auth headers not accessible";           risk = "safe"; }
  setCard("sender", result, "sender-dot", risk, "sender-body", body, "card-sender", risk);
  if (currentScanData) currentScanData.auth_result = result;
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
    renderResult((data.label || "unknown").toLowerCase());
    logScan((data.label || "unknown").toLowerCase());
  })
  .catch(() => {
    clearTimeout(timeout);
    const lower = bodyText.toLowerCase();
    const uc    = URGENCY_PHRASES.filter(p => lower.includes(p)).length;
    const hl    = /https?:\/\//i.test(bodyText);
    let label   = "ham";
    if (hl && uc >= 3) label = "phishing";
    else if (uc >= 2)  label = "spam";
    else if (hl && uc) label = "spam";
    renderResult(label);
    logScan(label);
  });
}

// ── Render final result ───────────────────────────────────────────
function renderResult(label) {
  const map = {
    ham:      { angle: -90, fill: 0.0,  sCls: "done",   status: "All clear"   },
    support:  { angle: -45, fill: 0.25, sCls: "done",   status: "Low risk"    },
    spam:     { angle:  45, fill: 0.75, sCls: "warn",   status: "Suspicious"  },
    phishing: { angle:  90, fill: 1.0,  sCls: "danger", status: "High risk"   },
  };
  const c = map[label] || { angle: 0, fill: 0.5, sCls: "done", status: "Scanned" };

  // Needle
  const needle = document.getElementById("needle");
  if (needle) needle.style.transform = "rotate(" + c.angle + "deg)";

  // Arc
  const arc = document.getElementById("risk-arc");
  if (arc) arc.style.strokeDashoffset = 251 - c.fill * 251;

  // Arc colour
  setArcColor(label);

  // Score number
  const score = RISK_SCORES[label] || 50;
  const col   = RISK_COLORS[label] || RISK_COLORS.ham;
  const num   = document.getElementById("score-number");
  if (num) { num.textContent = score; num.setAttribute("fill", col.start); }

  // Colour bar dot
  setBarDot(BAR_POSITIONS[label] || 50);

  // Verdict
  const vd = VERDICTS[label];
  const verdict = document.getElementById("verdict");
  if (verdict && vd) { verdict.textContent = vd.text; verdict.style.color = vd.color; }

  // Score card border glow
  const sc = document.getElementById("score-card");
  if (sc) sc.style.borderColor = col.start + "55";

  setStatus(c.status, c.sCls);
  if (currentScanData) currentScanData.label = label;
}

// ── Helpers ───────────────────────────────────────────────────────

function setArcColor(label) {
  const col = RISK_COLORS[label] || RISK_COLORS.ham;
  const s = document.getElementById("grad-start");
  const e = document.getElementById("grad-end");
  if (s) s.setAttribute("stop-color", col.start);
  if (e) e.setAttribute("stop-color", col.end);
}

function setBarDot(pct) {
  const dot = document.getElementById("bar-dot");
  if (dot) dot.style.left = Math.min(Math.max(pct, 1), 97) + "%";
}

// Update a full info card at once
function setCard(valueId, value, dotId, risk, bodyId, bodyText, cardId, cardRisk) {
  const el = document.getElementById(valueId);
  if (el) { el.textContent = value || "—"; el.className = "card-result " + (risk || "neutral"); }
  const dot = document.getElementById(dotId);
  if (dot) dot.className = "card-dot " + (risk || "");
  const bd = document.getElementById(bodyId);
  if (bd) bd.textContent = bodyText || "";
  const card = document.getElementById(cardId);
  if (card) card.className = "info-card " + (cardRisk ? cardRisk + "-card" : "");
}

function setCardBody(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
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
  btn.disabled = true;
  btn.innerHTML = "Sending...";
  fetch(BACKEND + "/report", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(Object.assign({}, currentScanData, { reported: true }))
  })
  .then(() => { btn.innerHTML = "✓ Reported to Sortify"; })
  .catch(() => { btn.innerHTML = "⚑ &nbsp;Mark as Suspicious"; btn.disabled = false; });
}

// ── Silent log ────────────────────────────────────────────────────
function logScan(label) {
  if (!currentScanData) return;
  fetch(BACKEND + "/log-scan", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(Object.assign({}, currentScanData, { label }))
  }).catch(() => {});
}
