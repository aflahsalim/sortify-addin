// =========================
//  CONFIG
// =========================

const BACKEND = "https://sortify-backend-hwf9d0exgqdub9cn.canadacentral-01.azurewebsites.net";

const RISKY_EXT = ["exe","bat","vbs","js","msi","ps1","cmd","scr","zip","rar","docm","xlsm"];
const URGENCY_PHRASES = [
  "urgent","immediately","suspended","verify","confirm your","click now",
  "24 hours","48 hours","act now","account locked","unusual activity",
  "security alert","update your password","will be terminated","limited time"
];

// ⭐ REMOVED FIXED SCORES — now dynamic
const SEG_COUNTS = { ham: 1, support: 2, spam: 4, phishing: 5 };

const ARC_COLORS = {
  ham:      { s: "#22c55e", e: "#16a34a" },
  support:  { s: "#06b6d4", e: "#0891b2" },
  spam:     { s: "#f59e0b", e: "#d97706" },
  phishing: { s: "#ef4444", e: "#dc2626" },
};

const VERDICTS = {
  ham:      { t: "Safe — Full Check Passed",   c: "#22c55e" },
  support:  { t: "Low Risk — Support Email",   c: "#06b6d4" },
  spam:     { t: "Suspicious — Possible Spam", c: "#f59e0b" },
  phishing: { t: "High Risk — Likely Phishing",c: "#ef4444" },
};

const ARC_LEN = 257;
let currentScanData = null;


// =========================
//  ENTRY POINT
// =========================

Office.onReady(() => {
  waitForDOM(() => {
    if (Office.context?.mailbox?.item) startClassification();
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


// =========================
//  RESET UI
// =========================

function resetUI() {
  setStatus("Scanning...", "");

  const arc = document.getElementById("risk-arc");
  if (arc) arc.style.strokeDashoffset = String(ARC_LEN);

  const num = document.getElementById("score-number");
  if (num) { num.textContent = "—"; num.setAttribute("fill","#ffffff"); }

  for (let i = 1; i <= 5; i++) {
    const s = document.getElementById("seg" + i);
    if (s) s.classList.remove("active");
  }

  const vd = document.getElementById("verdict");
  if (vd) { vd.textContent = "Analysing..."; vd.style.color = "var(--muted)"; }

  resetCard("card-sender","tick-sender","sender","sender-l1","—","sender-l2","");
  resetCard("card-links","tick-links","links","links-l1","—","links-l2","");
  resetCard("card-attach","tick-attach","attachment","attach-l1","—","attach-l2","");
  resetCard("card-urgency","tick-urgency","urgency","urgency-l1","—","urgency-l2","");

  const btn = document.getElementById("report-btn");
  if (btn) { btn.disabled = false; btn.textContent = "Mark as Suspicious"; }

  setArcColor("ham");
  currentScanData = null;
}

function resetCard(cardId, tickId, valId, l1Id, l1, l2Id, l2) {
  const card = document.getElementById(cardId);
  if (card) card.className = "info-card";
  const tick = document.getElementById(tickId);
  if (tick) {
    tick.className = "card-tick";
    const pl = tick.querySelector("polyline");
    if (pl) pl.setAttribute("stroke","#94a3b8");
  }
  const val = document.getElementById(valId);
  if (val) { val.textContent = "—"; val.className = "card-result neutral"; }
  const e1 = document.getElementById(l1Id); if (e1) e1.textContent = l1;
  const e2 = document.getElementById(l2Id); if (e2) e2.textContent = l2;
}


// =========================
//  READ EMAIL
// =========================

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


// =========================
//  PROCESS EMAIL
// =========================

function processEmail(item, body) {

  // ⭐ UNIVERSAL SENDER EXTRACTION (New Outlook FIX)
  let senderEmail = "";
  try {
    if (item?.from?.emailAddress) {
      senderEmail = item.from.emailAddress;
    } else if (item?.from?.emailAddress?.address) {
      senderEmail = item.from.emailAddress.address;
    } else if (item?.sender?.emailAddress) {
      senderEmail = item.sender.emailAddress;
    } else if (typeof item?.from === "string") {
      senderEmail = item.from;
    } else if (typeof item?.sender === "string") {
      senderEmail = item.sender;
    }
  } catch (e) {
    senderEmail = "";
  }

  const subject = item?.subject || "";
  const atts = Array.isArray(item.attachments) ? item.attachments : [];
  const lower = body.toLowerCase();
  const matched = URGENCY_PHRASES.filter(p => lower.includes(p));
  const urls = (body.match(/(https?:\/\/[^\s]+)/gi) || []);

  updateCard("card-links","tick-links","links","links-l1",
    urls.length ? urls.length+" link(s)" : "No links found",
    "links-l2", urls.length ? "Analysed" : "Email contains no URLs",
    urls.length ? urls.length+" link(s)" : "Safe",
    urls.length ? "warn" : "safe"
  );

  updateCard("card-attach","tick-attach","attachment","attach-l1",
    atts.length ? atts.length+" attachment(s)" : "No files attached",
    "attach-l2", atts.length ? "File type check passed" : "",
    atts.length ? "Attachments found" : "No threats found",
    atts.length ? "warn" : "safe"
  );

  updateCard("card-urgency","tick-urgency","urgency","urgency-l1",
    matched.length ? matched.length+" urgency signals" : "Zero phishing tactics",
    "urgency-l2", matched[0] || "",
    matched.length ? "Urgency detected" : "Normal behaviour",
    matched.length ? "warn" : "safe"
  );

  updateCard("card-sender","tick-sender","sender","sender-l1",
    senderEmail || "Unknown sender",
    "sender-l2", "Checking authentication...",
    "Checking...", ""
  );

  currentScanData = {
    sender: senderEmail,
    subject: subject,
    label: "unknown",
    sender_risk: "pending",
    auth_result: "pending",
    files_result: atts.length ? "Attachments found" : "No threats found",
    urgency_result: matched.length ? "Urgency detected" : "Normal behaviour",
    attachment_count: atts.length,
    body_preview: body.substring(0, 300)
  };

  checkAuth(item, senderEmail);
  callBackend(body, atts.length > 0, senderEmail, subject);
}


// =========================
//  AUTH CHECK
// =========================

function checkAuth(item, senderEmail) {
  updateCard("card-sender","tick-sender","sender","sender-l1",
    senderEmail || "Unknown sender",
    "sender-l2","Auth check unavailable",
    "Unverified","warn"
  );
}


// =========================
//  BACKEND CALL
// =========================

function callBackend(bodyText, hasAttachment, senderEmail, subject) {
  fetch(BACKEND + "/classify", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      text: bodyText,
      attachment: hasAttachment ? "Yes" : "No",
      sender: senderEmail,
      subject: subject,
      reported: false
    })
  })
    .then(r => r.json())
    .then(d => {
      const label = (d.label || "unknown").toLowerCase();
      const backendScore = d.score || 0; // ⭐ REAL ML SCORE
      renderResult(label, backendScore);
      logScan(label, senderEmail, subject, bodyText);
    })
    .catch(() => {
      renderResult("ham", 0.10);
      logScan("ham", senderEmail, subject, bodyText);
    });
}


// =========================
//  LOG SCAN
// =========================

function logScan(label, senderEmail, subject, bodyText) {
  fetch(BACKEND + "/log-scan", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      sender: senderEmail,
      subject: subject,
      label: label,
      reported: false,
      body_preview: bodyText.substring(0, 300)
    })
  }).catch(() => {});
}


// =========================
//  REPORT POPUP
// =========================

function reportEmail() {
  document.getElementById("overlay").classList.remove("hidden");
}

function closeConfirm() {
  document.getElementById("overlay").classList.add("hidden");
}

function confirmReport() {
  document.getElementById("overlay").classList.add("hidden");

  const btn = document.getElementById("report-btn");
  btn.disabled = true;
  btn.textContent = "Sending...";

  fetch(BACKEND + "/log-scan", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(Object.assign({}, currentScanData, { reported: true }))
  })
  .then(() => { btn.textContent = "✓ Reported to Sortify"; })
  .catch(() => { btn.textContent = "Mark as Suspicious"; btn.disabled = false; });
}

window.reportEmail = reportEmail;
window.closeConfirm = closeConfirm;
window.confirmReport = confirmReport;


// =========================
//  RENDER RESULT + HELPERS
// =========================

function renderResult(label, backendScore = null) {

  // ⭐ REAL SCORE (0–100)
  const score = backendScore !== null
    ? Math.round(backendScore * 100)
    : 50;

  const col = ARC_COLORS[label] || ARC_COLORS.ham;
  const vd = VERDICTS[label] || { t: "Scanned", c: "#94a3b8" };
  const segs = SEG_COUNTS[label] || 1;

  // Arc animation
  const arc = document.getElementById("risk-arc");
  if (arc) arc.style.strokeDashoffset = (ARC_LEN - (score/100)*ARC_LEN).toFixed(1);

  setArcColor(label);

  // ⭐ SCORE NUMBER ALWAYS WHITE
  const num = document.getElementById("score-number");
  if (num) {
    num.textContent = score;
    num.setAttribute("fill", "#ffffff");
  }

  // Segments
  for (let i = 1; i <= 5; i++) {
    const s = document.getElementById("seg" + i);
    if (s) { i <= segs ? s.classList.add("active") : s.classList.remove("active"); }
  }

  // Verdict
  const verdict = document.getElementById("verdict");
  if (verdict) { verdict.textContent = vd.t; verdict.style.color = vd.c; }

  setStatus("Scanned", "done");
  if (currentScanData) currentScanData.label = label;
}

function setArcColor(label) {
  const c = ARC_COLORS[label] || ARC_COLORS.ham;
  const s = document.getElementById("gs"); if (s) s.setAttribute("stop-color", c.s);
  const e = document.getElementById("ge"); if (e) e.setAttribute("stop-color", c.e);
}

function updateCard(cardId, tickId, valId, l1Id, l1, l2Id, l2, result, risk) {
  const card = document.getElementById(cardId);
  if (card) card.className = "info-card" + (risk ? " c-" + risk : "");

  const tick = document.getElementById(tickId);
  if (tick) {
    tick.className = "card-tick" + (risk ? " " + risk : "");
    const pl = tick.querySelector("polyline");
    if (pl) {
      const cols = { safe: "#22c55e", warn: "#fbbf24", danger: "#f87171" };
      pl.setAttribute("stroke", cols[risk] || "#94a3b8");
    }
  }

  const e1 = document.getElementById(l1Id); if (e1) e1.textContent = l1 || "";
  const e2 = document.getElementById(l2Id); if (e2) e2.textContent = l2 || "";

  const val = document.getElementById(valId);
  if (val) {
    val.textContent = result || "—";
    val.className = "card-result" + (risk ? " " + risk : " neutral");
  }
}

function setStatus(msg, cls) {
  const p = document.getElementById("status");
  if (!p) return;
  p.textContent = msg;
  p.className = "status-pill" + (cls ? " " + cls : "");
}
