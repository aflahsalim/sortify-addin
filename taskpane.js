const BACKEND = "https://sortify-backend-hwf9d0exgqdub9cn.canadacentral-01.azurewebsites.net";
const RISKY_EXT = ["exe","bat","vbs","js","msi","ps1","cmd","scr","zip","rar","docm","xlsm"];
const URGENCY_PHRASES = [
  "urgent","immediately","suspended","verify","confirm your","click now",
  "24 hours","48 hours","act now","account locked","unusual activity",
  "security alert","update your password","will be terminated","limited time"
];

// Risk score 0-100 shown as big number AND drives arc fill
const RISK_SCORES = { ham: 5, support: 22, spam: 71, phishing: 94 };

// How many of 5 segments light up
const SEG_COUNTS = { ham: 1, support: 2, spam: 4, phishing: 5 };

// Arc gradient colours
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

// Total SVG arc path length for "M18 100 A82 82 0 0 1 182 100"
// Arc = π * r = π * 82 ≈ 257.6 — use 257
const ARC_LEN = 257;

let currentScanData = null;

// ── Entry point ───────────────────────────────────────────────────
Office.onReady(() => {
  waitForDOM(() => {
    if (Office.context?.mailbox?.item) startClassification();
    else setStatus("Open an email to scan", "");
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

  // Arc starts empty (full offset = hidden)
  const arc = document.getElementById("risk-arc");
  if (arc) arc.style.strokeDashoffset = String(ARC_LEN);

  // Score number
  const num = document.getElementById("score-number");
  if (num) { num.textContent = "—"; num.setAttribute("fill","#22c55e"); }

  // Colour segments all dim
  for (let i = 1; i <= 5; i++) {
    const s = document.getElementById("seg" + i);
    if (s) s.classList.remove("active");
  }

  // Verdict
  const vd = document.getElementById("verdict");
  if (vd) { vd.textContent = "Analysing..."; vd.style.color = "var(--muted)"; }

  // Cards
  resetCard("card-sender",  "tick-sender",  "sender",     "sender-l1",  "—", "sender-l2",  "");
  resetCard("card-links",   "tick-links",   "links",      "links-l1",   "—", "links-l2",   "");
  resetCard("card-attach",  "tick-attach",  "attachment", "attach-l1",  "—", "attach-l2",  "");
  resetCard("card-urgency", "tick-urgency", "urgency",    "urgency-l1", "—", "urgency-l2", "");

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

// ── Read email ────────────────────────────────────────────────────
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

// ── All checks ────────────────────────────────────────────────────
function processEmail(item, body) {

  // Sender
  let senderEmail = "";
  try { senderEmail = item?.from?.emailAddress?.address || ""; } catch(e) {}
  const domain    = senderEmail.split("@")[1] || "";
  const freeDomains = ["gmail.com","yahoo.com","hotmail.com","outlook.com","live.com"];
  const isFree    = freeDomains.includes(domain.toLowerCase());
  let senderName  = "";
  try { senderName = item?.from?.emailAddress?.displayName || ""; } catch(e) {}

  // Attachments
  const atts = Array.isArray(item.attachments) ? item.attachments : [];
  let aL1, aL2, aResult, aRisk = "safe";
  if (atts.length === 0) {
    aL1 = "No files attached"; aL2 = ""; aResult = "No threats found";
  } else {
    const risky = atts.find(a => RISKY_EXT.includes((a.name||"").split(".").pop().toLowerCase()));
    if (risky) {
      aL1 = "1 attachment: " + risky.name; aL2 = "Executable/macro file type";
      aResult = "Risky file detected"; aRisk = "danger";
    } else {
      aL1 = atts.length + " attachment(s)"; aL2 = "File type check passed";
      aResult = "No threats found";
    }
  }

  // Urgency
  const lower   = body.toLowerCase();
  const matched = URGENCY_PHRASES.filter(p => lower.includes(p));
  let uL1, uL2, uResult, uRisk = "safe";
  if (matched.length >= 3) {
    uL1 = matched.length + " phishing signals found"; uL2 = "Urgency & threat language";
    uResult = "High pressure tactics"; uRisk = "danger";
  } else if (matched.length >= 1) {
    uL1 = "Mild pressure language"; uL2 = '"' + matched[0] + '"';
    uResult = "Low pressure detected"; uRisk = "warn";
  } else {
    uL1 = "Zero phishing tactics"; uL2 = "No urgent language detected";
    uResult = "Behaviour looks normal";
  }

  // Links
  const urls = (body.match(/(https?:\/\/[^\s]+)/gi) || []);
  let lL1, lL2, lResult, lRisk = "safe";
  if (urls.length === 0) {
    lL1 = "No links found"; lL2 = "Email contains no URLs"; lResult = "Safe";
  } else {
    lL1 = urls.length + " link(s) found"; lL2 = "Analysed against risk model";
    lResult = urls.length + " link(s)"; lRisk = "warn";
  }

  // Update cards
  updateCard("card-links",   "tick-links",   "links",      "links-l1",   lL1, "links-l2",   lL2, lResult,  lRisk);
  updateCard("card-attach",  "tick-attach",  "attachment", "attach-l1",  aL1, "attach-l2",  aL2, aResult,  aRisk);
  updateCard("card-urgency", "tick-urgency", "urgency",    "urgency-l1", uL1, "urgency-l2", uL2, uResult,  uRisk);
  updateCard("card-sender",  "tick-sender",  "sender",     "sender-l1",  senderName||senderEmail, "sender-l2", "Verifying...", "Checking...", "");

  currentScanData = {
    sender: senderEmail, subject: item.subject || "", label: "unknown",
    sender_risk: isFree ? "warn" : "safe", auth_result: "pending",
    files_result: aResult, urgency_result: uResult,
    attachment_count: atts.length, body_preview: body.substring(0, 300)
  };

  checkAuth(item, senderEmail, isFree, domain, senderName);
  callBackend(body, atts.length > 0);
}

// ── Auth ──────────────────────────────────────────────────────────
function checkAuth(item, senderEmail, isFree, domain, senderName) {
  const display = senderName || senderEmail;
  if (typeof item.getAllInternetHeadersAsync === "function") {
    const t = setTimeout(() => fallbackTrust(isFree, domain, display), 4000);
    item.getAllInternetHeadersAsync((r) => {
      clearTimeout(t);
      r.status === Office.AsyncResultStatus.Succeeded && r.value
        ? parseTrust(r.value, isFree, domain, display)
        : fallbackTrust(isFree, domain, display);
    });
  } else if (item.internetHeaders?.getAsync) {
    const t = setTimeout(() => fallbackTrust(isFree, domain, display), 4000);
    item.internetHeaders.getAsync(["Authentication-Results"], (r) => {
      clearTimeout(t);
      r.status === Office.AsyncResultStatus.Succeeded
        ? parseTrust(r.value?.["Authentication-Results"]||"", isFree, domain, display)
        : fallbackTrust(isFree, domain, display);
    });
  } else {
    fallbackTrust(isFree, domain, display);
  }
}

function parseTrust(raw, isFree, domain, display) {
  const h = raw.toLowerCase();
  const n = [h.includes("spf=pass"),h.includes("dkim=pass"),h.includes("dmarc=pass")].filter(Boolean).length;
  let l2, result, risk;
  if (n===3&&!isFree)      { l2="Verified by DMARC/SPF/DKIM"; result="Sender Verified ✓"; risk="safe"; }
  else if (n>=2&&!isFree)  { l2=n+"/3 auth checks passed";    result="Mostly verified";   risk="safe"; }
  else if (n>=1)           { l2=n+"/3 auth checks passed";    result="Partially verified"; risk="warn"; }
  else if (isFree)         { l2="Free email provider";        result="Personal address";   risk="warn"; }
  else                     { l2="Auth headers unavailable";   result="Unverified";          risk="warn"; }
  updateCard("card-sender","tick-sender","sender","sender-l1",display,"sender-l2",l2,result,risk);
  if (currentScanData) currentScanData.auth_result = result;
}

function fallbackTrust(isFree, domain, display) {
  let l2, result, risk;
  if (!domain)     { l2="Sender info missing";   result="Unknown sender";   risk="warn"; }
  else if (isFree) { l2="Free email provider";   result="Personal address"; risk="warn"; }
  else             { l2="From "+domain;           result=domain;            risk="safe"; }
  updateCard("card-sender","tick-sender","sender","sender-l1",display,"sender-l2",l2,result,risk);
  if (currentScanData) currentScanData.auth_result = result;
}

// ── Backend ───────────────────────────────────────────────────────
function callBackend(bodyText, hasAttachment) {
  const ctrl = new AbortController();
  const t    = setTimeout(() => ctrl.abort(), 10000);
  fetch(BACKEND+"/classify", {
    method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({text:bodyText, attachment:hasAttachment?"Yes":"No"}),
    signal: ctrl.signal
  })
  .then(r=>r.json())
  .then(d=>{ clearTimeout(t); renderResult((d.label||"unknown").toLowerCase()); logScan((d.label||"unknown").toLowerCase()); })
  .catch(()=>{
    clearTimeout(t);
    const uc = URGENCY_PHRASES.filter(p=>bodyText.toLowerCase().includes(p)).length;
    const hl = /https?:\/\//i.test(bodyText);
    let label="ham";
    if (hl&&uc>=3) label="phishing";
    else if (uc>=2) label="spam";
    else if (hl&&uc) label="spam";
    renderResult(label); logScan(label);
  });
}

// ── Render result ─────────────────────────────────────────────────
function renderResult(label) {
  const score  = RISK_SCORES[label]  || 50;
  const col    = ARC_COLORS[label]   || ARC_COLORS.ham;
  const vd     = VERDICTS[label]     || { t:"Scanned", c:"#94a3b8" };
  const segs   = SEG_COUNTS[label]   || 1;
  const sCls   = { ham:"done", support:"done", spam:"warn", phishing:"danger" };
  const status = { ham:"All clear", support:"Low risk", spam:"Suspicious", phishing:"High risk" };

  // Arc: offset = ARC_LEN - (score/100)*ARC_LEN
  // score 5  → 257 - 12.85 = 244.15 → tiny sliver ✓
  // score 94 → 257 - 241.6  = 15.4  → nearly full ✓
  const arc = document.getElementById("risk-arc");
  if (arc) arc.style.strokeDashoffset = (ARC_LEN - (score/100)*ARC_LEN).toFixed(1);

  // Arc gradient colour
  setArcColor(label);

  // Score number colour
  const num = document.getElementById("score-number");
  if (num) { num.textContent = score; num.setAttribute("fill", col.s); }

  // Colour bar segments
  for (let i=1; i<=5; i++) {
    const s = document.getElementById("seg"+i);
    if (s) { i<=segs ? s.classList.add("active") : s.classList.remove("active"); }
  }

  // Verdict text
  const verdict = document.getElementById("verdict");
  if (verdict) { verdict.textContent = vd.t; verdict.style.color = vd.c; }

  setStatus(status[label]||"Scanned", sCls[label]||"done");
  if (currentScanData) currentScanData.label = label;
}

function setArcColor(label) {
  const c = ARC_COLORS[label]||ARC_COLORS.ham;
  const s = document.getElementById("gs"); if(s) s.setAttribute("stop-color",c.s);
  const e = document.getElementById("ge"); if(e) e.setAttribute("stop-color",c.e);
}

// ── Card helper ───────────────────────────────────────────────────
function updateCard(cardId, tickId, valId, l1Id, l1, l2Id, l2, result, risk) {
  const card = document.getElementById(cardId);
  if (card) card.className = "info-card"+(risk?" c-"+risk:"");
  const tick = document.getElementById(tickId);
  if (tick) {
    tick.className = "card-tick"+(risk?" "+risk:"");
    const pl = tick.querySelector("polyline");
    if (pl) { const cols={safe:"#22c55e",warn:"#fbbf24",danger:"#f87171"}; pl.setAttribute("stroke",cols[risk]||"#94a3b8"); }
  }
  const e1=document.getElementById(l1Id); if(e1) e1.textContent=l1||"";
  const e2=document.getElementById(l2Id); if(e2) e2.textContent=l2||"";
  const val=document.getElementById(valId);
  if(val){ val.textContent=result||"—"; val.className="card-result"+(risk?" "+risk:" neutral"); }
}

function setStatus(msg, cls) {
  const p=document.getElementById("status");
  if(!p) return;
  p.textContent=msg;
  p.className="status-pill"+(cls?" "+cls:"");
}

// ── Report ────────────────────────────────────────────────────────
function reportEmail() {
  if (!currentScanData) return;
  document.getElementById("overlay").classList.remove("hidden");
}
function closeConfirm() {
  document.getElementById("overlay").classList.add("hidden");
}
function confirmReport() {
  closeConfirm();
  const btn=document.getElementById("report-btn");
  btn.disabled=true; btn.textContent="Sending...";
  fetch(BACKEND+"/report",{
    method:"POST",headers:{"Content-Type":"application/json"},
    body:JSON.stringify(Object.assign({},currentScanData,{reported:true}))
  })
  .then(()=>{ btn.textContent="✓ Reported to Sortify"; })
  .catch(()=>{ btn.textContent="Mark as Suspicious"; btn.disabled=false; });
}

function logScan(label) {
  if (!currentScanData) return;
  fetch(BACKEND+"/log-scan",{
    method:"POST",headers:{"Content-Type":"application/json"},
    body:JSON.stringify(Object.assign({},currentScanData,{label}))
  }).catch(()=>{});
}
