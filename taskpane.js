const BACKEND = "https://sortify-backend-hwf9d0exgqdub9cn.canadacentral-01.azurewebsites.net";

const RISKY_EXT = ["exe","bat","vbs","js","msi","ps1","cmd","scr","zip","rar","docm","xlsm"];
const URGENCY_PHRASES = [
  "urgent","immediately","suspended","verify","confirm your","click now",
  "24 hours","48 hours","act now","account locked","unusual activity",
  "security alert","update your password","will be terminated","limited time"
];

const RISK_SCORES = { ham: 95, support: 88, spam: 62, phishing: 21 };
const VERDICTS = {
  ham:      { text: "Safe — Full Check Passed",   color: "#22c55e", status: "safe" },
  support:  { text: "Low Risk — Support Email",   color: "#06b6d4", status: "safe" },
  spam:     { text: "Suspicious — Possible Spam", color: "#f59e0b", status: "warn" },
  phishing: { text: "High Risk — Likely Phishing",color: "#ef4444", status: "danger" },
};

let currentScan = null;

Office.onReady(() => {
  waitForDOM(() => {
    wireButton();
    if (Office.context?.mailbox?.item) {
      resetUI();
      startClassification();
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        () => { resetUI(); startClassification(); }
      );
    } else {
      setStatus("Open an email to scan", "");
    }
  });
});

function waitForDOM(cb) {
  if (document.getElementById("risk-arc")) cb();
  else requestAnimationFrame(() => waitForDOM(cb));
}

function wireButton() {
  const btn = document.getElementById("report-btn");
  if (!btn) return;
  btn.addEventListener("click", () => {
    if (!currentScan) return;
    btn.disabled = true;
    btn.textContent = "Reporting…";
    fetch(BACKEND + "/report", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(Object.assign({}, currentScan, { reported: true }))
    })
    .then(() => { btn.textContent = "✓ Reported to Sortify"; })
    .catch(() => { btn.disabled = false; btn.textContent = "Mark as Suspicious"; });
  });
}

function resetUI() {
  setStatus("Scanning…", "");
  setGauge(0.0, "--", "#22c55e");
  setVerdict("Analysing email…", "#f9fafb");

  setCard("sender",     "--", "Checking authentication…", "");
  setCard("links",      "--", "No analysis yet", "");
  setCard("attachment", "--", "No analysis yet", "");
  setCard("urgency",    "--", "No analysis yet", "");

  currentScan = null;
}

function startClassification() {
  const item = Office.context?.mailbox?.item;
  if (!item) return setStatus("Open an email to scan", "");

  setStatus("Reading email…", "");
  item.body.getAsync(Office.CoercionType.Text, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      return setStatus("Could not read email", "warn");
    }
    const body = res.value || "";
    processEmail(item, body);
  });
}

function processEmail(item, body) {
  // Sender
  let senderEmail = "";
  let senderName = "";
  try {
    senderEmail = item?.from?.emailAddress?.address || "";
    senderName  = item?.from?.emailAddress?.displayName || "";
  } catch(e) {}
  const domain = senderEmail.split("@")[1] || "";
  const freeDomains = ["gmail.com","yahoo.com","hotmail.com","outlook.com","live.com"];
  const isFree = freeDomains.includes(domain.toLowerCase());
  const senderDisplay = senderName || senderEmail || "Unknown sender";

  // Attachments
  const attachments = Array.isArray(item.attachments) ? item.attachments : [];
  let aLabel = "No files attached";
  let aSub   = "";
  let aRisk  = "";
  if (attachments.length > 0) {
    const risky = attachments.find(a => {
      const ext = (a.name || "").split(".").pop().toLowerCase();
      return RISKY_EXT.includes(ext);
    });
    if (risky) {
      aLabel = risky.name;
      aSub   = "Executable / macro file type";
      aRisk  = "danger";
    } else {
      aLabel = attachments.length + " attachment(s)";
      aSub   = "No risky file types detected";
      aRisk  = "safe";
    }
  }
  setCard("attachment", aLabel, aSub, aRisk);

  // Links
  const urls = (body.match(/(https?:\/\/[^\s]+)/gi) || []);
  let lLabel = "No links found";
  let lSub   = "Email contains no URLs";
  let lRisk  = "";
  if (urls.length > 0) {
    lLabel = urls.length + " link(s)";
    lSub   = "Links present in body";
    lRisk  = "warn";
  }
  setCard("links", lLabel, lSub, lRisk);

  // Urgency / behaviour
  const lower = body.toLowerCase();
  const matched = URGENCY_PHRASES.filter(p => lower.includes(p));
  let uLabel = "Normal behaviour";
  let uSub   = "No phishing tactics detected";
  let uRisk  = "";
  if (matched.length >= 3) {
    uLabel = "High pressure tactics";
    uSub   = matched.length + " phishing signals found";
    uRisk  = "danger";
  } else if (matched.length >= 1) {
    uLabel = "Mild pressure language";
    uSub   = '"' + matched[0] + '" detected';
    uRisk  = "warn";
  }
  setCard("urgency", uLabel, uSub, uRisk);

  // Sender card (auth will refine)
  setCard("sender", senderDisplay, "Verifying domain & headers…", "");

  currentScan = {
    sender: senderEmail,
    subject: item.subject || "",
    domain,
    attachment_count: attachments.length,
    body_preview: body.substring(0, 300),
    label: "unknown"
  };

  checkAuth(item, senderDisplay, isFree, domain);
  callBackend(body, attachments.length > 0);
}

function checkAuth(item, display, isFree, domain) {
  if (typeof item.getAllInternetHeadersAsync === "function") {
    const t = setTimeout(() => fallbackAuth(display, isFree, domain), 4000);
    item.getAllInternetHeadersAsync((r) => {
      clearTimeout(t);
      if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
        parseAuth(r.value, display, isFree, domain);
      } else {
        fallbackAuth(display, isFree, domain);
      }
    });
  } else {
    fallbackAuth(display, isFree, domain);
  }
}

function parseAuth(raw, display, isFree, domain) {
  const h = raw.toLowerCase();
  const n = [h.includes("spf=pass"), h.includes("dkim=pass"), h.includes("dmarc=pass")].filter(Boolean).length;
  let label, sub, risk;
  if (n === 3 && !isFree) {
    label = display;
    sub   = "Verified by SPF / DKIM / DMARC";
    risk  = "safe";
  } else if (n >= 2 && !isFree) {
    label = display;
    sub   = n + "/3 auth checks passed";
    risk  = "safe";
  } else if (n >= 1) {
    label = display;
    sub   = n + "/3 auth checks passed";
    risk  = "warn";
  } else if (isFree) {
    label = display;
    sub   = "Free email provider";
    risk  = "warn";
  } else {
    label = display;
    sub   = "Unverified sender";
    risk  = "warn";
  }
  setCard("sender", label, sub, risk);
}

function fallbackAuth(display, isFree, domain) {
  let sub, risk;
  if (!domain) {
    sub  = "Sender info missing";
    risk = "warn";
  } else if (isFree) {
    sub  = "Free email provider";
    risk = "warn";
  } else {
    sub  = "From " + domain;
    risk = "safe";
  }
  setCard("sender", display, sub, risk);
}

function callBackend(bodyText, hasAttachment) {
  setStatus("Classifying email…", "");
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), 10000);

  fetch(BACKEND + "/classify", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text: bodyText, attachment: hasAttachment ? "Yes" : "No" }),
    signal: ctrl.signal
  })
  .then(r => r.json())
  .then(d => {
    clearTimeout(t);
    const label = String(d.label || "unknown").toLowerCase();
    renderResult(label);
    logScan(label);
  })
  .catch(() => {
    clearTimeout(t);
    const lower = bodyText.toLowerCase();
    const uc = URGENCY_PHRASES.filter(p => lower.includes(p)).length;
    const hl = /https?:\/\//i.test(bodyText);
    let label = "ham";
    if (hl && uc >= 3) label = "phishing";
    else if (uc >= 2)  label = "spam";
    else if (hl && uc) label = "spam";
    renderResult(label);
    logScan(label);
  });
}

function renderResult(label) {
  const v = VERDICTS[label] || {
    text: "Scanned — Unknown verdict",
    color: "#e5e7eb",
    status: ""
  };
  const score = RISK_SCORES[label] ?? 50;
  const fill  = label === "phishing" ? 0.2 :
                label === "spam"     ? 0.45 :
                label === "support"  ? 0.78 :
                0.9; // ham default

  setGauge(fill, score, v.color);
  setVerdict(v.text, v.color);
  setStatus(v.text.split("—")[0].trim(), v.status);

  if (currentScan) currentScan.label = label;
}

function setGauge(fillRatio, score, color) {
  const arc = document.getElementById("risk-arc");
  if (arc) {
    const total = 245;
    arc.style.strokeDashoffset = total - fillRatio * total;
    arc.style.stroke = color;
  }
  const num = document.getElementById("score-number");
  if (num) num.textContent = typeof score === "number" ? score : "--";
}

function setVerdict(text, color) {
  const v = document.getElementById("verdict");
  if (v) {
    v.textContent = text;
    v.style.color = color || "#f9fafb";
  }
}

function setStatus(text, cls) {
  const pill = document.getElementById("status-pill");
  if (!pill) return;
  pill.textContent = text;
  pill.className = "status-pill" + (cls ? " " + cls : "");
}

function setCard(key, value, sub, risk) {
  const valEl = document.getElementById(key);
  const subEl = document.getElementById(key + "-sub");
  const dotEl = document.getElementById(key + "-dot");
  if (valEl) valEl.textContent = value || "--";
  if (subEl) subEl.textContent = sub || "";
  if (dotEl) {
    dotEl.className = "detail-dot";
    if (risk) dotEl.classList.add(risk);
  }
}

function logScan(label) {
  if (!currentScan) return;
  fetch(BACKEND + "/log-scan", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(Object.assign({}, currentScan, { label }))
  }).catch(() => {});
}
