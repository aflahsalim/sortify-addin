/* global Office, document */

Office.onReady(() => {
  const item = Office.context.mailbox.item;
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
        showResult({
          score: 0,
          label: "ham",
          sender: "--",
          links: "--",
          content: "No content",
          attachment: hasAttachment ? "Yes" : "No",
        });
        return;
      }

      classifyEmail(emailText, hasAttachment);
    } else {
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
      if (typeof data.attachment === "undefined") {
        data.attachment = hasAttachment ? "Yes" : "No";
      }
      showResult(data);
    })
    .catch((err) => {
      console.error("Fetch error:", err);
      setStatus("Error contacting backend.");
    });
}

function showResult(data) {
  const label = data.label || "unknown";
  const score = Number(data.score) || 0;

  let gaugeColor = "#00FF94"; // default green
  let needleAngle = -90;
  let badgeText = "SAFE";

  if (label === "phishing") {
    gaugeColor = "#FF4B4B";
    needleAngle = 90;
    badgeText = "PHISHING DETECTED";
  } else if (label === "spam") {
    gaugeColor = "#FFA500";
    needleAngle = 45;
    badgeText = "RISK";
  } else if (label === "support") {
    gaugeColor = "#00BFFF";
    needleAngle = 0;
    badgeText = "SUPPORT EMAIL";
  } else if (label === "ham") {
    gaugeColor = "#00FF94";
    needleAngle = -90;
    badgeText = "SAFE";
  }

  // Update gauge
  const needle = document.getElementById("needle");
  if (needle) {
    needle.setAttribute("transform", `rotate(${needleAngle} 100 100)`);
  }
  const arc = document.getElementById("risk-arc");
  if (arc) {
    arc.setAttribute("stroke", gaugeColor);
  }

  // Update risk label
  const scoreEl = document.querySelector(".score-value");
  if (scoreEl) {
    scoreEl.innerText = label === "spam" ? "RISK" : label.toUpperCase();
  }

  // Update badge
  const badge = document.querySelector(".status-badge");
  if (badge) {
    badge.innerText = badgeText;
    badge.className = "status-badge"; // reset classes
    if (label === "phishing") badge.classList.add("status-spam");
    else if (label === "spam") badge.classList.add("status-medium");
    else if (label === "support") badge.classList.add("status-support");
    else badge.classList.add("status-safe");
  }

  // Update analysis details
  document.getElementById("sender").innerText = data.sender || "--";
  document.getElementById("links").innerText = data.links || "--";
  document.getElementById("keywords").innerText = data.content || "--";
  document.getElementById("attachment").innerText = data.attachment || "--";
}

function setStatus(message) {
  const badge = document.querySelector(".status-badge");
  if (badge) {
    badge.innerText = message;
    badge.className = "status-badge status-loading";
  }
}
