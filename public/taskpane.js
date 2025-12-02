Office.onReady(() => {
  // Office is ready
});

function showResult(item) {
  // ✅ Get sender email with fallback
  const senderEmail =
    item?.from?.emailAddress?.address ||
    item?.sender?.emailAddress?.address ||
    "";

  const senderDomain = senderEmail.split("@")[1] || "";
  const isFreeDomain = ["gmail.com", "yahoo.com", "outlook.com", "hotmail.com"].includes(
    senderDomain?.toLowerCase()
  );

  const senderReputation = senderEmail
    ? (isFreeDomain ? "Suspicious" : "Trusted")
    : "Unknown";

  // ✅ Update UI
  document.getElementById("sender").textContent = senderReputation;

  // Hyperlink assessment
  const body = item?.body?.text || "";
  const hasLinks = body.includes("http://") || body.includes("https://");
  document.getElementById("links").textContent = hasLinks ? "Detected" : "None";

  // File assessment
  const attachments = item?.attachments || [];
  document.getElementById("attachment").textContent = attachments.length > 0 ? "Found" : "None";

  // Urgency assessment
  const urgencyKeywords = ["urgent", "immediately", "critical", "asap"];
  const urgencyLevel = urgencyKeywords.some((kw) => body.toLowerCase().includes(kw))
    ? "Critical"
    : "Normal";
  document.getElementById("urgency").textContent = urgencyLevel;

  // ✅ Gauge + badge update
  let riskScore = 0;
  if (senderReputation === "Suspicious") riskScore += 40;
  if (hasLinks) riskScore += 30;
  if (attachments.length > 0) riskScore += 20;
  if (urgencyLevel === "Critical") riskScore += 30;

  const arc = document.getElementById("risk-arc");
  const needle = document.getElementById("needle");
  const scoreLabel = document.getElementById("score-label");
  const resultButton = document.getElementById("result-button");

  const maxArc = 235;
  const offset = Math.max(maxArc - (riskScore / 100) * maxArc, 0);
  arc.style.strokeDashoffset = offset;

  const angle = -90 + (riskScore / 100) * 180;
  needle.style.transform = `rotate(${angle} 100 90)`;

  scoreLabel.textContent = `${riskScore}%`;

  if (riskScore < 25) {
    arc.setAttribute("stroke", "#28a745");
    resultButton.textContent = "Safe";
    resultButton.style.background = "#28a745";
  } else if (riskScore < 50) {
    arc.setAttribute("stroke", "#00bfff");
    resultButton.textContent = "Support";
    resultButton.style.background = "#00bfff";
  } else if (riskScore < 75) {
    arc.setAttribute("stroke", "#fd7e14");
    resultButton.textContent = "Spam";
    resultButton.style.background = "#fd7e14";
  } else {
    arc.setAttribute("stroke", "#dc3545");
    resultButton.textContent = "Phishing";
    resultButton.style.background = "#dc3545";
  }
}

// ✅ Run when item is ready
Office.context.mailbox.item && showResult(Office.context.mailbox.item);
