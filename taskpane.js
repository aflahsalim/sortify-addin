Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "block";

  const item = Office.context.mailbox.item;
  if (item) {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        classifyEmail(result.value);
      } else {
        document.getElementById("result").innerText = "Failed to read email body.";
      }
    });
  }
});

function classifyEmail(emailText) {
  fetch("https://sortify-backend.onrender.com/classify", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text: emailText })
  })
  .then(res => res.json())
  .then(data => showResult(data))
  .catch(err => {
    document.getElementById("result").innerText = "Error contacting backend.";
    console.error(err);
  });
}

function showResult(data) {
  document.getElementById("result").innerHTML =
    `<h3>Phishing Score: ${data.score}</h3><p>${data.label}</p>`;
}
