Office.initialize = function () {
  window.sendToSupport = sendToSupport;
};

function sendToSupport(event) {
  const item = Office.context.mailbox.item;
  if (!item) {
    console.error("No email item available.");
    event.completed();
    return;
  }

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailBody = result.value || "";
      const subject = item.subject || "Email for Support Review";

      Office.context.ui.displayDialogAsync(
        "https://sortify-addin.onrender.com/confirm.html",
        { height: 30, width: 40, displayInIframe: true },
        (dialogResult) => {
          const dialog = dialogResult.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (message) => {
            if (typeof message.message === "string" && message.message.toLowerCase() === "yes") {
              Office.context.mailbox.displayNewMessageForm({
                toRecipients: ["aflahsalim.bca@outlook.com"],
                subject: "Sortify Verification Request: " + subject,
                body: `Please verify this email manually:\n\n${emailBody}`
              });
            }
            dialog.close();
            event.completed();
          });

          dialog.addEventHandler(Office.EventType.DialogEventReceived, (event) => {
            console.log("Dialog closed or failed:", event);
            event.completed();
          });
        }
      );
    } else {
      console.error("Failed to read email body.");
      event.completed();
    }
  });
}
