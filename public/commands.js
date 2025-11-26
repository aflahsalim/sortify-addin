function sendToSupport(event) {
  const item = Office.context.mailbox.item;
  if (!item) {
    console.error("No email item available.");
    event.completed();
    return;
  }

  // Read the email body
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailBody = result.value || "";
      const subject = item.subject || "Email for Support Review";

      // Show confirmation dialog
      Office.context.ui.displayDialogAsync(
        "https://sortify-addin.onrender.com/confirm.html", // host a simple confirm.html
        { height: 30, width: 40, displayInIframe: true },
        (dialogResult) => {
          const dialog = dialogResult.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (message) => {
            if (message.message === "yes") {
              // Forward email to support team
              Office.context.mailbox.displayNewMessageForm({
                toRecipients: ["aflahsalim.bca@outlook.com"],
                subject: "Sortify Verification Request: " + subject,
                body: `Please verify this email manually:\n\n${emailBody}`
              });
            }
            dialog.close();
            event.completed();
          });

          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
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
