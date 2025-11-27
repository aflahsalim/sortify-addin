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

  // Get subject
  const subject = item.subject || "Email for Support Review";

  // Get body text
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailBody = result.value || "";

      // Collect attachments if any
      const attachments = item.attachments || [];

      // Show confirmation popup
      Office.context.ui.displayDialogAsync(
        "https://sortify-addin.onrender.com/confirm.html",
        { height: 30, width: 40, displayInIframe: true },
        (dialogResult) => {
          const dialog = dialogResult.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (message) => {
            if (typeof message.message === "string" && message.message.toLowerCase() === "yes") {
              // Build new email form
              Office.context.mailbox.displayNewMessageForm({
                toRecipients: ["aflahsalim.bca@outlook.com"],
                subject: "Sortify Verification Request: " + subject,
                htmlBody: `
                  <p>Hi Sortify Support,</p>
                  <p>Please verify this email manually:</p>
                  <hr>
                  <p><strong>Original Subject:</strong> ${subject}</p>
                  <p><strong>Original Body:</strong></p>
                  <pre>${emailBody}</pre>
                  <hr>
                  <p>Thanks,<br/>Sortify User</p>
                `,
                attachments: attachments.map(att => ({
                  type: "file",
                  name: att.name,
                  url: att.url
                }))
              });
            }
            dialog.close();
            event.completed();
          });

          dialog.addEventHandler(Office.EventType.DialogEventReceived, (evt) => {
            console.log("Dialog closed or failed:", evt);
            event.completed();
          });
        }
      );
    } else {
      console.error("Failed to read email body:", result.error);
      event.completed();
    }
  });
}
