function sendToSupport(event) {
  const item = Office.context.mailbox.item;
  if (!item) {
    console.error("No email item available.");
    event.completed();
    return;
  }

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailBody = result.value;
      const subject = item.subject || "Email for Support Review";

      Office.context.mailbox.item.displayReplyForm({
        toRecipients: ["support@sortify.ai"], // replace with your support email
        subject: "Sortify Verification Request: " + subject,
        htmlBody: `<p>Please verify this email manually:</p><pre>${emailBody}</pre>`
      });
    } else {
      console.error("Failed to read email body.");
    }
    event.completed();
  });
}
