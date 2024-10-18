function sendRequestNotification(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  const bossEmail = "boss@company.com"; // Boss' email
  const employeeEmail = sheet.getRange(lastRow, 2).getValue(); // Employee's email (column 2)
  const item = sheet.getRange(lastRow, 3).getValue(); // Item requested (column 3)
  const estimatedCost = sheet.getRange(lastRow, 4).getValue(); // Estimated cost (column 4)
  const itemLink = sheet.getRange(lastRow, 5).getValue(); // Link to the item (column 5)
  const photoUpload = sheet.getRange(lastRow, 6).getValue(); // Photo upload link (column 6)
  const reason = sheet.getRange(lastRow, 7).getValue(); // Reason for purchase (column 7)
  
  // Create HTML summary for the boss
  const htmlMessage = `
    <p><strong>New Purchase Request</strong></p>
    <p><strong>Employee:</strong> ${employeeEmail}</p>
    <p><strong>Item:</strong> ${item}</p>
    <p><strong>Estimated Cost:</strong> ${estimatedCost}</p>
    <p><strong>Item Link:</strong> <a href="${itemLink}">${itemLink}</a></p>
    <p><strong>Photo Upload:</strong> <a href="${photoUpload}">Uploaded File</a></p>
    <p><strong>Reason for Purchase:</strong> ${reason}</p>
    <p><strong>Approval Link:</strong> <a href="${sheet.getUrl()}">Approve/Deny Request</a></p>
  `;
  
  // Send the email to the boss
  MailApp.sendEmail({
    to: bossEmail,
    subject: "New Purchase Request Submitted",
    htmlBody: htmlMessage
  });
}

// Trigger: Send email on form submission
function onFormSubmit(e) {
  sendRequestNotification(e);
}

// Function to send email notification when request is approved
function sendApprovalNotification() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const approvalStatus = sheet.getRange(lastRow, 8).getValue(); // Approval status in column 8
  
  if (approvalStatus === "Approved") {
    const employeeEmail = sheet.getRange(lastRow, 2).getValue(); // Email (column 2)
    const item = sheet.getRange(lastRow, 3).getValue(); // Item (column 3)
    const estimatedCost = sheet.getRange(lastRow, 4).getValue(); // Cost (column 4)
    const itemLink = sheet.getRange(lastRow, 5).getValue(); // Link (column 5)
    const photoUpload = sheet.getRange(lastRow, 6).getValue(); // Photo upload (column 6)
    const reason = sheet.getRange(lastRow, 7).getValue(); // Reason (column 7)
    
    // Send an email to the employee about approval
    const htmlMessage = `
      <p>Your purchase request has been <strong>Approved</strong>.</p>
      <p><strong>Item:</strong> ${item}</p>
      <p><strong>Estimated Cost:</strong> ${estimatedCost}</p>
      <p><strong>Item Link:</strong> <a href="${itemLink}">${itemLink}</a></p>
      <p><strong>Photo Upload:</strong> <a href="${photoUpload}">Uploaded File</a></p>
      <p><strong>Reason for Purchase:</strong> ${reason}</p>
    `;
    
    MailApp.sendEmail({
      to: employeeEmail,
      subject: "Your Purchase Request Has Been Approved",
      htmlBody: htmlMessage
    });
  }
}
