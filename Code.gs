// Setup the system: adds approval column, creates necessary triggers, asks for additional options
function setup() {
  const htmlService = HtmlService.createHtmlOutputFromFile('SetupForm');
  SpreadsheetApp.getUi().showSidebar(htmlService);
}

// Add Approval Status column to the sheet
function addApprovalColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headerRange = sheet.getRange(1, 1); // Get the first header row
  if (headerRange.getValue() !== "Approval Status") {
    sheet.insertColumnBefore(1); // Insert a new column A
    sheet.getRange(1, 1).setValue("Approval Status"); // Set the new header
  }
}

// Create triggers for form submission and approval notifications
function createTriggers() {
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  
  // Time-driven trigger to check approvals periodically
  ScriptApp.newTrigger("sendApprovalNotification")
    .timeBased()
    .everyHours(1) // Runs every hour
    .create();
}

// Create custom menu for running the setup manually
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Setup Menu')
    .addItem('Run Setup', 'setup')
    .addToUi();
}

// Function to notify the supervisor of a new request with inline CSS optimized for mobile
function sendRequestNotification(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  const supervisorEmail = PropertiesService.getScriptProperties().getProperty('SUPERVISOR_EMAIL');
  const supervisorPhoneEmail = PropertiesService.getScriptProperties().getProperty('SUPERVISOR_PHONE_EMAIL');
  const ccSupervisor = PropertiesService.getScriptProperties().getProperty('CC_SUPERVISOR') === 'true';
  const employeeEmail = sheet.getRange(lastRow, 2).getValue(); // Employee's email (column 2)
  const requestedItem = sheet.getRange(lastRow, 3).getValue(); // Item requested (column 3)
  const estimatedCost = sheet.getRange(lastRow, 4).getValue(); // Estimated cost (column 4)
  const itemLink = sheet.getRange(lastRow, 5).getValue(); // Link to the item (column 5)
  const photoUpload = sheet.getRange(lastRow, 6).getValue(); // Photo upload link (column 6)
  const purchaseReason = sheet.getRange(lastRow, 7).getValue(); // Reason for purchase (column 7)
  const thirdPartyEmail = PropertiesService.getScriptProperties().getProperty('THIRD_PARTY_EMAIL');
  const responseSheetLink = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  // Insert the uploaded image if provided
  const imageTag = photoUpload ? `<p style="text-align:center;"><img src="${photoUpload}" alt="Uploaded image" style="max-width:100%; height:auto;"></p>` : '';

  // Create enhanced HTML message with inline CSS optimized for mobile
  const htmlMessage = `
    <div style="font-family:Arial, sans-serif; line-height:1.5; padding:10px; background-color:#333333; color:#f0f0f0;">
      <h2 style="color:#ffffff;">New Purchase Request</h2>
      <p><strong>Employee:</strong> ${employeeEmail}</p>
      <p><strong>Item:</strong> ${requestedItem}</p>
      <p><strong>Estimated Cost:</strong> $${estimatedCost}</p>
      <p><strong>Item Link:</strong> <a href="${itemLink}" style="color:#1a73e8;">${itemLink}</a></p>
      ${imageTag}
      <p><strong>Reason for Purchase:</strong> ${purchaseReason}</p>
      <p><strong>Approval Status:</strong> <em>Pending</em></p>
    </div>
  `;
  
  // Prepare the email options
  let emailOptions = {
    to: supervisorEmail,
    subject: `New Purchase Request from ${employeeEmail} - Action Required`,
    htmlBody: htmlMessage
  };
  
  if (ccSupervisor || thirdPartyEmail) {
    emailOptions.cc = [];
    if (ccSupervisor) emailOptions.cc.push(employeeEmail); // CC the employee if selected
    if (thirdPartyEmail) emailOptions.cc.push(thirdPartyEmail); // CC the third-party email if provided
    emailOptions.cc = emailOptions.cc.join(", ");
  }
  
  // Send the email to the supervisor
  MailApp.sendEmail(emailOptions);

  // Send text notification if enabled
  if (supervisorPhoneEmail) {
    MailApp.sendEmail({
      to: supervisorPhoneEmail,
      subject: "New Purchase Request",
      body: `New request from ${employeeEmail} for ${requestedItem}. Please check your email or the responses sheet for details: ${responseSheetLink}`
    });
  }
}

// Trigger: Send email on form submission
function onFormSubmit(e) {
  sendRequestNotification(e);
}

// Function to send email notification when request is approved
function sendApprovalNotification() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const approvalStatus = sheet.getRange(lastRow, 1).getValue(); // Approval status in column A (1st column)
  
  const employeeEmail = sheet.getRange(lastRow, 2).getValue(); // Email (column 2)
  const requestedItem = sheet.getRange(lastRow, 3).getValue(); // Item (column 3)
  const estimatedCost = sheet.getRange(lastRow, 4).getValue(); // Cost (column 4)
  const itemLink = sheet.getRange(lastRow, 5).getValue(); // Link (column 5)
  const photoUpload = sheet.getRange(lastRow, 6).getValue(); // Photo upload (column 6)
  const purchaseReason = sheet.getRange(lastRow, 7).getValue(); // Reason (column 7)
  const thirdPartyEmail = PropertiesService.getScriptProperties().getProperty('THIRD_PARTY_EMAIL');

  // Insert the uploaded image if provided
  const imageTag = photoUpload ? `<p style="text-align:center;"><img src="${photoUpload}" alt="Uploaded image" style="max-width:100%; height:auto;"></p>` : '';

  if (approvalStatus === "Approved") {
    // Create HTML message for the employee with inline CSS optimized for mobile
    const htmlMessage = `
      <div style="font-family:Arial, sans-serif; line-height:1.5; padding:10px; background-color:#333333; color:#f0f0f0;">
        <h2 style="color:#ffffff;">Your Purchase Request Has Been Approved</h2>
        <p><strong>Item:</strong> ${requestedItem}</p>
        <p><strong>Estimated Cost:</strong> $${estimatedCost}</p>
        <p><strong>Item Link:</strong> <a href="${itemLink}" style="color:#1a73e8;">${itemLink}</a></p>
        ${imageTag}
        <p><strong>Reason for Purchase:</strong> ${purchaseReason}</p>
        <p><strong>Status:</strong> <em>Approved</em></p>
      </div>
    `;
    
    // Prepare email options
    let emailOptions = {
      to: employeeEmail,
      subject: "Your Purchase Request Has Been Approved",
      htmlBody: htmlMessage
    };

    if (PropertiesService.getScriptProperties().getProperty('CC_EMPLOYEE') === 'true' || thirdPartyEmail) {
      emailOptions.cc = [];
      if (PropertiesService.getScriptProperties().getProperty('CC_EMPLOYEE') === 'true') emailOptions.cc.push(PropertiesService.getScriptProperties().getProperty('SUPERVISOR_EMAIL')); // CC the supervisor if selected
      if (thirdPartyEmail) emailOptions.cc.push(thirdPartyEmail); // CC the third-party email if provided
      emailOptions.cc = emailOptions.cc.join(", ");
    }
    
    // Send the email to the employee
    MailApp.sendEmail(emailOptions);
  }
}

// Function to save setup data from the HTML form
function saveSetupData(supervisorEmail, supervisorPhoneEmail, ccSupervisor, ccEmployee, thirdPartyEmail) {
  PropertiesService.getScriptProperties().setProperty('SUPERVISOR_EMAIL', supervisorEmail);
  PropertiesService.getScriptProperties().setProperty('SUPERVISOR_PHONE_EMAIL', supervisorPhoneEmail);
  PropertiesService.getScriptProperties().setProperty('CC_SUPERVISOR', ccSupervisor.toString());
  PropertiesService.getScriptProperties().setProperty('CC_EMPLOYEE', ccEmployee.toString());
  if (thirdPartyEmail) {
    PropertiesService.getScriptProperties().setProperty('THIRD_PARTY_EMAIL', thirdPartyEmail);
  } else {
    PropertiesService.getScriptProperties().deleteProperty('THIRD_PARTY_EMAIL');
  }
}
