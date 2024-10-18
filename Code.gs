// Constants
const APPROVAL_STATUSES = ["Pending", "Approved", "Denied"];
const COLUMNS = {APPROVAL: 1, EMAIL: 3, ITEM: 4, COST: 5, LINK: 6, PHOTO: 7, REASON: 8};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('[[Menu]]')
    .addItem('Settings', 'setup')
    .addItem('Authorize', 'authorizeScript')
    .addToUi();
}

function setup() {
  try {
    const template = HtmlService.createTemplateFromFile('SetupForm');
    template.data = getSetupData(); 
    const html = template.evaluate();
    console.log('HTML content:', html.getContent()); // Log the HTML content
    SpreadsheetApp.getUi().showSidebar(html.setTitle('Setup'));
  } catch (e) {
    console.error('Error in setup:', e);
    if (e.stack) {
      console.error('Stack trace:', e.stack);
    }
    SpreadsheetApp.getUi().alert('An error occurred while loading the Setup form. Please check the logs for more details.');
  }
}

function onSetupFormLoaded() {
  try {
    authorizationTest();
    createTriggers();
    return { success: true };
  } catch (e) {
    Logger.log('Error in onSetupFormLoaded: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

function authorizationTest() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets();
  GmailApp.getInboxUnreadCount();
  DriveApp.getFiles();
  Drive.Files.list({pageSize: 1});
}

function authorizeScript() {
  try {
    authorizationTest();
    SpreadsheetApp.getUi().alert('Script authorized successfully!');
  } catch (e) {
    Logger.log('Authorization failed: ' + e.toString());
    showReauthorizationPrompt();
  }
}

function showReauthorizationPrompt() {
  SpreadsheetApp.getUi().alert('Authorization Error', 'Please reauthorize the script using the Authorize function in the [[Menu]].', SpreadsheetApp.getUi().ButtonSet.OK);
}

function addApprovalColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getRange(1, 1).getValue() !== "Approval Status") {
    sheet.insertColumnBefore(1).getRange(1, 1).setValue("Approval Status");
    const approvalColumn = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1);
    approvalColumn.setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(APPROVAL_STATUSES, true).setAllowInvalid(false).build()
    ).setValue("Pending");
  }
}

function createTriggers() {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('TRIGGERS_CREATED') !== 'true') {
    const triggers = [
      { funcName: 'onFormSubmit', eventType: ScriptApp.EventType.ON_FORM_SUBMIT },
      { funcName: 'onEditHandler', eventType: ScriptApp.EventType.ON_EDIT }
    ];
    triggers.forEach(({ funcName, eventType }) => {
      if (!ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === funcName)) {
        ScriptApp.newTrigger(funcName)
          .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
          .on(eventType)
          .create();
      }
    });
    properties.setProperty('TRIGGERS_CREATED', 'true');
  }
}

function onFormSubmit(e) {
  try {
    authorizationTest();
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    sheet.getRange(row, COLUMNS.APPROVAL).setValue("Pending");
    const data = getRowData(sheet, row);
    const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    sendNotification(data, true, sheetUrl);
  } catch (error) { 
    Logger.log('Error in onFormSubmit: ' + error.message); 
  }
}

function onEditHandler(e) {
  try {
    authorizationTest();
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    const column = e.range.getColumn();
    if (sheet.getRange(1, 1).getValue() === 'Approval Status' && column === COLUMNS.APPROVAL && row > 1) {
      const approvalStatus = e.range.getValue();
      if (APPROVAL_STATUSES.includes(approvalStatus)) {
        const data = getRowData(sheet, row);
        data.approvalStatus = approvalStatus;
        sendNotification(data, false);
      }
    }
  } catch (error) {
    Logger.log('Error in onEditHandler: ' + error.message);
    showReauthorizationPrompt();
  }
}

function sendNotification(data, isNewRequest, sheetUrl) {
  const properties = PropertiesService.getScriptProperties().getProperties();
  const emailOptions = prepareEmailOptions(data, properties, isNewRequest, sheetUrl);
  GmailApp.sendEmail(emailOptions.to, emailOptions.subject, '', emailOptions);
  
  if (isNewRequest && properties.ENABLE_SMS === 'true' && properties.SUPERVISOR_PHONE_EMAIL) {
    sendSmsNotification(data.item, sheetUrl, properties.SUPERVISOR_PHONE_EMAIL);
  }
}



function generateEmailHTML(data, isRequest, sheetUrl, imageTag = '') {
  const { email, item, cost, link, reason, approvalStatus } = data;
  const title = isRequest ? 'New Purchase Request' : `Your Purchase Request Has Been ${approvalStatus}`;
  const footer = isRequest ? `<a href="${sheetUrl}" style="color:#1a73e8;">Click here to open the Purchase Request spreadsheet.</a>` : '';
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body{font-family:'Ubuntu',sans-serif;margin:0;padding:0;background-color:#f5f5f5;}
        .email-container{max-width:600px;margin:auto;background-color:#fff;border:1px solid #ddd;border-radius:4px;}
        .email-header{background-color:#333;color:#fff;padding:20px;text-align:center;}
        .email-content{padding:20px;}
        .email-content table{width:100%;border-collapse:collapse;}
        .email-content th,.email-content td{padding:10px;text-align:left;vertical-align:top;border-bottom:1px solid #eee;}
        .email-content th{background-color:#f0f0f0;}
        .email-footer{padding:10px 20px;background-color:#f0f0f0;text-align:center;font-size:12px;color:#888;}
      </style>
    </head>
    <body>
      <div class="email-container">
        <div class="email-header"><h2>${title}</h2></div>
        <div class="email-content">
          <table>
            ${isRequest ? `<tr><th>Requestor:</th><td>${email}</td></tr>` : ''}
            <tr><th>Item:</th><td>${item || 'Not specified'}</td></tr>
            <tr><th>Estimated Cost:</th><td>${cost ? `$${cost}` : 'Not specified'}</td></tr>
            <tr><th>Item Link:</th><td>${link ? `<a href="${link}" style="color:#1a73e8;">${link}</a>` : 'Not provided'}</td></tr>
            <tr><th>Reason for Purchase:</th><td>${reason || 'Not provided'}</td></tr>
            <tr><th>Status:</th><td><em>${isRequest ? 'Pending' : approvalStatus}</em></td></tr>
            ${imageTag ? `<tr><td colspan="2" style="text-align:center;">${imageTag}</td></tr>` : ''}
          </table>
        </div>
        <div class="email-footer">${footer}</div>
      </div>
    </body>
    </html>
  `;
}

function sendSmsNotification(requestedItem, sheetUrl, phoneEmail) {
  let smsMessage = `[[FORM]] Purchase Request - ${requestedItem} ${sheetUrl}`;
  GmailApp.sendEmail(phoneEmail, "", smsMessage);
}

function prepareEmailOptions(data, properties, isNewRequest, sheetUrl) {
  const imageTag = generateImageTag(data.photo);
  const htmlMessage = generateEmailHTML(data, isNewRequest, sheetUrl, imageTag);
  const emailOptions = {
    to: isNewRequest ? properties.SUPERVISOR_EMAIL : data.email,
    subject: `[[Form]] ${data.item || 'Purchase Request'} - ${isNewRequest ? 'New Request' : data.approvalStatus}`,
    htmlBody: htmlMessage,
    attachments: []
  };
  const ccList = [];
  if (isNewRequest ? properties.CC_EMPLOYEE === 'true' : properties.CC_SUPERVISOR === 'true') {
    ccList.push(isNewRequest ? data.email : properties.SUPERVISOR_EMAIL);
  }
  if (properties.ENABLE_THIRD_PARTY_EMAIL === 'true' && properties.THIRD_PARTY_EMAIL) {
    ccList.push(properties.THIRD_PARTY_EMAIL);
  }
  if (ccList.length > 0) {
    emailOptions.cc = ccList.join(", ");
  }
  return emailOptions;
}

function getRowData(sheet, row) {
  return {
    email: sheet.getRange(row, COLUMNS.EMAIL).getValue(),
    item: sheet.getRange(row, COLUMNS.ITEM).getValue(),
    cost: sheet.getRange(row, COLUMNS.COST).getValue(),
    link: sheet.getRange(row, COLUMNS.LINK).getValue(),
    photo: sheet.getRange(row, COLUMNS.PHOTO).getValue(),
    reason: sheet.getRange(row, COLUMNS.REASON).getValue()
  };
}

const isValidEmail = email => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);

function generateImageTag(photoUpload) {
  if (!photoUpload) return '';
  try {
    const fileIdMatch = photoUpload.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
    if (!fileIdMatch) {
      throw new Error('Invalid file ID format');
    }
    const fileId = fileIdMatch[0];
    const file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileMetadata = Drive.Files.get(fileId, { fields: 'thumbnailLink,webContentLink,mimeType' });
    let imageUrl;
    if (fileMetadata.thumbnailLink) {
      // Use the largest thumbnail size available
      imageUrl = fileMetadata.thumbnailLink.replace(/=s\d+/, '=s1000');
    } else if (fileMetadata.webContentLink) {
      imageUrl = fileMetadata.webContentLink;
    } else {
      throw new Error('No valid image URL found');
    }
    const mimeType = fileMetadata.mimeType;
    if (mimeType.startsWith('image/')) {
      return `<img src="${imageUrl}" alt="Uploaded image" style="max-width:100%;height:auto;border:1px solid #ddd;border-radius:4px;">`;
    } else if (mimeType === 'application/pdf') {
      return `<a href="${imageUrl}" target="_blank" style="display:inline-block;padding:10px;background-color:#f0f0f0;border-radius:4px;text-decoration:none;color:#333;">View PDF Document</a>`;
    } else {
      return `<a href="${imageUrl}" target="_blank" style="display:inline-block;padding:10px;background-color:#f0f0f0;border-radius:4px;text-decoration:none;color:#333;">View Uploaded File</a>`;
    }
  } catch (error) {
    Logger.log('Error generating image tag: ' + error.message);
    return `<p>Unable to display the uploaded file. <a href="${photoUpload}" target="_blank" style="color:#1a73e8;">Click here to view the original file</a>.</p>`;
  }
}


function saveSetupData(formData) {
  const currentData = getSetupData();
  const updatedData = {...currentData, ...formData};
  
  const errors = validateSetupData(updatedData);
  if (errors.length > 0) throw new Error(errors.join('\n'));

  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'SUPERVISOR_EMAIL': updatedData.SUPERVISOR_EMAIL,
    'ENABLE_SMS': updatedData.ENABLE_SMS.toString(),
    'SUPERVISOR_PHONE_EMAIL': updatedData.ENABLE_SMS ? updatedData.SUPERVISOR_PHONE_EMAIL : '',
    'CC_SUPERVISOR': updatedData.CC_SUPERVISOR.toString(),
    'CC_EMPLOYEE': updatedData.CC_EMPLOYEE.toString(),
    'ENABLE_THIRD_PARTY_EMAIL': updatedData.ENABLE_THIRD_PARTY_EMAIL.toString(),
    'THIRD_PARTY_EMAIL': updatedData.ENABLE_THIRD_PARTY_EMAIL ? updatedData.THIRD_PARTY_EMAIL : '',
  });

  if (properties.getProperty('SETUP_COMPLETED') !== 'true') {
    addApprovalColumn();
    createTriggers();
    properties.setProperty('SETUP_COMPLETED', 'true');
  }
}

function getSetupData() {
  const properties = PropertiesService.getScriptProperties();
  return {
    SUPERVISOR_EMAIL: properties.getProperty('SUPERVISOR_EMAIL') || '',
    ENABLE_SMS: properties.getProperty('ENABLE_SMS') === 'true',
    SUPERVISOR_PHONE_EMAIL: properties.getProperty('SUPERVISOR_PHONE_EMAIL') || '',
    CC_SUPERVISOR: properties.getProperty('CC_SUPERVISOR') === 'true',
    CC_EMPLOYEE: properties.getProperty('CC_EMPLOYEE') === 'true',
    ENABLE_THIRD_PARTY_EMAIL: properties.getProperty('ENABLE_THIRD_PARTY_EMAIL') === 'true',
    THIRD_PARTY_EMAIL: properties.getProperty('THIRD_PARTY_EMAIL') || ''
  };
}

function validateSetupData(data) {
  const errors = [];
  if (data.SUPERVISOR_EMAIL && !isValidEmail(data.SUPERVISOR_EMAIL)) errors.push('Invalid Supervisor Email.');
  if (data.ENABLE_SMS && data.SUPERVISOR_PHONE_EMAIL && !isValidEmail(data.SUPERVISOR_PHONE_EMAIL)) errors.push('Invalid Supervisor Phone Email.');
  if (data.ENABLE_THIRD_PARTY_EMAIL && data.THIRD_PARTY_EMAIL && !isValidEmail(data.THIRD_PARTY_EMAIL)) errors.push('Invalid Third-Party Email.');
  return errors;
}
