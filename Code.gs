const APPROVAL_STATUSES = ["Pending", "Approved", "Denied"];
const COLUMNS = {APPROVAL: 1, EMAIL: 3, ITEM: 4, COST: 5, LINK: 6, PHOTO: 7, REASON: 8};

function onOpen() {
  try {
    authorizationTest();
    SpreadsheetApp.getUi().createMenu('[[Menu]]').addItem('Settings', 'setup').addToUi();
  } catch (e) { Logger.log('Error in onOpen: ' + e.toString()); }
}

function setup() {
  const template = HtmlService.createTemplateFromFile('SetupForm');
  template.data = getSetupData() || {};
  SpreadsheetApp.getUi().showSidebar(template.evaluate().setTitle('Setup'));
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
  const triggers = {onFormSubmit: ScriptApp.EventType.ON_FORM_SUBMIT, onEdit: ScriptApp.EventType.ON_EDIT};
  Object.entries(triggers).forEach(([funcName, eventType]) => {
    if (!ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === funcName)) {
      ScriptApp.newTrigger(funcName).forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).on(eventType).create();
    }
  });
}

function sendRequestNotification(e) {
  try {
    const sheet = e.range.getSheet(), row = e.range.getRow(), data = getRowData(sheet, row);
    const properties = PropertiesService.getScriptProperties().getProperties();
    const shortUrl = shortenUrl(SpreadsheetApp.getActiveSpreadsheet().getUrl());
    const emailOptions = prepareEmailOptions(data, properties, true, shortUrl);
    GmailApp.sendEmail(emailOptions.to, emailOptions.subject, '', emailOptions);
    if (properties.ENABLE_SMS === 'true' && properties.SUPERVISOR_PHONE_EMAIL) {
      sendSmsNotification(data.requestedItem, shortUrl, properties.SUPERVISOR_PHONE_EMAIL);
    }
  } catch (error) { Logger.log('Error sending request notification: ' + error.message); }
}

function onFormSubmit(e) {
  try {
    authorizationTest();
    e.range.getSheet().getRange(e.range.getRow(), COLUMNS.APPROVAL).setValue("Pending");
    sendRequestNotification(e);
  } catch (error) { Logger.log('Error in onFormSubmit: ' + error.message); }
}

function onEdit(e) {
  try {
    authorizationTest();
    const sheet = e.range.getSheet(), row = e.range.getRow(), column = e.range.getColumn();
    if (sheet.getRange(1, 1).getValue() === 'Approval Status' && column === COLUMNS.APPROVAL && row > 1) {
      const approvalStatus = e.range.getValue();
      if (APPROVAL_STATUSES.includes(approvalStatus)) {
        const data = getRowData(sheet, row);
        data.approvalStatus = approvalStatus;
        const properties = PropertiesService.getScriptProperties().getProperties();
        GmailApp.sendEmail(data.employeeEmail, `[[Form]] ${data.requestedItem} - ${approvalStatus}`, '', 
                           prepareEmailOptions(data, properties, false));
      }
    }
  } catch (error) { Logger.log('Error in onEdit: ' + error.message); }
}

function saveSetupData(supervisorEmail, enableSMS, supervisorPhoneEmail, ccSupervisor, ccEmployee, enableThirdPartyEmail, thirdPartyEmail) {
  const errors = [];
  if (!isValidEmail(supervisorEmail)) errors.push('Invalid Supervisor Email.');
  if (enableSMS && !isValidEmail(supervisorPhoneEmail)) errors.push('Invalid Supervisor Phone Email.');
  if (enableThirdPartyEmail && !isValidEmail(thirdPartyEmail)) errors.push('Invalid Third-Party Email.');
  if (errors.length > 0) throw new Error(errors.join('\n'));

  const data = {supervisorEmail, enableSMS: enableSMS.toString(), supervisorPhoneEmail, ccSupervisor: ccSupervisor.toString(), 
                ccEmployee: ccEmployee.toString(), enableThirdPartyEmail: enableThirdPartyEmail.toString(), thirdPartyEmail};
  CacheService.getUserCache().put('setupData', JSON.stringify(data), 600);

  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'SUPERVISOR_EMAIL': supervisorEmail,
    'ENABLE_SMS': enableSMS.toString(),
    'SUPERVISOR_PHONE_EMAIL': enableSMS ? supervisorPhoneEmail : '',
    'CC_SUPERVISOR': ccSupervisor.toString(),
    'CC_EMPLOYEE': ccEmployee.toString(),
    'ENABLE_THIRD_PARTY_EMAIL': enableThirdPartyEmail.toString(),
    'THIRD_PARTY_EMAIL': enableThirdPartyEmail ? thirdPartyEmail : '',
  });

  if (properties.getProperty('SETUP_COMPLETED') !== 'true') {
    addApprovalColumn();
    createTriggers();
    properties.setProperty('SETUP_COMPLETED', 'true');
  }
}

function generateEmailHTML(data, isRequest, shortUrl = '') {
  const { employeeEmail, requestedItem, estimatedCost, itemLink, purchaseReason, approvalStatus, imageTag } = data;
  const title = isRequest ? 'New Purchase Request' : `Your Purchase Request Has Been ${approvalStatus}`;
  const footer = isRequest ? `<a href="${shortUrl}" style="color:#1a73e8;">Click here to open the Purchase Request spreadsheet.</a>` : '';
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
            ${isRequest ? `<tr><th>Requestor:</th><td>${employeeEmail}</td></tr>` : ''}
            <tr><th>Item:</th><td>${requestedItem}</td></tr>
            <tr><th>Estimated Cost:</th><td>$${estimatedCost}</td></tr>
            <tr><th>Item Link:</th><td><a href="${itemLink}" style="color:#1a73e8;">${itemLink}</a></td></tr>
            <tr><th>Reason for Purchase:</th><td>${purchaseReason}</td></tr>
            <tr><th>Status:</th><td><em>${isRequest ? 'Pending' : approvalStatus}</em></td></tr>
            ${imageTag || ''}
          </table>
        </div>
        <div class="email-footer">${footer}</div>
      </div>
    </body>
    </html>
  `;
}

function authorizationTest() {
  try {
    GmailApp.getInboxUnreadCount();
    DriveApp.getFiles();
    Drive.Files.list({pageSize: 1});
  } catch (e) {
    Logger.log('Authorization error: ' + e.toString());
    throw e;
  }
}

function sendSmsNotification(requestedItem, shortUrl, phoneEmail) {
  const smsMessage = `[[FORM]] Purchase Request - ${requestedItem.substring(0, 157 - shortUrl.length)}... ${shortUrl}`;
  GmailApp.sendEmail(phoneEmail, "", smsMessage);
}

function prepareEmailOptions(data, properties, isNewRequest, shortUrl = '') {
  const imageTag = generateImageTag(data.photoUpload);
  const htmlMessage = generateEmailHTML({ ...data, imageTag }, isNewRequest, shortUrl);
  const emailOptions = {
    to: isNewRequest ? properties.SUPERVISOR_EMAIL : data.employeeEmail,
    subject: `[[Form]] ${data.requestedItem} - ${isNewRequest ? 'New Request' : data.approvalStatus}`,
    htmlBody: htmlMessage,
    attachments: []
  };
  const ccList = [];
  if (isNewRequest ? properties.CC_EMPLOYEE === 'true' : properties.CC_SUPERVISOR === 'true') {
    ccList.push(isNewRequest ? data.employeeEmail : properties.SUPERVISOR_EMAIL);
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
  return ['EMAIL', 'ITEM', 'COST', 'LINK', 'PHOTO', 'REASON'].reduce((acc, key) => {
    acc[key.toLowerCase()] = sheet.getRange(row, COLUMNS[key]).getValue();
    return acc;
  }, {});
}

const isValidEmail = email => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);

function shortenUrl(longUrl) {
  try {
    const response = UrlFetchApp.fetch('https://ulvis.net/API/write/get?url=' + encodeURIComponent(longUrl) + '&type=json');
    const json = JSON.parse(response.getContentText());
    return (json.success === 1 && json.data && json.data.url) ? json.data.url : longUrl;
  } catch (error) {
    Logger.log('Error shortening URL: ' + error.message);
    return longUrl;
  }
}

function generateImageTag(photoUpload) {
  if (!photoUpload) return '';
  try {
    const fileId = photoUpload.match(/[-\w]{25,}/)[0];
    if (!fileId) throw new Error('Invalid file ID');
    DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileMetadata = Drive.Files.get(fileId, { fields: 'thumbnailLink, webContentLink' });
    const imageUrl = fileMetadata.thumbnailLink ? fileMetadata.thumbnailLink.replace(/=s220$/, '=s0') : fileMetadata.webContentLink;
    if (!imageUrl) throw new Error('No valid image URL found.');
    return `<tr><td colspan="2" style="text-align:center;"><img src="${imageUrl}" alt="Uploaded image" style="max-width:100%;height:auto;border:1px solid #ddd;border-radius:4px;"></td></tr>`;
  } catch (error) {
    Logger.log('Error generating image tag: ' + error.message);
    return `<tr><td colspan="2" style="text-align:center;"><p>Error loading image. Please check the original file in Google Drive.</p></td></tr>`;
  }
}

function getSetupData() {
  try {
    return PropertiesService.getScriptProperties().getProperties();
  } catch (e) {
    Logger.log('Authorization required or properties could not be accessed.');
    return {};
  }
}
