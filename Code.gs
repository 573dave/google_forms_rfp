// Global constants
const APPROVAL_STATUSES = ["Pending", "Approved", "Denied"];
const APPROVAL_COLUMN_INDEX = 1;
const EMPLOYEE_EMAIL_COLUMN = 3;
const REQUESTED_ITEM_COLUMN = 4;
const ESTIMATED_COST_COLUMN = 5;
const ITEM_LINK_COLUMN = 6;
const PHOTO_UPLOAD_COLUMN = 7;
const PURCHASE_REASON_COLUMN = 8;

// Show the custom menu on the Google Sheets toolbar when the sheet is opened
function onOpen() {
  try {
    authorizationTest();
    SpreadsheetApp.getUi()
      .createMenu('[[Menu]]')
      .addItem('Settings', 'setup')
      .addToUi();
  } catch (e) {
    Logger.log('Error in onOpen: ' + e.toString());
    // We can't show UI alerts in onOpen, so we'll just log the error
  }
}

// Function to be run initially, which only opens the setup form in a sidebar.
function setup() {
  const data = getSetupData();
  const template = HtmlService.createTemplateFromFile('SetupForm');
  template.data = data || {};
  const htmlOutput = template.evaluate().setTitle('Setup');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// Function to add approval column if it doesn't already exist, with data validation.
function addApprovalColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (sheet.getRange(1, 1).getValue() !== "Approval Status") {
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1).setValue("Approval Status");
    
    const lastRow = sheet.getLastRow();
    const approvalColumn = sheet.getRange(2, 1, lastRow - 1, 1);
    
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(APPROVAL_STATUSES, true)
      .setAllowInvalid(false)
      .build();
    approvalColumn.setDataValidation(rule);
    
    if (lastRow > 1) {
      approvalColumn.setValue("Pending");
    }
  }
}

// Function to create triggers, avoiding duplication.
function createTriggers() {
  const triggers = {
    onFormSubmit: ScriptApp.EventType.ON_FORM_SUBMIT,
    onEdit: ScriptApp.EventType.ON_EDIT
  };

  Object.entries(triggers).forEach(([funcName, eventType]) => {
    if (!ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === funcName)) {
      ScriptApp.newTrigger(funcName)
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .on(eventType)
        .create();
    }
  });
}

// Function to notify the supervisor of a new request
function sendRequestNotification(e) {
    try {
      const sheet = e.range.getSheet();
      const row = e.range.getRow();
      const data = getRowData(sheet, row);
      const shortUrl = shortenUrl(SpreadsheetApp.getActiveSpreadsheet().getUrl());
      
      const properties = PropertiesService.getScriptProperties().getProperties();
      const emailOptions = prepareEmailOptions(data, properties, true, shortUrl);
      
      GmailApp.sendEmail(emailOptions.to, emailOptions.subject, '', emailOptions);
      
      if (properties.ENABLE_SMS === 'true' && properties.SUPERVISOR_PHONE_EMAIL) {
        sendSmsNotification(data.requestedItem, shortUrl, properties.SUPERVISOR_PHONE_EMAIL);
      }
    } catch (error) {
      Logger.log('Error sending request notification: ' + error.message);
    }
}

// Trigger: Send email on form submission
function onFormSubmit(e) {
  try {
    authorizationTest();
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    sheet.getRange(row, APPROVAL_COLUMN_INDEX).setValue("Pending");
    sendRequestNotification(e);
  } catch (error) {
    Logger.log('Error in onFormSubmit: ' + error.message);
  }
}

// Function to send email notification when approval status changes
function onEdit(e) {
  try {
    authorizationTest();
    const sheet = e.range.getSheet();
    if (sheet.getRange(1, 1).getValue() !== 'Approval Status') return;
    
    const row = e.range.getRow();
    const column = e.range.getColumn();
    
    if (column === APPROVAL_COLUMN_INDEX && row > 1) {
      const approvalStatus = e.range.getValue();
      if (APPROVAL_STATUSES.includes(approvalStatus)) {
        const data = getRowData(sheet, row);
        data.approvalStatus = approvalStatus;
        
        const properties = PropertiesService.getScriptProperties().getProperties();
        const emailOptions = prepareEmailOptions(data, properties, false);
        
        GmailApp.sendEmail(emailOptions.to, emailOptions.subject, '', emailOptions);
      }
    }
  } catch (error) {
    Logger.log('Error in onEdit: ' + error.message);
  }
}

// Function to save setup data and perform setup actions.
function saveSetupData(
  supervisorEmail,
  enableSMS,
  supervisorPhoneEmail,
  ccSupervisor,
  ccEmployee,
  enableThirdPartyEmail,
  thirdPartyEmail
  ) {
  const properties = PropertiesService.getScriptProperties();

  // Server-side validation
  const errors = [];

  // Validate supervisorEmail
  if (!supervisorEmail || !isValidEmail(supervisorEmail)) {
    errors.push('Invalid Supervisor Email.');
  }

  // Validate supervisorPhoneEmail if SMS notifications are enabled
  if (enableSMS === true || enableSMS === 'true') {
    if (!supervisorPhoneEmail || !isValidEmail(supervisorPhoneEmail)) {
      errors.push('Invalid Supervisor Phone Email.');
    }
  }

  // Validate thirdPartyEmail if third-party notifications are enabled
  if (enableThirdPartyEmail === true || enableThirdPartyEmail === 'true') {
    if (!thirdPartyEmail || !isValidEmail(thirdPartyEmail)) {
      errors.push('Invalid Third-Party Email.');
    }
  }

  if (errors.length > 0) {
    throw new Error(errors.join('\n'));
  }

  // Update cache service first
  const cache = CacheService.getUserCache();
  const data = {
    supervisorEmail,
    enableSMS: enableSMS.toString(),
    supervisorPhoneEmail,
    ccSupervisor: ccSupervisor.toString(),
    ccEmployee: ccEmployee.toString(),
    enableThirdPartyEmail: enableThirdPartyEmail.toString(),
    thirdPartyEmail,
  };
  cache.put('setupData', JSON.stringify(data), 600); // Cache for 10 minutes

  // Store the properties
  properties.setProperties({
    'SUPERVISOR_EMAIL': supervisorEmail,
    'ENABLE_SMS': enableSMS.toString(),
    'SUPERVISOR_PHONE_EMAIL': (enableSMS === true || enableSMS === 'true') ? supervisorPhoneEmail : '',
    'CC_SUPERVISOR': ccSupervisor.toString(),
    'CC_EMPLOYEE': ccEmployee.toString(),
    'ENABLE_THIRD_PARTY_EMAIL': enableThirdPartyEmail.toString(),
    'THIRD_PARTY_EMAIL': (enableThirdPartyEmail === true || enableThirdPartyEmail === 'true') ? thirdPartyEmail : '',
  });

  // Check if initial setup has been completed
  const setupCompleted = properties.getProperty('SETUP_COMPLETED');

  if (!setupCompleted || setupCompleted !== 'true') {
    // Run the setup steps once
    addApprovalColumn();
    createTriggers();

    // Set the setup completed flag
    properties.setProperty('SETUP_COMPLETED', 'true');
  }
}

/**
 * Generates the complete HTML for email templates
 * @param {Object} data - The purchase request data
 * @param {boolean} isRequest - Whether this is a new request (true) or a status update (false)
 * @param {string} shortUrl - The shortened URL for the spreadsheet (only needed for new requests)
 * @return {string} The generated HTML string
 */
function generateEmailHTML(data, isRequest, shortUrl = '') {
  const {
    employeeEmail,
    requestedItem,
    estimatedCost,
    itemLink,
    purchaseReason,
    approvalStatus,
    imageTag
  } = data;

  // Set title and footer based on whether it's a new request or status update
  let title, footer;
  if (isRequest) {
    title = 'New Purchase Request';
    footer = `<a href="${shortUrl}" style="color:#1a73e8;">Click here to open the Purchase Request spreadsheet.</a>`;
  } else {
    title = `Your Purchase Request Has Been ${approvalStatus}`;
    footer = ''; 
  }

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        @import url('https://fonts.googleapis.com/css?family=Ubuntu');
        body {
          font-family: 'Ubuntu', Arial, sans-serif;
          margin: 0;
          padding: 0;
          background-color: #f5f5f5;
        }
        .email-container {
          width: 100%;
          max-width: 600px;
          margin: auto;
          background-color: #ffffff;
          border: 1px solid #ddd;
          border-radius: 4px;
          overflow: hidden;
        }
        .email-header {
          background-color: #333333;
          color: #ffffff;
          padding: 20px;
          text-align: center;
        }
        .email-content {
          padding: 20px;
        }
        .email-content table {
          width: 100%;
          border-collapse: collapse;
        }
        .email-content th, .email-content td {
          padding: 10px;
          text-align: left;
          vertical-align: top;
        }
        .email-content th {
          background-color: #f0f0f0;
          border-bottom: 1px solid #ddd;
        }
        .email-content td {
          border-bottom: 1px solid #eee;
        }
        .email-footer {
          padding: 10px 20px;
          background-color: #f0f0f0;
          text-align: center;
          font-size: 12px;
          color: #888888;
        }
        @media only screen and (max-width: 600px) {
          .email-content table, .email-content th, .email-content td {
            width: 100%;
            display: block;
          }
          .email-content th {
            text-align: left;
          }
        }
      </style>
    </head>
    <body>
      <div class="email-container">
        <div class="email-header">
          <h2>${title}</h2>
        </div>
        <div class="email-content">
          <table>
            ${isRequest ? `
              <tr>
                <th>Requestor:</th>
                <td>${employeeEmail}</td>
              </tr>
            ` : ''}
            <tr>
              <th>Item:</th>
              <td>${requestedItem}</td>
            </tr>
            <tr>
              <th>Estimated Cost:</th>
              <td>$${estimatedCost}</td>
            </tr>
            <tr>
              <th>Item Link:</th>
              <td><a href="${itemLink}" style="color:#1a73e8;">${itemLink}</a></td>
            </tr>
            <tr>
              <th>Reason for Purchase:</th>
              <td>${purchaseReason}</td>
            </tr>
            <tr>
              <th>Status:</th>
              <td><em>${isRequest ? 'Pending' : approvalStatus}</em></td>
            </tr>
            ${imageTag || ''}
          </table>
        </div>
        <div class="email-footer">
          ${footer}
        </div>
      </div>
    </body>
    </html>
  `;
}

// Authorization test function
function authorizationTest() {
  try {
    // Attempt to access services that require authorization
    GmailApp.getInboxUnreadCount();
    DriveApp.getFiles();
    Drive.Files.list({pageSize: 1});
  } catch (e) {
    // If an error occurs, log it and re-throw
    Logger.log('Authorization error: ' + e.toString());
    throw e;
  }
}

// Function to send SMS notification
function sendSmsNotification(requestedItem, shortUrl, phoneEmail) {
  let smsMessage = `[[FORM]] Purchase Request - ${requestedItem}. Details: ${shortUrl}`;
  if (smsMessage.length > 160) {
    const maxItemLength = 157 - shortUrl.length;
    smsMessage = `[[FORM]] Purchase Request - ${requestedItem.substring(0, maxItemLength)}... ${shortUrl}`;
  }
  GmailApp.sendEmail(phoneEmail, "", smsMessage);
}

// Helper function to prepare email options
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
  if (isNewRequest) {
    if (properties.CC_EMPLOYEE === 'true') ccList.push(data.employeeEmail);
  } else {
    if (properties.CC_SUPERVISOR === 'true') ccList.push(properties.SUPERVISOR_EMAIL);
  }
  
  if (properties.ENABLE_THIRD_PARTY_EMAIL === 'true' && properties.THIRD_PARTY_EMAIL) {
    ccList.push(properties.THIRD_PARTY_EMAIL);
  }
  
  if (ccList.length > 0) {
    emailOptions.cc = ccList.join(", ");
  }
  
  return emailOptions;
}

// Helper function to get row data
function getRowData(sheet, row) {
  return {
    employeeEmail: sheet.getRange(row, EMPLOYEE_EMAIL_COLUMN).getValue(),
    requestedItem: sheet.getRange(row, REQUESTED_ITEM_COLUMN).getValue(),
    estimatedCost: sheet.getRange(row, ESTIMATED_COST_COLUMN).getValue(),
    itemLink: sheet.getRange(row, ITEM_LINK_COLUMN).getValue(),
    photoUpload: sheet.getRange(row, PHOTO_UPLOAD_COLUMN).getValue(),
    purchaseReason: sheet.getRange(row, PURCHASE_REASON_COLUMN).getValue()
  };
}

function isValidEmail(email) {
  // Basic email validation regex
  var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

function shortenUrl(longUrl) {
  try {
    const apiUrl = 'https://ulvis.net/API/write/get?url=' + encodeURIComponent(longUrl) + '&type=json';
    const response = UrlFetchApp.fetch(apiUrl);
    const json = JSON.parse(response.getContentText());

    if (json.success === 1 && json.data && json.data.url) {
      return json.data.url;
    } else {
      Logger.log('Error shortening URL: ' + (json.error && json.error.msg ? json.error.msg : 'Unknown error'));
      return longUrl; // Fallback to the original URL
    }
  } catch (error) {
    Logger.log('Error shortening URL: ' + error.message);
    return longUrl; // Fallback to the original URL
  }
}

function generateImageTag(photoUpload) {
    let imageTag = '';
    if (photoUpload) {
      try {
        // Extract the file ID from the photoUpload URL
        const fileId = getFileIdFromUrl(photoUpload);
        if (fileId) {
          // Set the sharing permissions to 'Anyone with the link can view'
          const file = DriveApp.getFileById(fileId);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

          // Use the Advanced Drive Service to get the file metadata
          const fields = 'thumbnailLink, webContentLink';
          const fileMetadata = Drive.Files.get(fileId, { fields: fields });

          // Try to use the thumbnailLink first
          let imageUrl = fileMetadata.thumbnailLink;
          if (imageUrl) {
            // Replace '=s220' with '=s0' to get the full-size image
            imageUrl = imageUrl.replace(/=s220$/, '=s0'); 
          } else if (fileMetadata.webContentLink) {
            // Fallback to webContentLink
            imageUrl = fileMetadata.webContentLink;
          } else {
            throw new Error('No valid image URL found.');
          }

          // Create the image tag
          imageTag = `
            <tr>
              <td colspan="2" style="padding:10px 0; text-align:center;">
                <img src="${imageUrl}" alt="Uploaded image" style="max-width:100%; height:auto; border:1px solid #ddd; border-radius:4px;">
              </td>
            </tr>`;
        }
      } catch (error) {
        Logger.log('Error generating image tag: ' + error.message);
        imageTag = `
          <tr>
            <td colspan="2" style="padding:10px 0; text-align:center;">
              <p>Error loading image. Please check the original file in Google Drive.</p>
            </td>
          </tr>`;
      }
    }
    return imageTag;
}

function getFileIdFromUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function getSetupData() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const data = {
      supervisorEmail: properties.getProperty('SUPERVISOR_EMAIL'),
      enableSMS: properties.getProperty('ENABLE_SMS'),
      supervisorPhoneEmail: properties.getProperty('SUPERVISOR_PHONE_EMAIL'),
      ccSupervisor: properties.getProperty('CC_SUPERVISOR'),
      ccEmployee: properties.getProperty('CC_EMPLOYEE'),
      enableThirdPartyEmail: properties.getProperty('ENABLE_THIRD_PARTY_EMAIL'),
      thirdPartyEmail: properties.getProperty('THIRD_PARTY_EMAIL')
    };
    return data;
  } catch (e) {
    Logger.log('Authorization required or properties could not be accessed.');
    return {};
  }
}
