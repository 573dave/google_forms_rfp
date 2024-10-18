# Google Forms Purchase Request System

This project provides a Google Forms-based system for employees to request purchase approvals. The system is linked to a Google Sheet, where a script automates notifications to the supervisor for each new request and allows them to approve or deny requests. Once a request is approved or denied, the employee is automatically notified via email.

## Features

- **Google Form Integration**: Employees can easily request purchases using a Google Form, including an optional photo upload.
- **Automated Notifications**: Supervisor receives an email with a summary of the request and a link to the Google Sheet for approval.
- **Approval Process**: Once the supervisor approves or denies the request in the sheet, employees are notified automatically.
- **HTML Email Summaries**: Requests and approvals are sent as HTML emails for easy reading, including embedded images when available.
- **Configurable Notification Preferences**: Options to enable CC for supervisors or employees, as well as notifications for third-party stakeholders (e.g., managers or secretaries).
- **SMS Notification (Optional)**: Supervisors can also receive SMS notifications if their phone email-to-text gateway is provided.
- **Automatic Permission Handling**: The script automatically requests necessary permissions when the sheet is opened or edited.
- **Custom Menu**: A custom menu is added to the Google Sheet for easy access to setup and authorization functions.

## Setup Instructions

### 1. Create the Google Form
1. Go to [Google Forms](https://forms.google.com) and create a new form.
2. Add the following fields:
   - **Item to Purchase** (Short answer)
   - **Estimated Cost** (Short answer with number validation)
   - **Link to Item** (Short answer for a URL)
   - **Optional Photo Upload** (File upload field, requires sign-in)
   - **Reason for Purchase** (Paragraph for detailed explanation)

**Note:** The current code expects these fields to be present and in this specific order. If you need to add or modify fields, you will need to adjust the `COLUMNS` constant in the script accordingly.

### 2. Link to Google Sheets
1. Click on the **Responses** tab in the form and create a linked Google Sheet.

### 3. Add the Scripts
1. In the Google Sheet, go to **Extensions > Apps Script**.
2. In the Apps Script editor, delete any existing code in the default `Code.gs` file.
3. Copy the entire content of the [Code.gs](https://raw.githubusercontent.com/573dave/google_forms_rfp/main/Code.gs) file from the GitHub repository.
4. Paste the copied code into the `Code.gs` file in the Apps Script editor.
5. Create a new HTML file for the setup form:
   - Click on the **+** icon next to Files in the left sidebar.
   - Select **HTML** from the dropdown menu.
   - Name the new file `SetupForm`.
6. Copy the entire content of the [SetupForm.html](https://raw.githubusercontent.com/573dave/google_forms_rfp/main/SetupForm.html) file from the GitHub repository.
7. Paste the copied HTML into the newly created `SetupForm.html` file in the Apps Script editor.
8. Save both files (Code.gs and SetupForm.html).
9. Close the Apps Script editor.
10. Reload the Google Sheet to ensure the script updates are recognized.

### 4. Run the Setup Script
1. In the linked Google Sheet, you should see a new menu item called `[[Menu]]` in the toolbar.
2. Click on `[[Menu]]` and select **Settings**. This will open a sidebar in the Google Sheet.
3. Fill in the required supervisor email address, configure CC preferences, enable SMS notifications if needed, and optionally add a third-party email for additional notification.
4. Click "Save" to store your settings.

### 5. Grant Permissions
1. Click on `[[Menu]]` and select **Authorize** to grant the necessary permissions for the script to function.
2. Follow the prompts to review and grant the required permissions.

## System Flow Chart

Below is an ASCII flow chart representing the main processes of the Purchase Request System:

```
     +--------------+ +----------------+ +------------+
     |    Setup     | | Form Response  | |  OnEdit    |
     +--------------+ +----------------+ +------------+
            |                |                |
            v                v                v
 +--------------------+ +--------------+ +------------------+
 | Load SetupForm     | | Get Form Data| | Check Edited Cell|
 +--------------------+ +--------------+ +------------------+
            |                |                |
            v                v                v
 +--------------------+ +--------------+ +------------------+
 | Display Sidebar    | | Set 'Pending'| | Is Approval      |
 +--------------------+ +--------------+ | Status?          |
            |                |           +------------------+
            v                |                |
 +--------------------+      |           No   |   Yes
 | User Inputs        |      |        +-------+-------+
 | Settings           |      |        |               |
 +--------------------+      |        v               v
            |                |    (End)     +------------------+
            v                |               | Get Updated     |
 +--------------------+      |               | Status          |
 | Save Setup Data    |      |               +------------------+
 +--------------------+      |                       |
            |                |                       v
            v                v               +------------------+
 +--------------------+ +--------------+     | Prepare Email    |
 | Add Approval       | | Prepare Email|     | Content          |
 | Column             | | Content      |     +------------------+
 +--------------------+ +--------------+             |
            |                |                       v
            v                v               +------------------+
 +--------------------+ +--------------+     | Send Notification|
 | Create Triggers    | | Send to      |     | to Employee      |
 +--------------------+ | Supervisor   |     +------------------+
            |           +--------------+             |
            |                |                       v
            |                v               +------------------+
            |        +--------------+        | CC Options?      |
            |        | SMS Enabled? |        +------------------+
            |        +--------------+         No |          | Yes
            |         No |       | Yes          |          |
            |            |       |              |          v
            |            |       v              |  +------------------+
            |            |   +----------+       |  | Send CC Emails   |
            |            |   | Send SMS |       |  +------------------+
            |            |   +----------+       |          |
            |            |       |              |          |
            v            v       v              v          v
                       +----------+
                       |   End    |
                       +----------+
```

## Script Overview

### `onOpen()`
Sets up the custom menu when the sheet is opened.

### `setup()`
Runs the setup form in the sidebar, allowing configuration of email settings for the supervisor, SMS notifications, CC preferences, and optional third-party notifications.

### `onSetupFormLoaded()`
Performs authorization tests and creates necessary triggers when the setup form is loaded.

### `authorizationTest()`
Tests various Google services to ensure proper authorization.

### `authorizeScript()`
Manually triggers the authorization process.

### `addApprovalColumn()`
Adds an "Approval Status" column to the sheet if it doesn't exist.

### `createTriggers()`
Creates necessary triggers for form submission and edit events.

### `onFormSubmit(e)`
Triggered when a form submission is received. Sends a notification to the supervisor with the details of the request.

### `onEditHandler(e)`
Triggered when an edit is made to the sheet. Checks if the approval status has changed and sends notifications accordingly.

### `sendNotification(data, isNewRequest, sheetUrl)`
Handles sending notifications for new purchase requests or status updates, including email and optional SMS.

### `generateEmailHTML(data, isRequest, sheetUrl, imageTag)`
Generates the HTML content for email notifications, including embedded images when available.

### `sendSmsNotification(requestedItem, sheetUrl, phoneEmail)`
Sends SMS notifications for new purchase requests.

### `prepareEmailOptions(data, properties, isNewRequest, sheetUrl)`
Prepares email options including recipients, subject, and content.

### `getRowData(sheet, row)`
Retrieves data from a specific row in the sheet.

### `generateImageTag(photoUpload)`
Generates an HTML image tag or file link for the uploaded photo.

### `saveSetupData(formData)`
Handles storing configuration options set by the setup form.

### `getSetupData()`
Retrieves the current setup data from script properties.

### `validateSetupData(data)`
Validates the setup data before saving.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details. In summary, you are free to use, copy, modify, merge, publish, distribute, sublicense, and sell copies of the software, provided you include the original copyright notice. The software is provided "as is", without warranty of any kind, and the creators are not liable for any issues that arise from using the software.
