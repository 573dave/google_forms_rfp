# Google Forms Purchase Request System

This project provides a Google Forms-based system for employees to request purchase approvals. The system is linked to a Google Sheet, where a script automates notifications to the supervisor for each new request and allows them to approve or deny requests. Once a request is approved, the employee is automatically notified via email.

## Features

- **Google Form Integration**: Employees can easily request purchases using a Google Form, including an optional photo upload.
- **Automated Notifications**: Supervisor receives an email with a summary of the request and a link to the Google Sheet for approval.
- **Approval Process**: Once the supervisor approves or denies the request in the sheet, employees are notified automatically.
- **HTML Email Summaries**: Requests and approvals are sent as HTML emails for easy reading.
- **Configurable Notification Preferences**: Options to enable CC for supervisors or employees, as well as notifications for third-party stakeholders (e.g., managers or secretaries).
- **SMS Notification (Optional)**: Supervisors can also receive SMS notifications if their phone email-to-text gateway is provided.

## Setup Instructions

### 1. Create the Google Form
1. Go to [Google Forms](https://forms.google.com) and create a new form.
2. Add the following fields:
   - **Item to Purchase** (Short answer)
   - **Estimated Cost** (Short answer with number validation)
   - **Link to Item** (Short answer for a URL)
   - **Optional Photo Upload** (File upload field, requires sign-in)
   - **Reason for Purchase** (Paragraph for detailed explanation)

**Note:** The current code will only work if all of these fields are present and in the specified order, with no additional fields. If you need to add or modify fields, you will need to adjust the code accordingly.
   - **Item to Purchase** (Short answer)
   - **Estimated Cost** (Short answer with number validation)
   - **Link to Item** (Short answer for a URL)
   - **Optional Photo Upload** (File upload field)
   - **Reason for Purchase** (Paragraph for detailed explanation)

### 2. Link to Google Sheets
1. Click on the **Responses** tab in the form and create a linked Google Sheet.

### 3. Add the Scripts
1. In the Google Sheet, go to **Extensions > Apps Script**.
2. In the Apps Script editor, copy and paste the contents of [Code.gs](https://raw.githubusercontent.com/573dave/google_forms_rfp/refs/heads/main/Code.gs){:target="_blank"} from this repository, replacing the empty "myFunction".
3. Create another new file for the HTML form:
   - Click on the **+** icon again and select **HTML**.
   - Name it `SetupForm` and copy and paste the contents of [SetupForm.html](https://raw.githubusercontent.com/573dave/google_forms_rfp/refs/heads/main/SetupForm.html){:target="_blank"} from this repository into the editor.
4. Save both files and close the editor.
5. Reload the Google Sheet to ensure the script updates are recognized.

### 4. Run the Setup Script
1. In the linked Google Sheet, go to the newly added **Setup Menu** in the toolbar.
2. Select **Run Setup**. This will open a sidebar in the Google Sheet.
3. Fill in the required supervisor email address, configure CC preferences, enable SMS notifications if needed, and optionally add a third-party email for additional notification.

### 5. Test the System
1. Submit a test request via the Google Form.
2. Verify that the supervisor receives a notification with a summary and link to the sheet.
3. Approve or deny the request in the **Approval Status** column of the sheet and confirm that the employee is notified via email.

## Script Overview

### `setup()`
Runs the setup form in the sidebar, allowing configuration of email settings for the supervisor, SMS notifications, CC preferences, and optional third-party notifications. Also handles the creation of necessary triggers automatically.

### `onFormSubmit(e)`
Triggered when a form submission is received. Sends a notification to the supervisor with the details of the request.

### `sendApprovalNotification()`
Checks the approval status of the most recent request and notifies the employee if the request has been approved.

### `saveSetupData()`
Handles storing configuration options set by the setup form, including supervisor email, SMS notification setup, CC preferences, and third-party email settings.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details. In summary, you are free to use, copy, modify, merge, publish, distribute, sublicense, and sell copies of the software, provided you include the original copyright notice. The software is provided "as is", without warranty of any kind, and the creators are not liable for any issues that arise from using the software.
