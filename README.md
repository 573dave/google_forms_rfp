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
1. The first time you open the sheet or make an edit, you'll be prompted to grant the necessary permissions for the script to function.
2. Follow the prompts to review and grant the required permissions.


## Script Overview

### `onOpen()`
Sets up the custom menu and attempts to obtain necessary permissions when the sheet is opened.

### `setup()`
Runs the setup form in the sidebar, allowing configuration of email settings for the supervisor, SMS notifications, CC preferences, and optional third-party notifications.

### `onFormSubmit(e)`
Triggered when a form submission is received. Sends a notification to the supervisor with the details of the request.

### `onEdit(e)`
Triggered when an edit is made to the sheet. Checks if the approval status has changed and sends notifications accordingly.

### `sendRequestNotification(e)`
Handles sending notifications for new purchase requests, including email and optional SMS.

### `generateEmailHTML()`
Generates the HTML content for email notifications, including embedded images when available.

### `saveSetupData()`
Handles storing configuration options set by the setup form, including supervisor email, SMS notification setup, CC preferences, and third-party email settings.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details. In summary, you are free to use, copy, modify, merge, publish, distribute, sublicense, and sell copies of the software, provided you include the original copyright notice. The software is provided "as is", without warranty of any kind, and the creators are not liable for any issues that arise from using the software.
