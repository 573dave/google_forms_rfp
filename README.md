# Google Forms Purchase Request System

This project provides a Google Forms-based system for employees to request purchase approvals. The system is linked to a Google Sheet, where a script automates notifications to the boss for each new request and allows them to approve or deny requests. Once a request is approved, the employee is automatically notified via email.

## Features

- **Google Form Integration**: Employees can easily request purchases using a Google Form.
- **Automated Notifications**: Boss receives an email with a summary of the request and a link to the Google Sheet for approval.
- **Approval Process**: Once the boss approves or denies the request in the sheet, employees are notified automatically.
- **HTML Email Summaries**: Requests and approvals are sent as HTML emails for easy reading.

## Setup Instructions

### 1. Create the Google Form
1. Go to [Google Forms](https://forms.google.com) and create a new form.
2. Add the following fields:
   - **Email Address** (Email question type)
   - **Item to Purchase** (Short answer)
   - **Reason for Purchase** (Paragraph)
   - **Estimated Cost** (Short answer, number validation)
   - **Approval Status** (Multiple choice with options: "Pending", "Approved", "Denied")

### 2. Link to Google Sheets
1. Click on the **Responses** tab in the form, and click the green Sheets icon to create a linked response sheet.

### 3. Add the Script
1. In the Google Sheet, go to **Extensions > Apps Script**.
2. Copy and paste the script from this repository (`purchase_request_script.js`) into the Apps Script editor.
3. Save and close the editor.

### 4. Set Up Triggers
1. In the Apps Script editor, click on the **Triggers** icon (clock icon).
2. Set up a trigger for `onFormSubmit` to fire **On form submit**.
3. Set up a time-driven trigger for `sendApprovalNotification` to run at regular intervals.

### 5. Test the System
1. Submit a test request via the Google Form.
2. Verify that the boss receives a notification with a summary and link to the sheet.
3. Approve or deny the request in the **Approval Status** column of the sheet and confirm that the employee is notified via email.

## Script Overview

### `onFormSubmit(e)`
Triggered when a form submission is received. Sends a notification to the boss with the details of the request.

### `sendApprovalNotification()`
Checks the approval status of the most recent request and notifies the employee if the request has been approved.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
