📧 Automated Email Dispatch System using Excel VBA & Outlook

🚀 Overview

This project is an Excel-based automation solution designed to send bulk emails using Microsoft Outlook. It enables dynamic email generation, attachment handling, file renaming and batch processing, significantly reducing manual effort in repetitive communication workflows.

---

🧩 Problem Statement

Organizations often need to send large volumes of emails (e.g., invoices, reports, notifications). Manual processing is:

- Time-consuming
- Prone to errors
- Difficult to scale

In the current business scinario where this project is being applied, user get hundreds of invoices from finance department after digitally signed in a specific name format. The user then renames the files to fit to their requirement manually and writes seperate email to all customers attaching the files one by one.

---

💡 Solution

Developed a VBA-based automation tool that:

- Reads data from Excel
- Renames the files stared in a specific folder
- Generates personalized emails
- Attaches relevant documents
- Sends or previews emails via Outlook

---

✨ Key Features

- 📨 Bulk email dispatch from Excel
- 🔗 Outlook integration
- 📝 Dynamic subject and email body
- 📎 Automated attachment handling
- 🧾 Template-based email generation
- 👀 Preview mode (Display before sending)
    - Batch mode
    - Selected row mode
- ❌ Error handling:
    - Invalid email IDs
    - Missing attachments
    - Critical Error
- 📂 File name generation
- 🔄 File renaming automation within folders

---

⚙️ Technical Implementation

🔹 VBA Logic

- Outlook object model automation
- Loop-based email generation
- Conditional logic for batch vs selected processing
- Dynamic content insertion
- File handling (rename & attach)
- Error handling using validation checks

---

🔹 Excel Design

- User control sheet:
  - Batch processing
  - Row-level execution
  - Select files folder
  - Navigation Pane
  - Processed Dashboard
- Structured input sheet:
  - Input details of documnet (e.g., Invoice Number | Customer Name | Invoice Type | Invoice Value | etc.,)
  - Other fields populated using excel functions (e.g., Sneder Name | Sender Email ID | Customer Email ID | CC Email | File Names | File Path | etc.,)
- Customer data sheet:
  - Customer name
  - Email ID
  - Location
  - Business Area
- Sent Log sheet:
  - Document No
  - Document Name
  - Sent Date and Time

---

🔄 Workflow

Excel Input → VBA Processing → File Handling → Outlook → Email Sent/Preview

---

🧠 Challenges & Solutions

🔸 Bulk Email Handling

Challenge: Managing large datasets without performance issues
Solution: Optimized looping and minimized unnecessary object calls

---

🔸 Outlook Integration

Challenge: Ensuring stable interaction with Outlook
Solution: Used Outlook object model with controlled instantiation

---

🔸 Dynamic Email Content

Challenge: Personalizing subject and body for each recipient
Solution: Used cell-driven dynamic placeholders

---

🔸 Attachment Handling

Challenge: Missing or incorrect file paths
Solution: Implemented validation checks before sending

---

🔸 Preview vs Send Mode

Challenge: Allowing safe execution without sending emails
Solution: Built dual-mode logic (Display vs Send)

---

🔸 Selected Row Processing

Challenge: Sending email for only specific records
Solution: Added row-based execution logic

---

🔸 File Name Generation

Challenge: Creating dynamic file names based on data
Solution: Used concatenation logic from Excel inputs

---

🔸 File Renaming in Folder

Challenge: Managing file naming consistency
Solution: Implemented file system automation using VBA

---

🔸 Error Handling

Challenge: Avoiding runtime failures
Solution: Added validation and conditional checks

---

📈 Impact

- ⏱️ Reduced manual effort significantly
- 📤 Enabled bulk email automation
- ✅ Improved accuracy and consistency
- 🔁 Reusable communication workflow

---

🛠️ Tools Used

- Microsoft Excel
- VBA (Visual Basic for Applications)
- Microsoft Outlook

---

📂 Project Structure

Mail-Dispatch-Engine/
│
├── Email_Automation_Template.xlsm (This file can be customized to the specific usage)
├── Screenshots/
└── README.md

---

▶️ How to Use

1. Open the Excel template
2. Enter email details in the input sheet
3. Select mode (Batch / Selected row)
4. Click “Send / Preview Emails”
5. Emails are generated via Outlook

---

⚠️ Notes

- Enable macros before using the file
- Ensure Outlook is configured
- Use preview mode before sending

---

🔮 Future Enhancements

- Email tracking (sent status)
- Log sheet for errors
- HTML email templates
- Dashboard for monitoring

---

👨‍💻 Author

Rama Subba Rao Polamraju
