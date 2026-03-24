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
- Renames the files stored in a specific folder
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
  - Process Dashboard
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

Automated-Email-Dispatch-Engine/
│
├── Mail-Dispatch-Engine.xlsm (This file can be customized to the specific usage)
├── Screenshots/
├── SampleDocuments/
└── README.md

---

📸 Screenshots 

🔸 Process Flow

Control Panel
![Control Sheet](ScreenShots/User_Control_Sheet.png)

Maintain Customer Data
![Customer Sheet](ScreenShots/Customer_Data_Sheet.png)

Input Document Details
![Data Input Sheet](ScreenShots/Data_Input_Sheet.png)

File Names Generation
![Generate File Names](ScreenShots/File_Names_Generated_Sheet.png)

Files Before Renaming
![File Before Rename](ScreenShots/Files_Before_Renaming.png)

Files After Renaming
![File After Rename](ScreenShots/Files_After_Renaming.png)

Email Dispatch
![Dispatch Sheet](ScreenShots/Dispatch_Sheet.png)

Email Processing 
![Pracessing Email](ScreenShots/Mail_Generation_Processing.png)

Generated Email

![Generate Email](ScreenShots/Email_Generated.png)

Log of Sent mails
![Log](ScreenShots/Sent_Log.png)

🔸 Code

Choosing Path
![Choose Path](ScreenShots/Code_Choose_Path.png)

FileName Generation
![Gen FileName](ScreenShots/Code_FileName_Generation.png)

Get File Path
![Gile Path](ScreenShots/Code_Get_FilePath.png)

File Renaming
![Rename File](ScreenShots/Code_File_Rename.png)

Get Data into Dispatch Table
![Get Data](ScreenShots/Code_Get_Data.png)

Selecting Mail Template
![Mail Template](ScreenShots/Code_MailTemplate_Selection.png)

Preview or Selected Row
![Preview](ScreenShots/Code_Preview_Logic.png)

Insert Range to Outmail
![Insert Range](ScreenShots/Code_Insert_Range.png)

Insert Multiple Attachments
![Multi Attachment](ScreenShots/Code_Insert_Multi_Attachment.png)

Range Placeholder Replacement
![Copy Range](ScreenShots/Code_Range_Copy.png)

Email Sent Log
![Log Sent](ScreenShots/Code_Log_Sent.png)

Clearing Data
![Clear](ScreenShots/Code_Clear_Data.png)

---

▶️ How to Use

1. Open the Excel template
2. Select mode (Batch / Selected row)
3. Select Folder path
4. Maintain customer data
5. Enter documwent details in the input sheet
6. Click Send Email button
7. Emails are generated via Outlook

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
