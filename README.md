# Email Sending API From Excel File

This is a **Spring Boot** application to send emails to a list of recipients from an Excel file and update the Excel file with email status and timestamp.

---

## Features

- Read recipient emails from an Excel file.
- Send emails with attachments.
- Async email sending using `ThreadPoolTaskExecutor`.
- Update Excel file with **Mail Status** and **Mail Sent Time**.
- Auto-creates missing columns if needed.
- Handles Excel file safely with copy+replace (prevents issues if the file is open).

---

## Request Example

You can send a POST request to:


### JSON Body Example

{
  "subject": "Application for Full Stack Developer Position – Pranav Wandre",
  "body": "Hi,\nI’m Pranav, a full stack developer with almost 3 years of experience. I’m exploring new opportunities where I can contribute and grow.\nAttaching my resume for your review.\nThanks for your time!\n\nThanks & Regards,\nPranav Wandre\n📞 +91-910*******",
  "fileName": "D:/Email_Send_Resume/Pranav_Wandre_Resume_S.pdf",
  "excelFilePath": "D:/Email_Send_Resume/Hiring Recruiter Email id's.xlsx",
  "sheetName": "Java_Only Valid&Updated",
  "emailColumn": "HR Email id"
}

