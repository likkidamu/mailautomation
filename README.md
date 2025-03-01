# Email Automation Script

## Overview
This Python script automates the process of sending job application emails with an attachment to multiple recipients listed in an Excel file.

## Features
- Reads recipient names and email addresses from an Excel file (`Mails.xlsx`).
- Sends personalized emails with a predefined subject and body.
- Attaches a resume (`Damodhar_Resume.docx`) to each email.
- Includes CC recipients.
- Limits the number of emails sent per execution.
- Introduces a delay between emails to prevent spam detection.

## Prerequisites
Ensure you have the following installed before running the script:
- Python 3.x
- Required libraries:
  - `smtplib` (for sending emails)
  - `email` (for handling email content)
  - `os` (for file operations)
  - `pandas` (for reading Excel files)
  - `openpyxl` (for handling Excel files)
  - `time` (for introducing delays)

To install missing dependencies, run:
```sh
pip install pandas openpyxl
```

## Email Configuration
Modify the following variables in the script:
- `smtp_server`: SMTP server address (Gmail used in this case).
- `smtp_port`: SMTP server port (587 for TLS).
- `sender_email`: Your email address.
- `sender_password`: Your app-generated password (Gmail requires an App Password).
- `cc_recipients`: List of CC recipients.
- `attachment_path`: Path to the resume file.

## Excel File Format
The script expects an Excel file (`Mails.xlsx`) with the following columns:
| Name  | Mails           |
|-------|---------------|
| John  | john@email.com |
| Alice | alice@email.com |

## Execution Steps
1. Ensure `Mails.xlsx` and `Damodhar_Resume.docx` are in the script directory.
2. Run the script:
```sh
python script.py
```
3. The script will send emails with a delay of 5 minutes between each batch of 8 emails.

## Notes
- **Security Warning**: Never hardcode passwords in scripts. Instead, use environment variables or secure credential storage.
- **Email Limits**: The script stops sending after reaching the limit (`email_limit` variable, default: 80).
- **Troubleshooting**: Ensure `Less secure apps` access is enabled or use an App Password if using Gmail.

## Disclaimer
Use this script responsibly and comply with email service provider policies to avoid getting flagged as spam.

