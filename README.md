# Email Sender

This is an automated email sender that sends bulk emails using Python. It supports validating email addresses and sending custom emails to multiple recipients in one go, with support for templates and BCC.

## Features

- **Bulk Email Sending**: Send emails to multiple recipients using BCC to avoid exposing recipients' email addresses.
- **Email Validation**: Ensure that the email addresses are valid before sending.
- **Custom Email Templates**: Send personalized emails using HTML templates.
- **Spreadsheet Support**: Load recipient data from `.xlsx` files.
- **Environment Variable Support**: Store sensitive credentials securely in a `.env` file.

## Folder Structure
  email-sender 
    ├── vaildate_mail # Folder for email validation scripts 
    ├── CCC-Bulk-LATE.xlsx # Sample Excel file with recipient data 
    ├── bulk_bcc.py # Main script for sending bulk emails 
    ├── mailer.py # Individual email sending
    ├── email_template.html # HTML email template for personalized emails 
    ├── .gitignore # Gitignore file to exclude unnecessary files (e.g., .env) 
    ├── requirements.txt # Required Python packages 
    └── README.md # Documentation for the project
