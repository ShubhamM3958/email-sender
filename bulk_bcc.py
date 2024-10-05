import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from dotenv import load_dotenv


# File paths
input_file = 'CCC-Bulk-LATE.xlsx'
remaining_file = 'temp/Remaining-CCC-Bulk-LATE.xlsx'
invited_file = 'Invited-CCC-Bulk-LATE.xlsx'
html_template_file = 'email_template.html'

# Email settings
email_sender = 'gju.tpoffice@gmail.com'
password = '<tpoffice password>'  # contact me for this
smtp_server = 'smtp.gmail.com'
smtp_port = 587

load_dotenv()  # load env password

to_addresses = ['tpcell@gjust.org']
custom_bcc = ['tpcell@gjust.org', 'shubham.polar@gmail.com']

# Batch size of emails to send
batch_size = 98
# Load the remaining emails from 'Remaining_HR_List.xlsx' if it exists, otherwise load from the original file
if os.path.exists(remaining_file):
    df = pd.read_excel(remaining_file)
    if df.empty:
        print(f'{remaining_file} is empty. Loading original file.')
        df = pd.read_excel(input_file)
    else:
        print(f'Loaded remaining emails from {remaining_file}')
else:
    df = pd.read_excel(input_file)
    print(f'Loaded original emails from {input_file}')

# Filter valid emails (using regex to check email format)
email_regex = r'^[^\s@]+@[^\s@]+\.[^\s@]+$'

df['valid_email'] = df['HR Mail-ID'].str.contains(email_regex, na=False)
to_email_df = pd.DataFrame({'HR Mail-ID': custom_bcc, 'valid_email': [True] * len(custom_bcc)})
valid_emails_df = df[df['valid_email']]  # Dataframe with valid emails only

# Get the first batch of 100 emails
daily_emails_df = valid_emails_df[:batch_size]
daily_emails_df = pd.concat([daily_emails_df, to_email_df], ignore_index=True)


# Function to load HTML email template
def load_html_template(template_file):
    with open(template_file, 'r', encoding='utf-8') as file:
        html_content = file.read()
    return html_content


# Function to send emails
def send_emails(email_list, html_content):
    # Set up the SMTP server
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email_sender, password)

    # Email content
    subject = ("Invitation for Campus Placement Drives for YOP 2025 students | Guru Jambheswhar University ("
               "Hisar)")
    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = ', '.join(to_addresses)
    # msg['To'] = to_addresses
    msg['Subject'] = subject
    msg.attach(MIMEText(html_content, 'html'))

    try:
        # Send email with multiple BCC recipients
        server.sendmail(email_sender, email_list, msg.as_string())
        # print(email_list)
        print(f'Email sent to {len(email_list)} recipients in BCC.')
        server.quit()
        return email_list
    except Exception as e:
        print(f'Failed to send email. Error: {str(e)}')
        server.quit()
        return []


html_template = load_html_template(html_template_file)

# Send emails to the first 100 valid email addresses
sent_emails = send_emails(daily_emails_df['HR Mail-ID'], html_template)

# Update the DataFrame: mark "Invited: Yes" for sent emails
df.loc[df['HR Mail-ID'].isin(sent_emails), 'Invited'] = 'Yes'

# Save the updated invited list
if os.path.exists(invited_file):
    invited_df = pd.read_excel(invited_file)
    invited_df = pd.concat([invited_df, df[df['HR Mail-ID'].isin(sent_emails)]], ignore_index=True)
else:
    invited_df = df[df['HR Mail-ID'].isin(sent_emails)]

invited_df.to_excel(invited_file, index=False)
print(f'Updated invited list saved to {invited_file}')

# Save the remaining emails for the next batch
remaining_df = valid_emails_df[batch_size:].copy()
remaining_df.drop(columns=['valid_email'], inplace=True)  # Remove temp column

if remaining_df.empty:
    if os.path.exists(remaining_file):
        os.remove(remaining_file)
        print(f'No remaining emails. Deleted {remaining_file}.')
    else:
        print('No remaining emails and the file does not exist.')
else:
    remaining_df.to_excel(remaining_file, index=False)
    print(f'Remaining email list saved to {remaining_file}')
