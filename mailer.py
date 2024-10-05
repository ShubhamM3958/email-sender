import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import os

# Email settings
smtp_server = "smtp.gmail.com"
smtp_port = 587

# Input files
input_file = 'CCC-Bulk-LATE.xlsx'
remaining_file = 'temp/Remaining-CCC-Bulk-LATE.xlsx'
invited_file = 'Invited-CCC-Bulk-LATE.xlsx'
batch_size = 50


def load_html_template(template_file):
    with open(template_file, 'r', encoding='utf-8') as file:
        html_content = file.read()
    return html_content

# Function to send individual emails
def send_email(to_email, html_content, subject, email_sender, email_password):
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(email_sender, email_password)

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(html_content, 'html'))

    try:
        # Send email
        server.sendmail(email_sender, to_email, msg.as_string())
        print(f'Email sent to {to_email}.')
        server.quit()
        return True
    except Exception as e:
        print(f'Failed to send email to {to_email}. Error: {str(e)}')
        server.quit()
        return False

# Main function to process and send emails
def process_emails(input_file, remaining_file, invited_file, html_template_file, email_sender, email_password, subject):
    # Load the remaining contacts if the file exists, otherwise load the original input file
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

    # Filter valid emails using regex
    email_regex = r'^[^\s@]+@[^\s@]+\.[^\s@]+$'
    df['valid_email'] = df['HR Mail-ID'].str.contains(email_regex, na=False)
    valid_emails_df = df[df['valid_email']]

    # Get the first batch of valid emails
    daily_emails_df = valid_emails_df[:batch_size]

    # Load the HTML template
    html_template = load_html_template(html_template_file)

    # List to store successfully sent emails
    sent_emails = []

    # Send emails one by one
    for index, row in daily_emails_df.iterrows():
        to_email = row['HR Mail-ID']
        hr_person = row['HR Person']

        if pd.isna(hr_person) or hr_person.strip() == "":
            hr_person = "HR"

        personalized_html_content = html_template.replace("{{HR_Person}}", hr_person)

        if send_email(to_email, personalized_html_content, subject, email_sender, email_password):
            sent_emails.append(to_email)

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

# Example usage
if __name__ == "__main__":
    subject = ("Invitation for Campus Placement Drives for YOP 2025 students | Guru Jambheswhar University ("
               "Hisar)")
    # from_email = "shubhammalhotra92890@gmail.com"
    from_email = 'gju.tpoffice@gmail.com'
    # password = "szzr qhgm fslu qpmd"
    password = 'rhvm vsie jgmn kbiv' # tp
    template_file = "email_template.html"

    # process_emails(input_file, remaining_file, invited_file, from_email, password, template_file, subject)
    process_emails(input_file, remaining_file, invited_file, template_file, from_email, password, subject)
    # send_email(subject, to_email, from_email, password, template_file, template_vars)
