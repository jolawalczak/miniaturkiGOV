import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import datetime
import os
import configparser
import sys

# Ensure at least one argument is provided (filename)
if len(sys.argv) < 2:
    print("Usage: python sendmail.py <filename> [optional_body]")
    sys.exit(1)

try:
    path_to_file = sys.argv[1]
except ValueError:
    print("Error: Filename parameter must be a string type.")
    sys.exit(1)

# Extract optional body argument if provided
body = sys.argv[2] if len(sys.argv) > 2 else ""  # Default to empty string if no body provided

# Load configuration from config.ini
config = configparser.ConfigParser()
config.read('configMiniaturki.ini')

# Read sensitive data from the config file
sender_email = config['EMAIL']['sender_email']
sender_password = config['EMAIL']['sender_password']
receiver_email = config['EMAIL']['receiver_email']
SMTP_SERVER = config['EMAIL']['SMTP_SERVER']
SMTP_PORT = int(config['EMAIL']['SMTP_PORT'])  # Convert port to integer

yesterday_date = (datetime.datetime.today() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')

# Set the subject and body of the email
firmyType = path_to_file.split("/")[1].split("_")[0]
raportDate = path_to_file.split("E_")[1].split(".")[0]
subject = f'Raport Business GOV - firmy {firmyType} z dnia {raportDate}'

msg = MIMEMultipart()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = subject

msg.attach(MIMEText(body, 'plain'))

# Check if the file exists before attaching
if os.path.exists(path_to_file):
    with open(path_to_file, 'rb') as file:
        msg.attach(MIMEApplication(file.read(), Name=path_to_file.split("/")[1]))
else:
    print(f"Warning: Attachment {path_to_file} not found. Sending email without it.")

try:
    # Use SMTP_SSL for port 465
    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(sender_email, sender_password)  # No starttls() needed
        text = msg.as_string()
        
        server.sendmail(sender_email, receiver_email, text)
        print("✅ Email sent successfully!")
except Exception as e:
    print(f"❌ Error: unable to send email - {e}")