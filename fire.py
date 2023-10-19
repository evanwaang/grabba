import csv
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def try_send_email(row, smtp_server):
    email_formats = [
        f"{row['FirstName']}.{row['LastName']}@{row['Company']}.com",
        f"{row['FirstName'][0]}.{row['LastName']}@{row['Company']}.com",
        f"{row['FirstName']}.{row['LastName']}@{row['CompanyAbrv']}.com" if row.get('CompanyAbrv') else None,
        f"{row['FirstName'][0]}.{row['LastName']}@{row['CompanyAbrv']}.com" if row.get('CompanyAbrv') else None
    ]
    
    for email in email_formats:
        if email is None:
            continue
        try:
            msg = MIMEMultipart()
            msg['From'] = 'evanwaang2020@gmail.com'
            msg['To'] = 'evanwaang2020@gmail.com'
            msg['Subject'] = 'Your Subject Here'

            body = f"Hello {row['FirstName']},\n\nYour message here."
            msg.attach(MIMEText(body, 'plain'))
            
            smtp_server.sendmail('evanwaang2020@gmail.com', email, msg.as_string())
            return ('Yes', 'Sent successfully')
        except Exception as e:
            print(f"Error sending to {email}: {e}")
            continue
    return ('No', 'Failed to send')

# SMTP setup
smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
smtp_server.starttls()
smtp_server.login('evanwaang2020@gmail.com', 'cbse uicw hzjo jcmg')

# Read CSV
rows = []
with open('email_list.csv', 'r') as file:
    csv_reader = csv.DictReader(file)
    for row in csv_reader:
        rows.append(row)

# Update CSV
with open('email_list.csv', 'w') as file:
    fieldnames = ['FirstName', 'LastName', 'Company', 'CompanyAbrv', 'Sent', 'Notes']
    csv_writer = csv.DictWriter(file, fieldnames=fieldnames)
    csv_writer.writeheader()
    for row in rows:
        print(row.keys())
        if row['Sent'] == 'No':
            sent_status, notes = try_send_email(row, smtp_server)
            time.sleep(10)
            row['Sent'] = sent_status
            row['Notes'] = notes
        csv_writer.writerow(row)

smtp_server.quit()
