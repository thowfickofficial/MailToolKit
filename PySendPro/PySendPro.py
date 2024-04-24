import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
import os
import uuid
import datetime
from apscheduler.schedulers.background import BackgroundScheduler

# Define a dictionary of SMTP server and port options for various email providers
smtp_options = {
    'gmail': ('smtp.gmail.com', 587),
    'outlook': ('smtp.office365.com', 587),
    'yahoo': ('smtp.mail.yahoo.com', 587),
    'hostinger': ('smtp.hostinger.com', 587),
    'aol': ('smtp.aol.com', 587),
    'icloud': ('smtp.mail.me.com', 587),
    'zoho': ('smtp.zoho.com', 587),
    'mailgun': ('smtp.mailgun.org', 587),
    'sendgrid': ('smtp.sendgrid.net', 587),
    'yandex': ('smtp.yandex.com', 587),
    'protonmail': ('smtp.protonmail.com', 587),
    'mailchimp': ('smtp.mailchimp.com', 587),
    'mailjet': ('in-v3.mailjet.com', 587),
    'ses': ('email-smtp.us-east-1.amazonaws.com', 587),
    'fastmail': ('smtp.fastmail.com', 587),
    'gmx': ('smtp.gmx.com', 587),
    'zimbra': ('smtp.zimbra.com', 587),
    'inmotionhosting': ('smtp.inmotionhosting.com', 587),
    'networksolutions': ('smtp.networksolutions.com', 587),
    'web.de': ('smtp.web.de', 587),
    'comcast': ('smtp.comcast.net', 587)
}

def send_email(sender_email, sender_password, recipient_email, subject, content, attachments=None, inline_images=None, is_html=False, smtp_server="smtp.gmail.com", smtp_port=587):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email

    if is_html:
        msg.attach(MIMEText(content, 'html'))
    else:
        msg.attach(MIMEText(content, 'plain'))

    msg['Subject'] = subject

    if attachments:
        for attachment_path in attachments:
            attachment = open(attachment_path, 'rb')
            part = MIMEApplication(attachment.read(), Name=os.path.basename(attachment_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
            msg.attach(part)

    if inline_images:
        for image_path in inline_images:
            with open(image_path, 'rb') as image_file:
                image_id = str(uuid.uuid4())
                image_data = image_file.read()
                image = MIMEImage(image_data)
                image.add_header('Content-ID', f'<{image_id}>')
                msg.attach(image)
                content = content.replace(image_path, f'cid:{image_id}')

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)
    server.sendmail(sender_email, recipient_email, msg.as_string())
    server.quit()

def schedule_emails(email_data):
    scheduler = BackgroundScheduler()

    for data in email_data:
        recipient_email, subject, content, attachments, inline_images, is_html, smtp_server, smtp_port, scheduled_time = data
        scheduler.add_job(send_email, 'date', run_date=scheduled_time, args=(sender_email, sender_password, recipient_email, subject, content, attachments, inline_images, is_html, smtp_server, smtp_port))

    scheduler.start()

def main():
    global sender_email, sender_password
    print("Welcome to PySendPro!")
    sender_email = input("Enter your email address: ")
    sender_password = input("Enter your email password: ")  

    print("\nAvailable SMTP Options:")
    for i, provider in enumerate(smtp_options.keys(), start=1):
        print(f"{i}. {provider.capitalize()}")
    
    smtp_option = input("\nEnter the number of the SMTP option you want to use: ")

    if smtp_option in [str(i) for i in range(1, len(smtp_options) + 1)]:
        provider = list(smtp_options.keys())[int(smtp_option) - 1]
        smtp_server, smtp_port = smtp_options[provider]

        print(f"\nYou have selected {provider.capitalize()} as your SMTP option.")

        excel_file_path = input("Enter the path to the Excel file containing email data: ")
        is_html = input("Send emails as HTML? (y/n): ").lower() == 'y'

        try:
            wb = openpyxl.load_workbook(excel_file_path)
            sheet = wb.active

            default_subject_template = "Default Subject: Hello {name}"
            default_content_template = "Default Content: Hi {name},\n\nThis is a default email."

            email_data = []

            attachments = []

            attachment_option = input("\nDo you want to attach files? (y/n): ").lower()
            if attachment_option == 'y':
                while True:
                    attachment_path = input("Enter the path to an attachment (leave empty to finish): ")
                    if not attachment_path:
                        break
                    attachments.append(attachment_path)

            inline_images = []

            inline_image_option = input("\nDo you want to attach inline images? (y/n): ").lower()
            if inline_image_option == 'y':
                while True:
                    image_path = input("Enter the path to an image (leave empty to finish): ")
                    if not image_path:
                        break
                    inline_images.append(image_path)

            for row in sheet.iter_rows(min_row=2, values_only=True):
                recipient_email, subject_template, content_template, scheduled_time = row

                # Replace placeholders in templates
                subject = subject_template.format(name=recipient_email.split('@')[0]) if subject_template else default_subject_template
                content = content_template.format(name=recipient_email.split('@')[0]) if content_template else default_content_template

                if not scheduled_time:
                    send_email(sender_email, sender_password, recipient_email, subject, content, attachments, inline_images, is_html, smtp_server, smtp_port)
                    print(f"Email sent to: {recipient_email}")
                else:
                    scheduled_time = datetime.datetime.strptime(scheduled_time, '%Y-%m-%d %H:%M:%S')
                    email_data.append((recipient_email, subject, content, attachments, inline_images, is_html, smtp_server, smtp_port, scheduled_time))

            if email_data:
                schedule_emails(email_data)
                print("\nEmails scheduled!")

        except Exception as e:
            print(f"An error occurred: {e}")

    else:
        print("\nInvalid SMTP option. Please select a valid number.")

if __name__ == "__main__":
    main()
