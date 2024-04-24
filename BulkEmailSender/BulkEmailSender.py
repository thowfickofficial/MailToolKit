import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

def send_email(sender_email, sender_password, recipient_email, subject, content):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(content, 'plain'))

    # Connect to Gmail's SMTP server
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender_email, sender_password)
    server.sendmail(sender_email, recipient_email, msg.as_string())
    server.quit()

def main():
    print("Welcome to BulkEmailSender!")
    sender_email = input("Enter your Gmail address: ")
    sender_password = input("Enter your Gmail password: ")  
    excel_file_path = input("Enter the path to the Excel file: ")

    try:
        wb = openpyxl.load_workbook(excel_file_path)
        sheet = wb.active

        for row in sheet.iter_rows(min_row=2, values_only=True):
            recipient_email, subject, content = row
            send_email(sender_email, sender_password, recipient_email, subject, content)
            print(f"Email sent to: {recipient_email}")

        print("Bulk email sending completed!")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
