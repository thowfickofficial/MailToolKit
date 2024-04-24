import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from ttkthemes import ThemedStyle

def send_email(sender_email, sender_password, recipient_email, subject, content):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(content, 'plain'))

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender_email, sender_password)
    server.sendmail(sender_email, recipient_email, msg.as_string())
    server.quit()

def send_emails():
    sender_email = sender_email_entry.get()
    sender_password = sender_password_entry.get()
    excel_file_path = filedialog.askopenfilename(title="Select Excel File")

    try:
        wb = openpyxl.load_workbook(excel_file_path)
        sheet = wb.active

        for row in sheet.iter_rows(min_row=2, values_only=True):
            recipient_email, subject, content = row
            send_email(sender_email, sender_password, recipient_email, subject, content)
            status_label.config(text=f"Email sent to: {recipient_email}")

        status_label.config(text="Bulk email sending completed!")

    except Exception as e:
        status_label.config(text=f"An error occurred: {e}")

# Create the GUI window
root = tk.Tk()
root.title("Bulk Email Sender")

# Apply themed style with the "black" theme
style = ThemedStyle(root)
style.set_theme("black")

# Set background color for the frame
style.configure("TFrame", background="#202020")

# Create and style the widgets using ttk
frame = ttk.Frame(root)
frame.pack(padx=20, pady=20)

sender_email_label = ttk.Label(frame, text="Your Gmail address:", foreground="white")
sender_email_label.pack(pady=10)

sender_email_entry = ttk.Entry(frame, foreground="black")  # Change text color here
sender_email_entry.pack(pady=5)

sender_password_label = ttk.Label(frame, text="Your Gmail password:", foreground="white")
sender_password_label.pack()

sender_password_entry = ttk.Entry(frame, show="*", foreground="black")  # Change text color here
sender_password_entry.pack(pady=5)

# Create a button-like label
select_excel_button = tk.Label(frame, text="Select Excel File", bg="#007ACC", fg="white", cursor="hand2")
select_excel_button.pack(pady=10)

# Bind the label to the button functionality
select_excel_button.bind("<Button-1>", lambda event: send_emails())

status_label = ttk.Label(frame, text="", font=("Helvetica", 12, "italic"), foreground="white")
status_label.pack(pady=10)

root.mainloop()
