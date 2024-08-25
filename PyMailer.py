import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os

email_service = input("Which email service are you using? (Gmail/Outlook): ").strip().lower()

if email_service == "gmail":
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
elif email_service == "outlook":
    smtp_server = "smtp-mail.outlook.com"
    smtp_port = 587
else:
    print("Unsupported email service. Please use Gmail or Outlook.")
    exit()

sender_email = input('Your Email Addresss: ')
receiver_email = input("Recepient's Email address: ")
password = input('What is your Password? ')
subject = input('What would you like the subject of this email to be? ')
body = input('What would you like the body of this email to be? ')

message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receiver_email
message["Subject"] = subject

message.attach(MIMEText(body, "plain"))

attach_file = input("Do you want to attach a file? (yes/no): ").strip().lower()

if attach_file == "yes":
    file_path = input("Enter the file path: ").strip()
    if os.path.isfile(file_path):
        try:
            with open(file_path, "rb") as attachment:
                part = MIMEApplication(attachment.read(), Name=os.path.basename(file_path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
                message.attach(part)
        except Exception as e:
            print(f"Could not attach the file: {e}")
    else:
        print("Invalid file path. Proceeding without attachment.")

try:
    server = smtplib.SMTP("smtp-mail.outlook.com", 587)
    server.starttls()
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, message.as_string())
    print("Email sent successfully!")
except Exception as e:
    print(f"Error: {e}")
finally:
    server.quit()
