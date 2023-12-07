import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
from datetime import datetime


def attach_file(msg, filepath):
    if filepath:  # Only attach if there's a file
        with open(filepath, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filepath.split('/')[-1]}")
            msg.attach(part)


def format_names(names_list):
    if len(names_list) == 1:
        return names_list[0]
    elif len(names_list) == 2:
        return f"{names_list[0]} and {names_list[1]}"
    else:
        all_except_last = ', '.join(names_list[:-1])
        return f"{all_except_last}, and {names_list[-1]}"


def send_email(subject, recipient_emails, cc_emails, message, attachment_path):
    from_email = "xyz@email.com"
    password = "your_pw_here"  # pw o app-gen pw

    msg = MIMEMultipart()
    msg["From"] = from_email
    msg["To"] = recipient_emails
    msg["CC"] = "boss@email.com, boss_assistant@email.com".join(cc_emails) # put myself on cc
    msg["Subject"] = subject
    msg.attach(MIMEText(message, "html"))

    attach_file(msg, attachment_path)

    recipients = recipient_emails + cc_emails
    server = smtplib.SMTP("smtp.office365.com", 587)
    server.starttls()
    server.login(from_email, password)
    server.sendmail(from_email, recipients, msg.as_string())
    server.quit()


def main():
    excel_file = "recipients.xlsx"
    template_file = "email_template.html"
    current_month = datetime.now().strftime('%B')
    subject = f"TEST {current_month} Reporting Updates"
    cc_emails = {"boss@email.com, boss_assistant@email.com"}

    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active

    with open(template_file, "r") as f:
        template = f.read()

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if len(row) >= 4:
            names_string, email_string, agency, attachment_path, custom_text = row[:5]
            names = [name.strip() for name in names_string.split(',')]
            formatted_names = format_names(names)
            recipient_emails = [email.strip() for email in email_string.split(',')]
            attachment_path = row[3] if len(row) > 3 else None
            custom_text = row[4] if len(row) > 4 else ""
            personalized_message = template.format(name=formatted_names, agency=agency, custom_text=custom_text)
            send_email(subject, recipient_emails, cc_emails, personalized_message, attachment_path)
            print(f"Email sent to {', '.join(recipient_emails)} ({agency})")
        else:
            print("Invalid row format:", row)


if __name__ == "__main__":
    main()
