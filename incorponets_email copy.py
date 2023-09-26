import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl


def send_email(subject, recipient, message):
    from_email = "your_email_here"
    password = "your_pw_here"
    to_email = recipient["email"]

    msg = MIMEMultipart()
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(message, "html"))

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(from_email, password)
    server.sendmail(from_email, to_email, msg.as_string())
    server.quit()


def main():
    global name, email, company
    excel_file = "recipients.xlsx"
    template_file = "email_template.html"
    subject = "Operacion Trinity"

    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active

    with open(template_file, "r") as f:
        template = f.read()

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if len(row) == 3:
            name, email, company = row
            personalized_message = template.format(name=name, company=company)
            recipient = {"email": email}
            send_email(subject, recipient, personalized_message)
            print(f"Email sent to {name} ({company}): {email}")
        else:
            print("Invalid row format:", row)



if __name__ == "__main__":
    main()
