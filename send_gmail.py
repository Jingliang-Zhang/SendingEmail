import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import basename
from pandas import read_excel

# ====== Global Settings to modify ======
subject = "On-running Account Statement_20240229"
body = """
Dear Customer:
        Please find attached your account statement as of Feb,29, 2024. 
        Please Kindly settle all overdue invoices by the end of this month.
        If you have any questions or missing any invoices, please let me know.
    
Cheers, 
Bowie
"""
sender = "bowie.yang@on-running.com"
password = "qhcqfveqvtrldkgq"
parentDir = "202311"


# ====== Code context, do not modify ======
def send_email(subject, body, sender, recipients, password, files):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    msg.attach(MIMEText(body))

    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
        smtp_server.login(sender, password)
        smtp_server.sendmail(sender, recipients, msg.as_string())
    print("Message sent!")


# read by default 1st sheet of an excel file
df = read_excel('customer statement name list.xlsx')
print(df)

for i in range(len(df)):
    recipients = df.iloc[i, 1].split(',')
    files = ["{}/{}".format(parentDir, df.iloc[i, 0])]
    print("Sending email to {account} with attachment {attachment}".format(
        account=recipients,
        attachment=files))

    # Following line is only for DEBUG purpose, DO NOT USE IT IN WORK!
    #recipients = ["bowie.yang@on-running.com", "657541672@qq.com"]

    send_email(subject=subject,
               body=body,
               sender=sender,
               recipients=recipients,
               password=password,
               files=files)
