import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from os.path import basename

sender = 'no-reply@dell.com'
receivers = ['phil.zhang@dell.com']

msg = MIMEMultipart()
msg['From'] = sender
msg['To'] = ", ".join(receivers)
msg['Date'] = formatdate(localtime=True)
msg['Subject'] = "SMTP email example"
text = """
This is a test message from rebot
"""
msg.attach(MIMEText(text))

files = ["gdb.pdf"]
for f in files or []:
    with open(f, "rb") as fil:
        part = MIMEApplication(
            fil.read(),
            Name=basename(f)
        )
    # After the file is closed
    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
    msg.attach(part)

try:
    smtpObj = smtplib.SMTP('exch2019mail.ins.dell.com')
    smtpObj.sendmail(sender, receivers, msg.as_string())
    print("Successfully sent email")
except smtplib.SMTPException:
    pass