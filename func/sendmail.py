import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def guimail(chude, at1, at2, danhsachgui, email, password):
    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = danhsachgui
    msg['Subject'] = chude

    body = "Kính gửi!"
    msg.attach(MIMEText(body, 'plain'))

    if at2 == "":
        attachment = open(at1, 'rb')
        part = MIMEApplication(attachment.read())
        attachment.close()
        part.add_header('Content-Disposition', 'attachment', filename=at1.split('\\')[-1])
        msg.attach(part)
    else:
        attachment1 = open(at1, 'rb')
        part1 = MIMEApplication(attachment1.read())
        attachment1.close()
        part1.add_header('Content-Disposition', 'attachment', filename=at1.split('\\')[-1])
        msg.attach(part1)

        attachment2 = open(at2, 'rb')
        part2 = MIMEApplication(attachment2.read())
        attachment2.close()
        part2.add_header('Content-Disposition', 'attachment', filename=at1.split('\\')[-1])
        msg.attach(part2)

    smtp = smtplib.SMTP('smtp.gmail.com', 587)
    smtp.starttls()
    smtp.login(email, password)
    smtp.send_message(msg)
    smtp.quit()

