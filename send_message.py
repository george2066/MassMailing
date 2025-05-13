import smtplib
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def send_message(login, password_imap, to_email, header, text, file_pathes):
    message = MIMEMultipart()
    message['Subject'] = Header(header, 'utf-8')
    message['From'] = login
    message['To'] = to_email
    s = smtplib.SMTP('smtp.yandex.ru', 587, timeout=10)

    append_files(message, file_pathes)
    message.attach(MIMEText(f'{text}', 'plain', 'utf-8'))

    try:
        s.starttls()
        s.login(login, password_imap)
        s.sendmail(message['From'], message['To'], message.as_string())
    except Exception as e:
        print(e)
    finally:
        s.quit()

def append_files(message, files):
    for file_path in files:
        with open(file_path, "rb") as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={file_path}")
        message.attach(part)