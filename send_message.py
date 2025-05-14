import smtplib
from app import ConsoleApp
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd



def send_message_for_email(login, password_imap, text, excel, files, console_app: ConsoleApp):
    df = pd.read_excel(excel, sheet_name='Лист1')

    row = 1
    while True:
        header = df.iloc[row, 6]
        text = text
        email = df.iloc[row, 3]
        if header == 'Генеральному директору ООО ""' or type(header) == str:
            console_app.log_message('\nБольше компаний не найдено😁\nИли может вы забыли добавить в столбец премичаний заглавия письма?🧐\n')
            break
        s = smtplib.SMTP('smtp.yandex.ru', 587, timeout=10)
        s.starttls()

        s.login(login, password_imap)
        message = MIMEMultipart()
        message['Subject'] = header
        message['From'] = login
        message['To'] = email
        message = append_files(message, files)
        message.attach(MIMEText(f'{text}', 'plain', 'utf-8'))
        try:
            s.sendmail(message['From'], message['To'], message.as_string())
        except Exception as e:
            console_app.log_message(f'Мыло {email} недоступен😒: {e}')
            row += 1
            continue
        console_app.log_message(f'{header} отправлено✉️ на {email}')
        row += 1
        s.quit()


def append_files(message, files):
    for file_path in files:
        with open(file_path, "rb") as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={file_path}")
        message.attach(part)
    return message