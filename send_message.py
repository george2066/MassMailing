import smtplib

from PyQt6.QtWidgets import QTextEdit

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

def send_message_for_email(login, password_imap, header, text, login_from, files):
    # Создаем объект MIMEMultipart
    msg = MIMEMultipart()
    msg['From'] = login
    msg['To'] = login_from
    msg['Subject'] = header

    # Добавляем текст сообщения
    msg.attach(MIMEText(text, 'plain'))

    files = files[0]

    # Добавляем вложения
    for file_path in files:
        try:
            attachment = open(file_path, 'rb')
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={file_path.split("/")[-1]}')
            msg.attach(part)
            attachment.close()
        except Exception as e:
            return 'Чё-то файл недоступен🤔\n По логике он должен находится в папке files в корне проекта.'

    # Отправка сообщения
    try:
        with smtplib.SMTP('smtp.yandex.ru', 587) as server:  # Замените на SMTP-сервер Вашей почты
            server.starttls()
            server.login(login, password_imap)
            server.send_message(msg)
            return True
    except Exception as e:
        return f'Мыло {login_from} недоступен😒: {e}'

def send_message_from_excel(login, password, excel):
    text = '''
                    Добрый день!

                    Предлагаем рассмотреть сотрудничество между нашими компаниями в области бюджетной платной парковки, без парковочных стоек и без паркоматов.
                    Используемые идентификаторы-номер автомобиля и номер телефона. Оплата производится безналичным способом через Bэб-кассу.

                    Высылаем ссылку на техническую документацию по оборудованию Telepark.
                    Основной сайт Cardpark.su, на нем представлены классические системы производства нашей компании.
                    https://yar.diskstation.me/dokuwiki/doku.php?id=wiki:parking:dealer:modem
                    Пароль guest
                    Логин guest

                    Во вложении небольшая презентация и уникальное торговое предложение на самую бюджетную парковочную систему на рынке.

                    От вас необходимо получить стоимость установки типового шлагбаума на бетонное основание.


                    С Уважением,
                    Директор по развитию
                    Полуэктов Дмитрий Вячеславович
                    тел. +7 994 402 43 85, +7 910 202 19 91
                    e-mail: dirrazv@cardpark.su
                    '''
    file_pathes = [
        'files\\ParkSoftDemo.zip',
        'files\\photo_2025-04-10_16-54-25.jpg',
        'files\\photo_2025-04-10_16-56-07.jpg',
        'files\\Изображение WhatsApp 2025-04-15 в 12.05.33_c66ddf15.jpg',
        'files\\Страница Cardpark.pdf',
        'files\\Телепарк дооборудование c gsm модемом.pdf',
        'files\\Телепарк дооборудование  на видеокамерах.pdf'
    ],
    # Если необходимо, можно обработать Excel-файл
    try:
        df = pd.read_excel(excel)
        row = 0
        while True:
            login_from, header = df.iloc[row, 3], df.iloc[row, 6]
            if not type(login_from) == str:
                break
            try:
                send_message_for_email(login, password, header, text, login_from, file_pathes)
            except Exception as e:
                return e
            row += 1
            # Здесь можно добавить логику работы с данными из Excel, если это нужно
        return True
    except Exception as e:
        return f'Не удалось открыть Excel файл {excel}: {e}'
