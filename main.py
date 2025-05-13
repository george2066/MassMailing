import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

import secret
from send_message import append_files

file_pathes = [
    'files\\ParkSoftDemo.zip',
    'files\\photo_2025-04-10_16-54-25.jpg',
    'files\\photo_2025-04-10_16-56-07.jpg',
    'files\\Изображение WhatsApp 2025-04-15 в 12.05.33_c66ddf15.jpg',
    'files\\Страница Cardpark.pdf',
    'files\\Телепарк дооборудование c gsm модемом.pdf',
    'files\\Телепарк дооборудование  на видеокамерах.pdf'
]

where_are_from = 'sales111@cardpark.su'
password = secret.PASSWORD_IMAP
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

df = pd.read_excel('data.xlsx', sheet_name='Лист1')

row = 1

while True:
    header = df.iloc[row, 6]
    text = text
    email = df.iloc[row, 3]
    if header == 'Генеральному директору ООО ""':
        print('Больше компаний не найдено😁')
        break
    s = smtplib.SMTP('smtp.yandex.ru', 587, timeout=10)
    s.starttls()
    s.login(where_are_from, password)
    message = MIMEMultipart()
    message['From'] = where_are_from
    message['To'] = email
    message = append_files(message, file_pathes)
    message.attach(MIMEText(f'{text}', 'plain', 'utf-8'))
    try:
        s.sendmail(message['From'], message['To'], message.as_string())
    except Exception as e:
        print(f'Мыло {email} недоступен😒: {e}')
        row += 1
        continue
    print(f'{header} отправлено✉️ на {email}')
    row += 1
    s.quit()



