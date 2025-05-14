import sys
from email.mime.base import MIMEBase

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QTextEdit, QPushButton, QHBoxLayout, QFileDialog, \
    QLineEdit, QLabel

import pandas as pd

from send_message import send_message_from_excel


class EmailMassMailingApp(QWidget):
    def __init__(self):
        super().__init__()

        self.excel=''

        width_input = 300
        width_button = 200
        width_text = 500
        height_text_field = 500

        self.setWindowTitle("Приложение для рассылки КП")

        # Основной вертикальный layout
        layout = QHBoxLayout()
        # Горизонтальный layout для кнопки
        right_layout = QVBoxLayout()
        right_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        # Горизонтальный layout для ввода email и password_imap
        left_layout = QVBoxLayout()

        # Поля ввода
        self.email_input = QLineEdit()
        self.password_imap_input = QLineEdit()
        self.email_input.setFixedWidth(width_input)
        self.password_imap_input.setFixedWidth(width_input)
        self.email_input.setPlaceholderText('Введите ваш Email🐶')
        self.password_imap_input.setPlaceholderText('Введите ваш пароль IMAP🔐')
        self.password_imap_input.setEchoMode(QLineEdit.EchoMode.Password)

        # Ну и я решил в левый layout добавить кнопку выбора файла
        button_file = QPushButton("Выбрать файл")
        button_file.setFixedWidth(width_button)
        button_file.clicked.connect(self.show_file_dialog)

        self.excel_path_output = QTextEdit()
        self.excel_path_output.setFixedWidth(width_text)
        self.excel_path_output.setFixedHeight(height_text_field)
        self.excel_path_output.setReadOnly(True)  # Делаем текстовое поле только для чтения
        self.excel_path_output.setStyleSheet("background-color: white; color: black; border: 2px solid black; font-size: 10px;")  # Черный фон, белый текст

        # Добавляем виджеты в левый layout
        left_layout.addWidget(QLabel('Email: '))
        left_layout.addWidget(self.email_input)
        left_layout.addWidget(QLabel('Password IMAP: '))
        left_layout.addWidget(self.password_imap_input)
        left_layout.addWidget(button_file)
        left_layout.addStretch()
        left_layout.addWidget(self.excel_path_output)

        # QTextEdit для вывода сообщений
        self.console_output = QTextEdit()
        self.console_output.setFixedWidth(width_text)
        self.console_output.setFixedHeight(height_text_field)
        self.console_output.setReadOnly(True)  # Делаем текстовое поле только для чтения
        self.console_output.setStyleSheet("background-color: black; color: white;")  # Черный фон, белый текст

        # Пример поля для ввода файла
        button_send_messages = QPushButton("Отослать")
        button_send_messages.setFixedWidth(width_button)
        button_send_messages.clicked.connect(
            lambda: send_message_from_excel(
                login=self.email_input.text(),
                password=self.password_imap_input.text(),
                excel=self.excel,
            )
        )

        # Добавляем кнопки в right_layout
        right_layout.addWidget(self.console_output)
        right_layout.addWidget(button_send_messages)
        right_layout.setAlignment(button_send_messages, Qt.AlignmentFlag.AlignHCenter)

        layout.addLayout(left_layout)
        layout.addLayout(right_layout)

        # Устанавливаем основной layout
        self.setLayout(layout)

    def show_file_dialog(self):
        # Пример вывода сообщений в консольное окно
        console_app.log_consol("Выберите файл📄")
        # Открываем диалог выбора файла
        options = QFileDialog.Option.DontUseNativeDialog
        excel_file, _ = QFileDialog.getOpenFileName(self, "Выбрать файл", "", "Все файлы (*);;Текстовые файлы (*.txt)", options=options)

        if excel_file:
            if not (excel_file.endswith('xlsx') or excel_file.endswith('xls')):
                console_app.log_consol(f'Выбранный файл должен быть формата xlsx либо xls😡')
                return False
            # Если файл выбран, обновляем метку с путем к файлу
            console_app.log_consol(f'Выбранный файл: {excel_file}')
            console_app.log_head_excel(excel_file)
            self.excel = excel_file
        return True

    def log_consol(self, message):
        """Метод для добавления сообщения в консольное окно."""
        print(message)
        self.console_output.append(message)

    def log_head_excel(self, excel_file):
        df = pd.read_excel(excel_file, sheet_name='Лист1')
        head = str(df.head().to_html(index=False))
        self.excel_path_output.append(f"<p>Не волнуйтесь,что видите только пять строк.<br>Это только пример<br>На самом деле он всё видит😁<br></p><b>Файл:</b> {excel_file}<br><br>{head}<br>")
        self.excel_path_output.setHtml(self.excel_path_output.toHtml())



if __name__ == "__main__":
    app = QApplication(sys.argv)
    console_app = EmailMassMailingApp()
    console_app.show()
    sys.exit(app.exec())
