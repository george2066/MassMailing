import sys

from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QTextEdit, QPushButton, QHBoxLayout, QFileDialog

from send_message import send_message


class ConsoleApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Консольное приложение")
        self.setGeometry(100, 100, 600, 400)

        # Основной вертикальный layout
        layout = QVBoxLayout()

        # QTextEdit для вывода сообщений
        self.console_output = QTextEdit()
        self.console_output.setReadOnly(True)  # Делаем текстовое поле только для чтения
        self.console_output.setStyleSheet("background-color: black; color: white;")  # Черный фон, белый текст

        layout.addWidget(self.console_output)

        # Горизонтальный layout для кнопок
        multi_format_layout = QHBoxLayout()

        # Пример поля для ввода файла

        # Пример кнопок
        button_file = QPushButton("Выбрать файл")
        button_file.clicked.connect(self.show_file_dialog)
        button1 = QPushButton("Отослать")

        # Добавляем кнопки в горизонтальный layout
        multi_format_layout.addWidget(button_file)
        multi_format_layout.addWidget(button1)

        layout.addLayout(multi_format_layout)

        # Устанавливаем основной layout
        self.setLayout(layout)

    def show_file_dialog(self):
        # Пример вывода сообщений в консольное окно
        console_app.log_message("Выберите файл📄")
        # Открываем диалог выбора файла
        options = QFileDialog.Option.DontUseNativeDialog
        excel, _ = QFileDialog.getOpenFileName(self, "Выбрать файл", "", "Все файлы (*);;Текстовые файлы (*.txt)", options=options)

        if excel:
            if not (excel.endswith('xlsx') or excel.endswith('xls')):
                console_app.log_message(f'Выбранный файл должен быть формата xlsx либо xls😡')
                return None
            # Если файл выбран, обновляем метку с путем к файлу
            console_app.log_message(f'Выбранный файл: {excel}')
        return None

    def log_message(self, message):
        """Метод для добавления сообщения в консольное окно."""
        print(message)
        self.console_output.append(message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    console_app = ConsoleApp()
    console_app.show()
    sys.exit(app.exec())
