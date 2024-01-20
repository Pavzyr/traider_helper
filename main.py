import ctypes
import os
import sys
import logging

from PyQt5.QtCore import QDate, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QFont
from selenium import webdriver
from general_downloader import make_hrefs_list, Forex4you, Litefinance

from general_download import run_main
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QGridLayout, QPlainTextEdit, \
    QDateEdit, QProgressBar


def open_file(msg, path):
    print(msg)
    os.startfile(path)


def open_folder(msg, path):
    print(msg)
    os.system(path)


def clean_folder(msg, folder_path):
    files = os.listdir(folder_path)
    for file in files:
        file_path = os.path.join(folder_path, file)
        os.remove(file_path)
    print(msg)


class WorkerThread(QThread):
    progress_update = pyqtSignal(int, bool, int, int, name='progressUpdate')  # Указываем сигналу имя для использования при соединении с методом

    def __init__(self, max_iterations, run_type):
        super().__init__()
        self.max_iterations = max_iterations
        self.ex = False
        self.run_type = run_type

    def open_browser(self):
        options = webdriver.ChromeOptions()
        options.add_argument('chromedriver_binary.chromedriver_filename')
        # options.add_argument('headless')
        options.add_argument("window-size=1920,1080")
        driver = webdriver.Chrome(options=options)
        driver.maximize_window()
        return driver

    def run(self):
        i = 0
        try:
            driver = self.open_browser()
            if self.run_type == 'all':
                for site in input_lists:
                    for href in site:
                        if href is None:
                            continue
                        elif 'forex4you' in href.value:
                            forex4you = Forex4you(
                                current_dir,
                                bd_dir,
                                driver,
                                href,
                                'forex4you',
                                forex4you_xpathes
                            )
                            forex4you.scrap_all()
                        elif 'litefinance' in href.value:
                            litefinance = Litefinance(
                                current_dir,
                                bd_dir,
                                driver,
                                href,
                                'litefinance',
                                litefinance_xpathes
                            )
                            litefinance.scrap_all()
                        i += 1
                        self.progress_update.emit(i, self.ex, self.max_iterations, i)
            elif self.run_type == 'litefinance':
                for href in litefinance_list:
                    if href is None:
                        continue
                    elif 'litefinance' in href.value:
                        litefinance = Litefinance(
                            current_dir,
                            bd_dir,
                            driver,
                            href,
                            'litefinance',
                            litefinance_xpathes
                        )
                        litefinance.scrap_all()
                    i += 1
                    self.progress_update.emit(i, self.ex, self.max_iterations, i)
            elif self.run_type == 'forex4you':
                for href in forex4you_list:
                    if href is None:
                        continue
                    elif 'forex4you' in href.value:
                        forex4you = Forex4you(
                            current_dir,
                            bd_dir,
                            driver,
                            href,
                            'forex4you',
                            forex4you_xpathes
                        )
                        forex4you.scrap_all()
                    i += 1
                    self.progress_update.emit(i, self.ex, self.max_iterations, i)
            driver.quit()
        except Exception as error:
            driver.quit()
            self.ex = True
            logging.error(error)
        self.progress_update.emit(0, self.ex, self.max_iterations, i)  # Завершаем выполнение операции


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Помошник трейдера')

        layout = QVBoxLayout()

        self.general_download_button = QPushButton('Скачать БАЗУ', self)
        self.general_download_button.clicked.connect(self.general_download)
        layout.addWidget(self.general_download_button)

        self.single_download_button = QPushButton('Скачать выборочно', self)
        self.single_download_button.clicked.connect(self.single_downloader)
        layout.addWidget(self.single_download_button)

        self.setLayout(layout)

        self.show()

    def general_download(self):
        self.hide()
        self.newWindow = GeneralDownload()
        self.newWindow.setFixedSize(670, 860)
        self.newWindow.show()

    def single_downloader(self):
        self.hide()
        self.newWindow = SingleDownload()
        self.newWindow.setFixedSize(670, 860)
        self.newWindow.show()


class GeneralDownload(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Скачать БАЗУ'
        self.initUI()


    def initUI(self):
        self.setWindowTitle(self.title)
        self.date_label = QLabel(self)
        self.date_label.setFixedSize(315, 45)
        self.date_label.setFont(QFont('Arial', 14))
        self.date_label.move(20, 10)
        self.date_label.setText('Выберите дату завершения сделки, '
                                '\nдо которой считывать сигналы:')

        self.work_progress_label = QLabel(self)
        self.work_progress_label.setFixedSize(315, 45)
        self.work_progress_label.setFont(QFont('Arial', 14))
        self.work_progress_label.move(20, 10)
        self.work_progress_label.setText('Процесс не запущен.')

        self.date_edit = QDateEdit(self)
        self.date_edit.setFixedSize(150, 35)
        self.date_edit.setFont(QFont('Arial', 14))
        default_date = QDate.currentDate()
        self.date_edit.setDate(default_date)
        self.date_edit.move(20, 70)
        self.date_edit.setCalendarPopup(True)

        self.show_start_button = QPushButton('🚀 Запуск', self)
        self.show_start_button.setFont(QFont('Arial', 14))
        self.show_start_button.setFixedSize(150, 36)
        self.show_start_button.move(20, 70)
        self.show_start_button.clicked.connect(self.main_proc)

        self.show_input_excel_button = QPushButton('   🧾 Открыть excel со ссылками', self)
        self.show_input_excel_button.setStyleSheet("text-align: left;")
        self.show_input_excel_button.setFont(QFont('Arial', 14))
        self.show_input_excel_button.setFixedSize(315, 36)
        self.show_input_excel_button.move(20, 125)
        self.show_input_excel_button.clicked.connect(lambda: open_file(
            "Открываю excel файл, где хранится список трейдеров, сигналы которых будут считываться.\n",
            rf"{current_dir}\resources\input\input hrefs.xlsx"))

        self.show_output_excel_button = QPushButton('   📁 Открыть папку с excel', self)
        self.show_output_excel_button.setFont(QFont('Arial', 14))
        self.show_output_excel_button.setStyleSheet("text-align: left;")
        self.show_output_excel_button.setFixedSize(315, 36)
        self.show_output_excel_button.move(20, 181)
        self.show_output_excel_button.clicked.connect(lambda: open_folder(
            "Открываю директорию, где хранятся сформированные файлы excel\n",
            rf"explorer.exe {current_dir}\resources\output excel"))

        self.show_output_htm_button = QPushButton('   📂 Открыть папку с htm', self)
        self.show_output_htm_button.setFont(QFont('Arial', 14))
        self.show_output_htm_button.setStyleSheet("text-align: left;")
        self.show_output_htm_button.setFixedSize(315, 36)
        self.show_output_htm_button.move(20, 237)
        self.show_output_htm_button.clicked.connect(lambda: open_folder(
            "Открываю директорию, где хранятся сформированные файлы htm\n",
            fr"explorer.exe {current_dir}\resources\output htm"))

        self.clean_output_htm_button = QPushButton('   🗑️ Очистить папку с htm', self)
        self.clean_output_htm_button.setFont(QFont('Arial', 14))
        self.clean_output_htm_button.setStyleSheet("text-align: left;")
        self.clean_output_htm_button.setFixedSize(315, 36)
        self.clean_output_htm_button.move(20, 237)
        self.clean_output_htm_button.clicked.connect(lambda: clean_folder('🗑️ Папка с htm очищена успешно\n',
                                                                          rf"{current_dir}\resources\output htm"))

        self.clean_output_excel_button = QPushButton('   🗑️ Очистить папку с excel', self)
        self.clean_output_excel_button.setFont(QFont('Arial', 14))
        self.clean_output_excel_button.setStyleSheet("text-align: left;")
        self.clean_output_excel_button.setFixedSize(315, 36)
        self.clean_output_excel_button.move(20, 237)
        self.clean_output_excel_button.clicked.connect(lambda: clean_folder('🗑️ Папка с excel очищена успешно\n',
                                                                            rf"{current_dir}\resources\output excel"))

        self.text_edit = QPlainTextEdit()
        self.text_edit.setFixedSize(650, 600)
        self.text_edit.setFont(QFont('Arial', 14))
        sys.stdout = self

        vbox1 = QGridLayout()
        vbox1.setSpacing(10)

        vbox = QGridLayout()
        vbox.setSpacing(10)
        vbox.addWidget(self.date_label, 0, 0)
        vbox.addLayout(vbox1, 0, 1, 1, 1)
        vbox1.addWidget(self.date_edit, 1, 0)
        vbox1.addWidget(self.show_start_button, 1, 1)
        vbox.addWidget(self.show_input_excel_button, 4, 0)
        vbox.addWidget(self.work_progress_label, 4, 1)
        vbox.addWidget(self.show_output_excel_button, 5, 1)
        vbox.addWidget(self.clean_output_excel_button, 5, 0)
        vbox.addWidget(self.show_output_htm_button, 6, 1)
        vbox.addWidget(self.clean_output_htm_button, 6, 0)
        vbox.addWidget(self.text_edit, 8, 0)
        self.setLayout(vbox)

    def write(self, text):
        self.text_edit.insertPlainText(text)

    def main_proc(self):
        date = self.date_edit.date().toString('dd.MM.yyyy')
        self.date_label.setText('Выбранная дата: ' + date)
        run_main(date)


class SingleDownload(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'Скачивание сигналов'
        self.widget = QWidget()
        layout = QVBoxLayout()
        self.widget.setLayout(layout)

        self.show_input_excel_button = QPushButton(
            '   🧾 Открыть excel litefinance', self)
        self.show_input_excel_button.setStyleSheet("text-align: left;")
        self.show_input_excel_button.setFont(QFont('Arial', 14))
        self.show_input_excel_button.clicked.connect(lambda: open_file(
            "Открываю excel файл, где хранится список ссылок для litefinance\n",
            rf"{current_dir}\resources\БАЗА ДАННЫХ\litefinance hrefs.xlsx"))
        layout.addWidget(self.show_input_excel_button)

        self.show_input_excel_button1 = QPushButton(
            '   🧾 Открыть excel forex4you', self)
        self.show_input_excel_button1.setStyleSheet("text-align: left;")
        self.show_input_excel_button1.setFont(QFont('Arial', 14))
        self.show_input_excel_button1.clicked.connect(lambda: open_file(
            "Открываю excel файл, где хранится список ссылок для forex4you\n",
            rf"{current_dir}\resources\БАЗА ДАННЫХ\forex4you hrefs.xlsx"))
        layout.addWidget(self.show_input_excel_button1)

        self.show_log_button = QPushButton('   🧾 Открыть логи', self)
        self.show_log_button.setStyleSheet("text-align: left;")
        self.show_log_button.setFont(QFont('Arial', 14))
        self.show_log_button.clicked.connect(lambda: open_file(
            "Открываю log файл, где хранятся ошибки работы робота",
            rf"{current_dir}\main.log"))
        layout.addWidget(self.show_log_button)

        self.show_output_excel_button = QPushButton(
            '   📁 Открыть папку litefinance', self)
        self.show_output_excel_button.setFont(QFont('Arial', 14))
        self.show_output_excel_button.setStyleSheet("text-align: left;")
        self.show_output_excel_button.clicked.connect(lambda: open_folder(
            "Открываю директорию, где хранятся сформированные файлы excel\n",
            rf"explorer.exe {current_dir}\resources\БАЗА ДАННЫХ\litefinance"))
        layout.addWidget(self.show_output_excel_button)

        self.show_output_htm_button = QPushButton(
            '   📂 Открыть папку forex4you', self)
        self.show_output_htm_button.setFont(QFont('Arial', 14))
        self.show_output_htm_button.setStyleSheet("text-align: left;")
        self.show_output_htm_button.clicked.connect(lambda: open_folder(
            "Открываю директорию, где хранятся сформированные файлы htm\n",
            fr"explorer.exe {current_dir}\resources\БАЗА ДАННЫХ\forex4you"))
        layout.addWidget(self.show_output_htm_button)

        self.label = QLabel('Ожидаю запуск')
        self.label.setFont(QFont('Arial', 14))
        layout.addWidget(self.label)

        self.progress_bar_loading = QProgressBar(self)
        self.progress_bar_loading.setFont(QFont('Arial', 14))
        self.progress_bar_loading.setVisible(False)
        layout.addWidget(self.progress_bar_loading)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.start_button = QPushButton('🚀 Общий запуск', self)
        self.start_button.setFont(QFont('Arial', 14))
        self.start_button.clicked.connect(self.run_all)
        layout.addWidget(self.start_button)

        self.start_button_lifefanance = QPushButton('🚀 Запуск только litefinance', self)
        self.start_button_lifefanance.setFont(QFont('Arial', 14))
        self.start_button_lifefanance.clicked.connect(self.run_litefinance)
        layout.addWidget(self.start_button_lifefanance)

        self.start_button_forex4you = QPushButton(
            '🚀 Запуск только forex4you', self)
        self.start_button_forex4you.setFont(QFont('Arial', 14))
        self.start_button_forex4you.clicked.connect(self.run_forex4you)
        layout.addWidget(self.start_button_forex4you)

        self.setCentralWidget(self.widget)

    def run_all(self):
        max_iterations = len(litefinance_list) + len(forex4you_list)
        self.progress_bar.setMaximum(max_iterations)
        self.start_operation(max_iterations, 'all')

    def run_litefinance(self):
        max_iterations = len(litefinance_list)
        self.progress_bar.setMaximum(max_iterations)
        self.start_operation(max_iterations, 'litefinance')

    def run_forex4you(self):
        max_iterations = len(forex4you_list)
        self.progress_bar.setMaximum(max_iterations)
        self.start_operation(max_iterations, 'forex4you')

    def start_operation(self, max_iterations, run_type):
        self.label.setText('Процесс выполняется...')
        self.progress_bar_loading.setVisible(True)
        self.progress_bar_loading.setRange(0, 0)
        self.widget.setEnabled(False)  # Блокируем кнопки
        self.worker_thread = WorkerThread(max_iterations, run_type)
        self.worker_thread.progressUpdate.connect(
        self.update_progress)  # Соединяем сигнал с обработчиком
        self.worker_thread.finished.connect(self.operation_completed)  # Соединяем сигнал о завершении с методом
        self.worker_thread.start()

    def update_progress(self, value, ex, max_iterations, current_itteration):
        self.progress_bar.setValue(value)
        self.label.setText(f'Процесс выполняется... Обработан {current_itteration} из {max_iterations}')
        if value == 0 and ex is False:
            self.progress_bar_loading.setVisible(False)
            self.label.setText('✅ Процесс завершен успешно')
        elif value == 0 and ex is True:
            self.progress_bar_loading.setVisible(False)
            self.label.setText('Процесс завершен с ошибкой. \nПроверьте логи!')

    def operation_completed(self):
        self.widget.setEnabled(True)  # Разблокируем кнопки


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.ERROR,
        filename='main.log',
        datefmt='%d.%m.%Y %H:%M:%S',
        filemode='w',
        format='%(asctime)s, %(levelname)s, %(message)s'
    )

    current_dir = os.path.dirname(os.path.abspath(__file__))
    bd_dir = current_dir + r'\resources\БАЗА ДАННЫХ'
    litefinance_list = make_hrefs_list(bd_dir + r'\litefinance hrefs.xlsx')
    forex4you_list = make_hrefs_list(bd_dir + r'\forex4you hrefs.xlsx')
    input_lists = [make_hrefs_list(bd_dir + r'\litefinance hrefs.xlsx'),
                   make_hrefs_list(bd_dir + r'\forex4you hrefs.xlsx')]
    litefinance_xpathes = {
        'trader_name': fr'//div[@class = "page_header_part traders_body"]//h2'
    }
    forex4you_xpathes = {
        'trader_name': fr'//span[@data-ng-bind= "::$headerCtrl.leader.displayName"]'
    }

    # скрываем консоль при запуске из батника
    console_window = ctypes.windll.kernel32.GetConsoleWindow()
    ctypes.windll.user32.ShowWindow(console_window, 6)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    ex = MainWindow()
    ex.setFixedSize(315, 200)
    sys.exit(app.exec_())