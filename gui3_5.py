from PyQt6.QtGui import QCloseEvent, QIcon
from PyQt6.QtWidgets import (QApplication,QHBoxLayout, QWidget, QLabel, QLineEdit, QPushButton,
                             QVBoxLayout,QFileDialog,QTableWidget, QTableWidgetItem, QMessageBox,QStackedWidget, QTabWidget)
import json
from modules.scaner_module import ScanerParser
from modules.rubic_module import RubiScraper
from modules.pik_module import PIKScraper
import os

class ScanerScraperGUI(QWidget):
    def __init__(self):
        super().__init__()

        self.settings_filename = 'settings.json'
        icon_path = './icon.ico'
        self.setWindowIcon(QIcon(icon_path))
        self.settings = {
            'username':'',
            'password':'',
            'excel_filename':'',
            'excel_filename_rubic':'',
            'excel_filename_pik':'',
            'chromedriver':'',
        }
        self.init_ui()
        self.load_settings()
    
    def init_ui(self):
        self.setWindowTitle('Web Scraper GUI')

        # основной слой
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # создание side bar 
        self.sidebar_menu = QWidget()
        self.sidebar_layout = QVBoxLayout(self.sidebar_menu)
        self.sidebar_menu.setStyleSheet("""
            background-color: #2C3E50;
            color: #ECF0F1;
            font-size: 18px;
        """)

        self.button_scaner = QPushButton('Сканер')
        self.button_scaner.clicked.connect(lambda: self.change_tab(0))

        self.button_generator = QPushButton('Генератор')
        self.button_generator.clicked.connect(lambda: self.change_tab(1))

        self.button_pik = QPushButton('ПИК')
        self.button_pik.clicked.connect(lambda:self.change_tab(2))

        self.button_rubicon = QPushButton('Рубикон')
        self.button_rubicon.clicked.connect(lambda: self.change_tab(3))

        self.button_settings = QPushButton('Настройки')
        self.button_settings.clicked.connect(lambda: self.change_tab(4))

        # Кнопка "Сбор всего"
        self.button_scrape_all = QPushButton('Сбор всего')
        self.button_scrape_all.clicked.connect(self.scrape_all_data)

        # Добовление кнопок на виджет
        self.sidebar_layout.addWidget(self.button_scaner)
        self.sidebar_layout.addWidget(self.button_generator)
        self.sidebar_layout.addWidget(self.button_pik)
        self.sidebar_layout.addWidget(self.button_rubicon)
        self.sidebar_layout.addWidget(self.button_settings)
        self.sidebar_layout.addWidget(self.button_scrape_all)

        # хранение вкладок
        self.tab_widget = QStackedWidget(self)

        # Инициализация каждой вкладки
        self.init_main_tab()
        self.init_generator_tab()
        self.init_pik_tab()
        self.init_rubicon_tab()
        self.init_settings_tab()
        
        # вкладки в сложенный виджет
        self.tab_widget.addWidget(self.main_tab)
        self.tab_widget.addWidget(self.generator_tab)
        self.tab_widget.addWidget(self.pik_tab)
        self.tab_widget.addWidget(self.rubicon_tab)
        self.tab_widget.addWidget(self.settings_tab)

        layout.addWidget(self.sidebar_menu, 1)
        layout.addWidget(self.tab_widget, 5)

        self.setLayout(layout)

        self.load_settings()

    def init_main_tab(self):
        self.label_urls = QLabel('Ссылки для Сканера:')

        self.urls_table = QTableWidget()
        self.urls_table.setColumnCount(1)
        self.urls_table.setHorizontalHeaderLabels(['Ссылки'])
        self.urls_table.setRowCount(1)
        self.urls_table.setObjectName('urls_table_sacaner')
        self.urls_table.setColumnWidth(0, 240)

        self.button_add_url = QPushButton('Добавить новую ссылку')
        self.button_add_url.clicked.connect(self.add_url_entry)

        self.button_delete_url = QPushButton('Удалить выбранную ссылку')
        self.button_delete_url.clicked.connect(self.delete_url_entry)

        self.button_scrape = QPushButton('Собрать данные')
        self.button_scrape.clicked.connect(self.scrape_data)

        self.main_tab = QWidget()
        main_layout = QVBoxLayout(self.main_tab)

        main_layout.addWidget(self.label_urls)
        main_layout.addWidget(self.urls_table)

        main_layout.addWidget(self.button_add_url)
        main_layout.addWidget(self.button_delete_url)
        main_layout.addWidget(self.button_scrape)

    def init_settings_tab(self):
        self.label_login_settings = QLabel('Логин:')
        self.label_password_settings = QLabel('Пароль:')
        self.label_excel_path_settings = QLabel('Путь к Excel файлу:')
        self.label_excel_path_rubic_settings = QLabel('Путь к Excel файлу рубика')
        self.label_excel_path_pik_settings = QLabel('Путь к Excel файлу pik')
        self.label_chromedriver_path_settings = QLabel('Путь к chromedriver:')

        self.entry_login_settings = QLineEdit()
        self.entry_password_settings = QLineEdit()

        self.settings_entry_excel_path = QLineEdit()
        self.settings_entry_excel_path.setReadOnly(True)

        self.settings_entry_excel_rubic_path = QLineEdit()
        self.settings_entry_excel_rubic_path.setReadOnly(True)

        self.settings_entry_excel_pik_path = QLineEdit()
        self.settings_entry_excel_pik_path.setReadOnly(True)

        self.settings_entry_chromedriver_path = QLineEdit()
        self.settings_entry_chromedriver_path.setReadOnly(True)

        self.button_browse_excel_path = QPushButton('Обзор')
        self.button_browse_excel_path.clicked.connect(self.browse_excel_path)

        self.button_browse_excel_path_rubic = QPushButton('Обзор до рубика')
        self.button_browse_excel_path_rubic.clicked.connect(self.browse_excel_path_rubic)

        self.button_browse_excel_path_pik = QPushButton('Обзор до pik')
        self.button_browse_excel_path_pik.clicked.connect(self.browse_excel_path_pik)

        self.button_browse_chromedriver_path = QPushButton('Обзор chromedriver')
        self.button_browse_chromedriver_path.clicked.connect(self.browse_chromedriver_path)

        self.button_save_settings = QPushButton('Сохранить настройки')
        self.button_save_settings.clicked.connect(self.save_settings)

        self.settings_tab = QWidget()
        settings_layout = QVBoxLayout(self.settings_tab)

        settings_layout.addWidget(self.label_login_settings)
        settings_layout.addWidget(self.entry_login_settings)

        settings_layout.addWidget(self.label_password_settings)
        settings_layout.addWidget(self.entry_password_settings)

        settings_layout.addWidget(self.label_excel_path_settings)
        settings_layout.addWidget(self.settings_entry_excel_path)
        settings_layout.addWidget(self.button_browse_excel_path)

        settings_layout.addWidget(self.label_excel_path_rubic_settings)
        settings_layout.addWidget(self.settings_entry_excel_rubic_path)
        settings_layout.addWidget(self.button_browse_excel_path_rubic)

        settings_layout.addWidget(self.label_excel_path_pik_settings)
        settings_layout.addWidget(self.settings_entry_excel_pik_path)
        settings_layout.addWidget(self.button_browse_excel_path_pik)

        settings_layout.addWidget(self.label_chromedriver_path_settings)
        settings_layout.addWidget(self.settings_entry_chromedriver_path)
        settings_layout.addWidget(self.button_browse_chromedriver_path)

        settings_layout.addWidget(self.button_save_settings)

        # Добавим разделитель между блоками
        separator = QWidget()
        separator.setFixedHeight(1)
        separator.setStyleSheet("background-color: #BDC3C7;")

        settings_layout.addWidget(separator)

    def init_generator_tab(self):
        self.generator_label_urls = QLabel('Ссылки для Генератора:')

        self.generator_urls_table = QTableWidget()
        self.generator_urls_table.setColumnCount(1)
        self.generator_urls_table.setHorizontalHeaderLabels(['Ссылки'])
        self.generator_urls_table.setRowCount(1)
        self.generator_urls_table.setObjectName('urls_table_generator')
        self.generator_urls_table.setColumnWidth(0, 240)

        self.generator_button_add_url = QPushButton('Добавить новую ссылку')
        self.generator_button_add_url.clicked.connect(self.generator_add_url_entry)

        self.generator_button_delete_url = QPushButton('Удалить выбранную ссылку')
        self.generator_button_delete_url.clicked.connect(self.generator_delete_url_entry)

        self.generator_button_scrape = QPushButton('Собрать данные')
        # self.generator_button_scrape.clicked.connect(self.generator_scrape_data)
        
        self.generator_tab = QWidget()
        generator_layout = QVBoxLayout(self.generator_tab)

        generator_layout.addWidget(self.generator_label_urls)
        generator_layout.addWidget(self.generator_urls_table)

        generator_layout.addWidget(self.generator_button_add_url)
        generator_layout.addWidget(self.generator_button_delete_url)
        generator_layout.addWidget(self.generator_button_scrape)

    def init_pik_tab(self):
        self.pik_label_urls = QLabel("Ссылки для ПИК'а:")

        self.pik_urls_table = QTableWidget()
        self.pik_urls_table.setColumnCount(1)
        self.pik_urls_table.setHorizontalHeaderLabels(['Ссылки'])
        self.pik_urls_table.setRowCount(1)

        self.pik_urls_table.setColumnWidth(0, 240)

        self.pik_button_add_url = QPushButton('Добавить новую ссылку')
        self.pik_button_add_url.clicked.connect(self.pik_add_url_entry)

        self.pik_button_delete_url = QPushButton('Удалить выбранную ссылку')
        self.pik_button_delete_url.clicked.connect(self.pik_delete_url_entry)

        self.pik_button_scrape = QPushButton('Собрать данные')
        self.pik_button_scrape.clicked.connect(self.pik_scrape_data)

        self.pik_button_exit = QPushButton('Выход')
        self.pik_button_exit.clicked.connect(self.close)

        self.pik_tab = QWidget()
        pik_layout = QVBoxLayout(self.pik_tab)

        pik_layout.addWidget(self.pik_label_urls)
        pik_layout.addWidget(self.pik_urls_table)

        pik_layout.addWidget(self.pik_button_add_url)
        pik_layout.addWidget(self.pik_button_delete_url)
        pik_layout.addWidget(self.pik_button_scrape)

    def init_rubicon_tab(self):
        self.rubicon_label_urls = QLabel('Ссылки для Рубикона:')

        self.rubicon_urls_table = QTableWidget()
        self.rubicon_urls_table.setColumnCount(1)
        self.rubicon_urls_table.setHorizontalHeaderLabels(['Ссылки'])
        self.rubicon_urls_table.setRowCount(1)

        self.rubicon_urls_table.setColumnWidth(0, 240)

        self.rubicon_button_add_url = QPushButton('Добавить новую ссылку')
        self.rubicon_button_add_url.clicked.connect(self.rubicon_add_url_entry)

        self.rubicon_button_delete_url = QPushButton('Удалить выбранную ссылку')
        self.rubicon_button_delete_url.clicked.connect(self.rubicon_delete_url_entry)

        self.rubicon_button_scrape = QPushButton('Собрать данные')
        self.rubicon_button_scrape.clicked.connect(self.rubicon_scrape_data)

        self.rubicon_button_exit = QPushButton('Выход')
        self.rubicon_button_exit.clicked.connect(self.close)

        self.rubicon_tab = QWidget()
        rubicon_layout = QVBoxLayout(self.rubicon_tab)

        rubicon_layout.addWidget(self.rubicon_label_urls)
        rubicon_layout.addWidget(self.rubicon_urls_table)

        rubicon_layout.addWidget(self.rubicon_button_add_url)
        rubicon_layout.addWidget(self.rubicon_button_delete_url)
        rubicon_layout.addWidget(self.rubicon_button_scrape)

    #добавление url для сканера        
    def add_url_entry(self):
        current_row = self.urls_table.rowCount()
        self.urls_table.insertRow(current_row)
        self.urls_table.setItem(current_row, 0, QTableWidgetItem(''))

    #добавление url для генератора
    def generator_add_url_entry(self):
        generator_current_row = self.generator_urls_table.rowCount()
        self.generator_urls_table.insertRow(generator_current_row)
        self.generator_urls_table.setItem(generator_current_row, 0, QTableWidgetItem(''))

    #добавление url для рубикона  
    def rubicon_add_url_entry(self):
        rubicon_current_row = self.rubicon_urls_table.rowCount()
        self.rubicon_urls_table.insertRow(rubicon_current_row)
        self.rubicon_urls_table.setItem(rubicon_current_row, 0, QTableWidgetItem(''))
    
    #добавление url для PIK'a
    def pik_add_url_entry(self):
        pik_current_row = self.pik_urls_table.rowCount()
        self.pik_urls_table.insertRow(pik_current_row)
        self.rubicon_urls_table.setItem(pik_current_row, 0, QTableWidgetItem(''))

    #удаление url для сканера
    def delete_url_entry(self):
        selected_rows = self.urls_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Предупреждение', 'Выберите строку для удаления.')
            return

        for row in reversed(selected_rows):
            self.urls_table.removeRow(row.row())

    #удаление url для генератора
    def generator_delete_url_entry(self):
        selected_rows = self.generator_urls_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Предупреждение', 'Выберите строку для удаления.')
            return

        for row in reversed(selected_rows):
            self.generator_urls_table.removeRow(row.row())

    #удаление url для PIK'a
    def pik_delete_url_entry(self):
        selected_rows = self.pik_urls_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Предупреждение', 'Выберите строку для удаления.')
            return

        for row in reversed(selected_rows):
            self.pik_urls_table.removeRow(row.row()) 

    #удаление url для рубикона            
    def rubicon_delete_url_entry(self):
        selected_rows = self.rubicon_urls_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Предупреждение', 'Выберите строку для удаления.')
            return

        for row in reversed(selected_rows):
            self.rubicon_urls_table.removeRow(row.row())   

    #парсинг Сканера
    def scrape_data(self):
        username = self.entry_login_settings.text()
        password = self.entry_password_settings.text()
        excel_filename = self.settings_entry_excel_path.text()
        chromedriver_path = self.settings_entry_chromedriver_path.text()
        if not username or not password:
            QMessageBox.warning(self, 'Ошибка', 'Введите логин и пароль.')
            return

        urls = [self.urls_table.item(i, 0).text() for i in range(self.urls_table.rowCount())]

        
        web_scraper = ScanerParser(chromedriver_path, username, password,excel_filename)

        try:
            web_scraper.login()
            data_list = web_scraper.scrape_data(urls)
            web_scraper.update_excel(data_list)
            QMessageBox.information(self, 'Успех', 'Данные успешно добавлены в Excel файл.')
        except Exception as e:
            QMessageBox.warning(self, 'Ошибка', f'Произошла ошибка: {e}')

    #парсинг генератора
    def generator_scrape_data(self):
        username = self.entry_login_settings.text()
        password = self.entry_password_settings.text()
        excel_filename = self.settings_entry_excel_path.text()
        chromedriver_path = self.settings_entry_chromedriver_path.text()
        if not username or not password:
            QMessageBox.warning(self, 'Ошибка', 'Введите логин и пароль.')
            return

        urls = [self.generator_urls_table.item(i, 0).text() for i in range(self.generator_urls_table.rowCount())]
        
        web_scraper = ScanerParser(chromedriver_path, username, password,excel_filename)

        try:
            web_scraper.login()
            data_list = web_scraper.scrape_data(urls)
            web_scraper.update_excel(data_list)
            QMessageBox.information(self, 'Успех', 'Данные успешно добавлены в Excel файл.')
        except Exception as e:
            QMessageBox.warning(self, 'Ошибка', f'Произошла ошибка: {e}')

    #парсинг пик'а
    def pik_scrape_data(self):
        username = self.entry_login_settings.text()
        password = self.entry_password_settings.text()
        excel_filename = self.settings_entry_excel_pik_path.text()
        chromedriver_path = self.settings_entry_chromedriver_path.text()
        if not username or not password:
            QMessageBox.warning(self, 'Ошибка', 'Введите логин и пароль.')
            return

        urls = [self.pik_urls_table.item(i, 0).text() for i in range(self.pik_urls_table.rowCount())]

        
        web_scraper = PIKScraper(chromedriver_path, username, password,excel_filename)

        try:
            web_scraper.login()
            data_list = web_scraper.scrape_data(urls)
            web_scraper.update_excel(data_list)
            QMessageBox.information(self, 'Успех', 'Данные успешно добавлены в Excel файл.')
        except Exception as e:
            QMessageBox.warning(self, 'Ошибка', f'Произошла ошибка: {e}')

    #парсинг рубикона
    def rubicon_scrape_data(self):
        username = self.entry_login_settings.text()
        password = self.entry_password_settings.text()
        excel_filename = self.settings_entry_excel_rubic_path.text()
        chromedriver_path = self.settings_entry_chromedriver_path.text()
        if not username or not password:
            QMessageBox.warning(self, 'Ошибка', 'Введите логин и пароль.')
            return

        urls = [self.rubicon_urls_table.item(i, 0).text() for i in range(self.rubicon_urls_table.rowCount())]

        
        web_scraper = RubiScraper(chromedriver_path, username, password,excel_filename)

        try:
            web_scraper.login()
            data_list = web_scraper.scrape_data(urls)
            web_scraper.update_excel(data_list)
            QMessageBox.information(self, 'Успех', 'Данные успешно добавлены в Excel файл.')
        except Exception as e:
            QMessageBox.warning(self, 'Ошибка', f'Произошла ошибка: {e}')

    #расположение excel сканера
    def browse_excel_path(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, 'Выберите файл Excel', '', 'Excel Files (*.xlsx *.xls)')
        if file_path:
            self.settings_entry_excel_path.setText(file_path)

    #расположение excel рубика
    def browse_excel_path_rubic(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, 'Выберите файл Excel', '', 'Excel Files (*.xlsx *.xls)')
        if file_path:
            self.settings_entry_excel_rubic_path.setText(file_path)

    #расположение excel pik'a
    def browse_excel_path_pik(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, 'Выберите файл Excel', '', 'Excel Files (*.xlsx *.xls)')
        if file_path:
            self.settings_entry_excel_pik_path.setText(file_path)

    def browse_chromedriver_path(self):
        file_dialog = QFileDialog()
        file_path,_ = file_dialog.getOpenFileName(self, 'Выберите расположение chromedriver', '', 'chromedriver Files (*.exe)')
        if file_path:
            self.settings_entry_chromedriver_path.setText(file_path)

    def save_settings(self):
        self.settings['username'] = self.entry_login_settings.text()
        self.settings['password'] = self.entry_password_settings.text()
        self.settings['excel_filename'] = self.settings_entry_excel_path.text()
        self.settings['excel_filename_rubic'] = self.settings_entry_excel_rubic_path.text()
        self.settings['excel_filename_pik'] = self.settings_entry_excel_pik_path.text()
        self.settings['chromedriver'] = self.settings_entry_chromedriver_path.text()

        try:
            with open(self.settings_filename, 'w') as file:
                json.dump(self.settings, file)
            QMessageBox.information(self, 'Настройки сохранены', 'Настройки успешно сохранены.')
        except IOError as e:
            QMessageBox.warning(self, 'Ошибка', f'Ошибка при сохранении настроек: {e}')

    def load_settings(self):
        if os.path.exists(self.settings_filename):
            try:
                with open(self.settings_filename, 'r') as file:
                    self.settings = json.load(file)
            except json.JSONDecodeError as e:
                QMessageBox.warning(self, 'Ошибка', f'Ошибка при чтении файла настроек: {e}')
        else:
            #Создание если нет
            self.create_settings_file()

        self.entry_login_settings.setText(self.settings['username'])
        self.entry_password_settings.setText(self.settings['password'])
        self.settings_entry_excel_path.setText(self.settings['excel_filename'])
        self.settings_entry_excel_rubic_path.setText(self.settings['excel_filename_rubic'])
        self.settings_entry_excel_pik_path.setText(self.settings['excel_filename_pik'])
        self.settings_entry_chromedriver_path.setText(self.settings['chromedriver'])

    def change_tab(self, index):
        self.tab_widget.setCurrentIndex(index)

    def create_settings_file(self):
        try:
            with open(self.settings_filename, 'w') as file:
                json.dump(self.settings, file)
        except IOError as e:
            QMessageBox.warning(self, 'Ошибка', f'Ошибка при создании файла настроек: {e}')

    def scrape_all_data(self):
        #Сбор данных со сканера
        self.scrape_data()
        #Сбор данных с генератора
        #self.scrape_data_gen()
        self.pik_scrape_data()
        self.rubicon_scrape_data()
        
if __name__ == '__main__':
    app = QApplication([])
    app.setStyle('Fusion')
    gui = ScanerScraperGUI()
    gui.show()
    app.exec()