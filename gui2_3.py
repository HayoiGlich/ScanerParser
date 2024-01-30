from PyQt6.QtGui import QCloseEvent, QIcon
from PyQt6.QtWidgets import (QApplication,QHBoxLayout, QWidget, QLabel, QLineEdit, QPushButton,
                             QVBoxLayout,QFileDialog,QTableWidget, QTableWidgetItem, QMessageBox, QTabWidget)
import json
from scaner_module import ScanerParser
import pandas as pd
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
            'chromedriver':'',
        }
        self.init_ui()
        self.load_settings()
    
    def init_ui(self):
        self.setWindowTitle('Web Parser GUI')

        #виджет с вкладками
        self.tab_widget = QTabWidget(self)

        #вкладка "Сканер"
        self.init_main_tab()

        #вкладка "Настройки"
        self.init_settings_tab()

        #добавление вкладок
        self.tab_widget.addTab(self.main_tab, "Сканер")
        self.tab_widget.addTab(self.settings_tab, "Настройки")

        layout = QVBoxLayout(self)
        layout.addWidget(self.tab_widget)
        self.setLayout(layout)

        self.load_settings()
        
    def init_main_tab(self):
        self.label_urls = QLabel('Ссылки:')

        self.urls_table = QTableWidget()
        self.urls_table.setColumnCount(1)
        self.urls_table.setHorizontalHeaderLabels(['Ссылки'])
        self.urls_table.setRowCount(1)

        self.urls_table.setColumnWidth(0, 240)

        self.button_add_url = QPushButton('Добавить новую ссылку')
        self.button_add_url.clicked.connect(self.add_url_entry)

        self.button_delete_url = QPushButton('Удалить выбранную ссылку')
        self.button_delete_url.clicked.connect(self.delete_url_entry)

        self.button_scrape = QPushButton('Собрать данные')
        self.button_scrape.clicked.connect(self.scrape_data)

        # self.button_exit = QPushButton('Выход')
        # self.button_exit.clicked.connect(self.close)

        self.main_tab = QWidget()
        main_layout = QVBoxLayout(self.main_tab)

        main_layout.addWidget(self.label_urls)
        main_layout.addWidget(self.urls_table)

        main_layout.addWidget(self.button_add_url)
        main_layout.addWidget(self.button_delete_url)
        main_layout.addWidget(self.button_scrape)
        # main_layout.addWidget(self.button_exit)

    def init_settings_tab(self):
        self.label_login_settings = QLabel('Логин:')
        self.label_password_settings = QLabel('Пароль:')
        self.label_excel_path_settings = QLabel('Путь к Excel файлу:')
        self.label_chromedriver_path_settings = QLabel('Путь к chromedriver:')

        self.entry_login_settings = QLineEdit()
        self.entry_password_settings = QLineEdit()

        self.settings_entry_excel_path = QLineEdit()
        self.settings_entry_excel_path.setReadOnly(True)

        self.settings_entry_chromedriver_path = QLineEdit()
        self.settings_entry_chromedriver_path.setReadOnly(True)

        self.button_browse_excel_path = QPushButton('Обзор')
        self.button_browse_excel_path.clicked.connect(self.browse_excel_path)

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

        settings_layout.addWidget(self.label_chromedriver_path_settings)
        settings_layout.addWidget(self.settings_entry_chromedriver_path)
        settings_layout.addWidget(self.button_browse_chromedriver_path)

        settings_layout.addWidget(self.button_save_settings)

    def add_url_entry(self):
        current_row = self.urls_table.rowCount()
        self.urls_table.insertRow(current_row)
        self.urls_table.setItem(current_row, 0, QTableWidgetItem(''))

    def delete_url_entry(self):
        selected_rows = self.urls_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'Предупреждение', 'Выберите строку для удаления.')
            return

        for row in reversed(selected_rows):
            self.urls_table.removeRow(row.row())

    def browse_excel_path(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, 'Выберите файл Excel', '', 'Excel Files (*.xlsx *.xls)')
        if file_path:
            self.settings_entry_excel_path.setText(file_path)

    def browse_chromedriver_path(self):
        file_dialog = QFileDialog()
        file_path,_ = file_dialog.getOpenFileName(self, 'Выберите расположение chromedriver', '', 'chromedriver Files (*.exe)')
        if file_path:
            self.settings_entry_chromedriver_path.setText(file_path)

    # def closeEvent(self, event):
    #     event.accept()

    def save_settings(self):
        self.settings['username'] = self.entry_login_settings.text()
        self.settings['password'] = self.entry_password_settings.text()
        self.settings['excel_filename'] = self.settings_entry_excel_path.text()
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
            self.create_settings_file()

        self.entry_login_settings.setText(self.settings['username'])
        self.entry_password_settings.setText(self.settings['password'])
        self.settings_entry_excel_path.setText(self.settings['excel_filename'])
        self.settings_entry_chromedriver_path.setText(self.settings['chromedriver'])

    def create_settings_file(self):
        try:
            with open(self.settings_filename, 'w') as file:
                json.dump(self.settings, file)
        except IOError as e:
            QMessageBox.warning(self, 'Ошибка', f'Ошибка при создании файла настроек: {e}')

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

if __name__ == '__main__':
    app = QApplication([])
    gui = ScanerScraperGUI()
    gui.show()
    app.exec()