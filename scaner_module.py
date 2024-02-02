from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import re
import pandas as pd
from datetime import datetime, timedelta, date
import calendar

class ScanerParser:
    def __init__(self, chromedriver_path, username, password,excel_filename):
        self.chromedriver_path = chromedriver_path
        self.options = webdriver.ChromeOptions()
        self.options.add_argument('--headless')
        self.options.add_argument('--disable-logging')
        self.browser = webdriver.Chrome(executable_path=self.chromedriver_path, options=self.options)
        #спрятать консоль webdriver'а
        script = '''
        var consoleLog = console.log;
        console.log = function() {};
        '''
        self.browser.execute_cdp_cmd("Runtime.evaluate", {"expression": script})
        self.username = username
        self.password = password
        self.excel_filename = excel_filename

    def login(self):
        self.browser.get(r'https://tasks.etecs.ru/login')
        email = self.browser.find_element(By.ID, "username")
        password_field = self.browser.find_element(By.ID, "password")
        login_button = self.browser.find_element(By.NAME, 'login')

        email.send_keys(self.username)
        password_field.send_keys(self.password)
        login_button.click()

    def scrape_data(self, urls):
        data_list = []

        for url in urls:
            self.browser.get(url)
            required_html = self.browser.page_source
            soup = BeautifulSoup(required_html, 'lxml')
            current_date = datetime.now().strftime("%d.%m.%Y")

            name_org = self.parse_name(soup)
            ip_mass_1, ip_mass_2 = self.parse_ip(soup)
            serial_numbers = self.parse_serial_number(soup)
            date_production_str = self.parse_date(soup, current_date)
            date_ranges = self.parse_date_end(soup)
            date_end = soup.find('div', class_='cf_48').find('div', class_='value').getText()
            execution = self.parse_execution(soup)
            postavka = self.parse_postavka(soup, execution)

            #отслеживания добавленных серийных номеров и дат
            added_serial_numbers = set()
            added_years_dict = {}

            #обработка серийных номеров в зависимости от их количества
            if len(serial_numbers) == 1:
                #если один серийный номер
                serial_number = serial_numbers[0]
                if serial_number not in added_serial_numbers:
                    added_serial_numbers.add(serial_number)
                    print(f"До записи в excel: {serial_number}")
                    for date_range in date_ranges:
                        if self.check_and_add_year(serial_number, date_range[0].year, added_years_dict):
                            data_list.append({
                                'Наименование организации': name_org,
                                'S/N': serial_number,
                                'Исполнение': execution,
                                'Дата производства': date_production_str,
                                'Дата начала лицензии': date_range[0].strftime("%d.%m.%Y"),
                                'Дата окончания лицензии': date_range[1].strftime("%d.%m.%Y"),
                                'IP': ip_mass_1,
                                'Доп.IP': ' ',
                                'IP дубль': ip_mass_1,
                                'Дата окончания': date_end,
                                'Серийник дубль': serial_number,
                                'Наименовение организации': name_org.join("''"),
                                'Тип поставки': postavka
                            })

                            if ip_mass_2 is not None:
                                data_list.append({
                                    'Наименование организации': name_org,
                                    'S/N': serial_number,
                                    'Исполнение': execution,
                                    'Дата производства': date_production_str,
                                    'Дата начала лицензии': date_range[0].strftime("%d.%m.%Y"),
                                    'Дата окончания лицензии': date_range[1].strftime("%d.%m.%Y"),
                                    'IP': ip_mass_2,
                                    'Доп.IP': 'Дополнительных',
                                    'IP дубль': int(ip_mass_1) + int(ip_mass_2),
                                    'Дата окончания': date_end,
                                    'Серийник дубль': serial_number,
                                    'Наименовение организации': name_org.join("''"),
                                    'Тип поставки': postavka
                                })

            else:
                #если несколько серийных номеров
                for serial_number in serial_numbers:
                    if serial_number not in added_serial_numbers:
                        added_serial_numbers.add(serial_number)
                        print(f"До записи в excel: {serial_number}")
                        for date_range in date_ranges:
                            if self.check_and_add_year(serial_number, date_range[0].year, added_years_dict):
                                data_list.append({
                                    'Наименование организации': name_org,
                                    'S/N': serial_number,
                                    'Исполнение': execution,
                                    'Дата производства': date_production_str,
                                    'Дата начала лицензии': date_range[0].strftime("%d.%m.%Y"),
                                    'Дата окончания лицензии': date_range[1].strftime("%d.%m.%Y"),
                                    'IP': ip_mass_1,
                                    'Доп.IP': ' ',
                                    'IP дубль': ip_mass_1,
                                    'Дата окончания': date_range,
                                    'Серийник дубль': serial_number,
                                    'Наименовение организации': name_org.join("''"),
                                    'Тип поставки': postavka
                                })

                                if ip_mass_2 is not None:
                                    data_list.append({
                                        'Наименование организации': name_org,
                                        'S/N': serial_number,
                                        'Исполнение': execution,
                                        'Дата производства': date_production_str,
                                        'Дата начала лицензии': date_range[0].strftime("%d.%m.%Y"),
                                        'Дата окончания лицензии': date_range[1].strftime("%d.%m.%Y"),
                                        'IP': ip_mass_2,
                                        'Доп.IP': 'Дополнительных',
                                        'IP дубль': int(ip_mass_1) + int(ip_mass_2),
                                        'Дата окончания': date_end,
                                        'Серийник дубль': serial_number,
                                        'Наименовение организации': name_org.join("''"),
                                        'Тип поставки': postavka
                                    })

        return data_list
    
    def check_and_add_year(self, serial_number, year, added_years_dict):
        if serial_number not in added_years_dict:
            added_years_dict[serial_number] = {year}
            return True
        elif year not in added_years_dict[serial_number]:
            added_years_dict[serial_number].add(year)
            return True
        else:
            return False
        
    def parse_date(self, soup, current_date):
        data_start = soup.find('div', class_='cf_46').find('div', class_='value').getText()

        #преобразование даты в формат datetime
        date_start = datetime.strptime(data_start, "%d.%m.%Y")

        #если date_start меньше current_date
        if date_start < datetime.combine(date.today(), datetime.min.time()):
            date_production = date_start
            #"Дата производства" на один день раньше "Дата начала лицензии"
            date_production = date_production - timedelta(days=1)
            if date_start.weekday() == calendar.MONDAY:
                date_production = date_production - timedelta(days=2)
        else:
            date_start = datetime.combine(date.today(), datetime.min.time())
            date_production = date_start
            if date_start.weekday() == calendar.MONDAY:
                date_production = date_production + timedelta(days=4)

        date_production_str = date_production.strftime("%d.%m.%Y")

        return date_production_str
        
    def parse_date_end(self, soup):
        date_start_str = soup.find('div', class_='cf_46').find('div', class_='value').getText()
        date_end_str = soup.find('div', class_='cf_48').find('div', class_='value').getText()

        date_start = datetime.strptime(date_start_str, "%d.%m.%Y")
        date_end = datetime.strptime(date_end_str, "%d.%m.%Y")

        #проверка на разницу в годах
        year_diff = date_end.year - date_start.year

        if year_diff > 0:
            #последовательность дат
            date_ranges = self.generate_date_range(date_start, date_end)
            return date_ranges
        elif year_diff == 0:
            #если разница в годах = 0 то возвращаем диапазон
            return [(date_start, date_end)]
        else:
            #ошибка если разница отрицательна
            raise ValueError("Ошибка в разнице годов. Дата окончания раньше даты начала.")

    def generate_date_range(self, start_date, end_date):
        date_range = []

        #если разница в годах > 1, ген даты по порядку года
        if end_date.year - start_date.year > 1:
            current_date = start_date
            while current_date.year < end_date.year:
                next_year_date = datetime(current_date.year + 1, start_date.month, start_date.day)
                date_range.append((current_date, next_year_date - timedelta(days=1)))
                current_date = next_year_date
        else:
            date_range.append((start_date, end_date))

        return date_range
    
    def parse_name(self, soup):
        title_text = soup.title.text
        match = re.search("для .*? -", title_text)

        if match:
            name_org = match.group()[4:-1]
            return name_org

        return None

    def parse_ip(self, soup):
        ip_mass = soup.find('div', class_="cf_43").find('div', class_='value').getText()
        if ip_mass == None or ip_mass == "-":
            ip_mass = soup.find('div', class_="wiki").find('p').getText()
            match = re.search(r'(\d+)\D*\+\D*(\d+)', ip_mass)
            if match:
                ip_mass_1, ip_mass_2 = match.groups()
            else:
                ip_mass_1 = ' '.join(re.findall(r'\d+', ip_mass))
                ip_mass_2 = None
        else:
            ip_mass_1 = re.search(r'\d+', ip_mass).group()
            ip_mass_2 = None

        return ip_mass_1, ip_mass_2
    
    def parse_serial_number(self, soup):
        serial_number = soup.find('div', class_='cf_51').find('div', class_='value').getText()
        # регулярное выражение совмещающее патерн серийников
        pattern = re.compile(r'(.*?-(\d+)-.*?-(\d+))|(0060601\.22\.(\d+)-0060601\.22\.(\d+)|(0060601\.21\.(\d+)-0060601\.21\.(\d+)))')
        match = pattern.search(serial_number)

        result_list = []

        if match:
            if match.group(2) is not None and match.group(3) is not None:
                serial_number1 = int(match.group(2))
                serial_number2 = int(match.group(3))
                result_list.extend([f"ЭФ2204-{number:06d}" for number in range(serial_number1, serial_number2 + 1)])
            elif match.group(4) is not None and match.group(5) is not None:
                serial_number3 = int(match.group(4))
                serial_number4 = int(match.group(5))
                result_list.extend([f"0060601.22.{number:04d}" for number in range(serial_number3, serial_number4 + 1)])
            else:
                result_list.append(serial_number)
        else:
            parts = serial_number.split('-')
            if len(parts) == 2 and all(part.isdigit() for part in parts):
                start_value, end_value = map(int, parts)
                result_list.extend([f"{int(number):09d}" for number in range(start_value, end_value + 1)])
            else:
                return [serial_number]

        return result_list

    def parse_execution(self, soup):
        execution = soup.title
        execution = re.search(r"[И,и]нспектор", execution.text)

        return str(2) if execution else str(1)

    def parse_postavka(self,soup, execution):
        postavka = soup.find('div',class_='cf_25').find('div',class_='value').getText()
        if postavka == 'ФСТЭК':
            return f'f hidef' if execution == '1' else f'fi hidef'
        elif postavka == 'МО РФ' or postavka == 'МО РФ с ВП':
            return f'm hidef' if execution == '1' else f'mi hidef'

    def update_excel(self, data_list):
        try:
            df = pd.read_excel(self.excel_filename,dtype={'S/N': str,'Серийник дубль':str},)
        except FileNotFoundError:
            df = pd.DataFrame()

        for new_data in data_list:
            if_not_empty_df = not df.empty

            if if_not_empty_df:
                if_have_double = df.apply(lambda row: 
                    row['Наименование организации'] == new_data['Наименование организации'] and 
                    any(sn in row['S/N'] if isinstance(row['S/N'], list) else row['S/N'] == sn for sn in new_data['S/N']),
                    axis=1)

                if if_have_double.any():
                    new_rows = pd.DataFrame([{'Наименование организации': new_data['Наименование организации'],
                                            'S/N': sn} for sn in new_data['S/N']])
                    df = pd.concat([df, new_rows], ignore_index=True)
                else:
                    df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
            else:
                df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
        writer = pd.ExcelWriter(self.excel_filename, engine='xlsxwriter') 
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        text_format = writer.book.add_format({'num_format': '@'})

        worksheet = writer.sheets['Sheet1']
        for idx, col in enumerate(df):
            series = df[col]
            max_len = max((series.astype(str).map(len).max(), len(str(col)))) + 2
            worksheet.set_column(idx, idx, max_len,text_format)

        writer.close()
        # df.to_excel(self.excel_filename, index=False)