from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import re
import pandas as pd
from datetime import datetime, timedelta, date
import calendar

class PIKScraper:
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
            execution = self.parse_ispolnenie(soup)
            serial_numbers = self.parse_serial_number(soup)
            data_prodaction = self.parse_date(soup)
            date_ranges = self.parse_date_end(soup)

            added_serial_numbers = set()
            added_years_dict = {}    

            if len(serial_numbers) == 1:
                #если один серийный номер
                serial_number = serial_numbers[0]
                if serial_number not in added_serial_numbers:
                    added_serial_numbers.add(serial_number)
                    for date_range in date_ranges:
                        if self.check_and_add_year(serial_number, date_range[0].year, added_years_dict):
                            data_list.append({
                                                'заказчик': name_org,
                                                'Дата производства': data_prodaction,
                                                'серийный': serial_number,
                                                'Исполнение': execution,
                                                'дата1': date_range[0].strftime("%d.%m.%Y"),
                                                'дата2': date_range[1].strftime("%d.%m.%Y"),
                                            })
                            
            else:
                #если несколько серийных номеров
                for serial_number in serial_numbers:
                    if serial_number not in added_serial_numbers:
                        added_serial_numbers.add(serial_number)
                        for date_range in date_ranges:
                            if self.check_and_add_year(serial_number, date_range[0].year, added_years_dict):
                                data_list.append({
                                            'заказчик': name_org,
                                            'Дата производства': data_prodaction,
                                            'серийный': serial_number,
                                            'Исполнение': execution,
                                            'дата1': date_range[0].strftime("%d.%m.%Y"),
                                            'дата2': date_range[1].strftime("%d.%m.%Y"),
                                        })
            
        return data_list
    
    def parse_name(self,soup):
        title_text = soup.title.text
        match = re.search("для .*? -", title_text)

        if match:
            name_org = match.group()[4:-1]
            return name_org

        return None
    
    def parse_ispolnenie(self, soup):
        title_text = soup.title.text
        if re.search(r'\b(lite|Lite|-Lite|-lite)\b', title_text):
            return str(2)
        return str(1)
    
    def parse_serial_number(self, soup):
        serial_number = soup.find('div', class_='cf_51').find('div', class_='value').getText()
        
        regul_dict = {
            'pattern1': re.compile(r'(025120102\.22\.(\d+)-025120102\.22\.(\d+))'),
            'pattern2': re.compile(r'(025120102\.21\.(\d+)-025120102\.21\.(\d+))'),
            'pattern3': re.compile(r'(025120102\.23\.(\d+)-025120102\.23\.(\d+))'),
            'pattern4': re.compile(r'(025120102\.24\.(\d+)-025120102\.24\.(\d+))'),
            'pattern5': re.compile(r'(025120101\.21\.(\d+)-025120101\.21\.(\d+))'),
            'pattern6': re.compile(r'(025120101\.22\.(\d+)-025120101\.22\.(\d+))'),
            'pattern7': re.compile(r'(025120101\.23\.(\d+)-025120101\.23\.(\d+))'),
            'pattern8': re.compile(r'(025120101\.24\.(\d+)-025120101\.24\.(\d+))'),
            'pattern9': re.compile(r'(\d+)-(\d+)'),
            'pattern10': re.compile(r'(ЭМ3752-(\d+)-ЭМ3752-(\d+))'),
            'pattern11': re.compile(r'(ЭМ3753-(\d+)-ЭМ3753-(\d+))'),
            'pattern12': re.compile(r'(ЭМ3754-(\d+)-ЭМ3754-(\d+))')
        }

        result_list = []

        for key, pattern in regul_dict.items():
            match = pattern.search(serial_number)
            if match:
                if match.group(2) is not None:
                    start_value, end_value = int(match.group(2)), int(match.group(3))
                    if start_value <= end_value:
                        result_list.extend([f"025120102.22.{number:04d}" for number in range(start_value, end_value + 1)])
                elif match.group(4) is not None:
                    result_list.extend([f"025120102.21.{number:04d}" for number in range(int(match.group(4)), int(match.group(5)) + 1)])
                elif match.group(6) is not None:
                    result_list.extend([f"025120102.23.{number:04d}" for number in range(int(match.group(6)), int(match.group(7)) + 1)])
                elif match.group(8) is not None:
                    result_list.extend([f"025120102.24.{number:04d}" for number in range(int(match.group(8)), int(match.group(9)) + 1)])
                elif match.group(10) is not None:
                    result_list.extend([f"025120101.21.{number:04d}" for number in range(int(match.group(10)), int(match.group(11)) + 1)])
                elif match.group(12) is not None:
                    result_list.extend([f"025120101.22.{number:04d}" for number in range(int(match.group(12)), int(match.group(13)) + 1)])
                elif match.group(14) is not None:
                    result_list.extend([f"025120101.23.{number:04d}" for number in range(int(match.group(14)), int(match.group(15)) + 1)])
                elif match.group(16) is not None:
                    result_list.extend([f"025120101.24.{number:04d}" for number in range(int(match.group(16)), int(match.group(17)) + 1)])
                elif match.group(18) is not None:
                    result_list.append("00000")
                elif match.group(19) is not None:
                    result_list.append(f"ЭМ3752-{int(match.group(19)):06d}")
                elif match.group(20) is not None:
                    result_list.append(f"ЭМ3753-{int(match.group(20)):06d}")
                elif match.group(21) is not None:
                    result_list.append(f"ЭМ3754-{int(match.group(21)):06d}")
                break
        else:
            parts = serial_number.split('-')
            if len(parts) == 2 and all(part.isdigit() for part in parts):
                start_value, end_value = map(int, parts)
                result_list.extend([f"{number:09d}" for number in range(start_value, end_value + 1)])
            else:
                return [serial_number]

        return result_list
   
    def parse_date(self,soup):
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
        date_end_str = soup.find('div', class_='cf_50').find('div', class_='value').getText()

        if date_end_str:
            date_start = datetime.strptime(date_start_str, "%d.%m.%Y")
            date_end = datetime.strptime(date_end_str, "%d.%m.%Y")

            # Проверка на разницу в годах
            year_diff = date_end.year - date_start.year

            if year_diff > 0:
                # Последовательность дат
                date_ranges = self.generate_date_range(date_start, date_end)
                return date_ranges
            elif year_diff == 0:
                # Если разница в годах = 0, то возвращаем диапазон
                return [(date_start, date_end)]
            else:
                # Ошибка если разница отрицательна
                raise ValueError("Ошибка в разнице годов. Дата окончания раньше даты начала.")
        else:
            # Если date_end_str пустая строка, вернуть None или выполнить другие действия по вашему усмотрению
            return []
 
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
    
    def check_and_add_year(self, serial_number, year, added_years_dict):
        if serial_number not in added_years_dict:
            added_years_dict[serial_number] = {year}
            return True
        elif year not in added_years_dict[serial_number]:
            added_years_dict[serial_number].add(year)
            return True
        else:
            return False
        
    def update_excel(self, data_list):
        try:
            df = pd.read_excel(self.excel_filename,dtype={'серийный': str})
        except FileNotFoundError:
            df = pd.DataFrame()

        for new_data in data_list:
            if_not_empty_df = not df.empty

            if if_not_empty_df:
                if_have_double = df.apply(lambda row: 
                    row['заказчик'] == new_data['заказчик'] and 
                    any(sn in row['серийный'] if isinstance(row['серийный'], list) else row['серийный'] == sn for sn in new_data['серийный']),
                    axis=1)

                if if_have_double.any():
                    new_rows = pd.DataFrame([{'заказчик': new_data['заказчик'],
                                            'серийный': sn} for sn in new_data['серийный']])
                    df = pd.concat([df, new_rows], ignore_index=True)
                else:
                    df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
            else:
                df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
        writer = pd.ExcelWriter(self.excel_filename, engine='xlsxwriter') 
        df.to_excel(writer, index=False, sheet_name='Лист1')
        
        text_format = writer.book.add_format({'num_format': '@'})

        worksheet = writer.sheets['Лист1']
        for idx, col in enumerate(df):
            series = df[col]
            max_len = max((series.astype(str).map(len).max(), len(str(col)))) + 2
            worksheet.set_column(idx, idx, max_len,text_format)

        writer.close()

    