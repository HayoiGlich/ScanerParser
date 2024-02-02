from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import re
import pandas as pd
from datetime import datetime, timedelta, date
import calendar
import locale

class RubiScraper:
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
            name_org_front = self.parse_name_second(soup)
            serial_number = self.parse_serial_number(soup)
            ispolnenie = self.parse_ispolnenie(soup)
            ispolnenie2 = f"{name_org_front} {ispolnenie}"
            data_start = self.parse_date_start(soup)
            data_end = self.parse_date_end(soup)
            data_list.append([data_start,data_end])

        return data_list
    
    def parse_name(self, soup):
        title_text = soup.title.text
        match = re.search("для .*? -", title_text)

        if match:
            name_org = match.group()[4:-1].strip()
            return name_org

        return None
    
    def parse_ispolnenie(self, soup):
        ispolnenie = soup.find('div', class_='cf_40').find('div', class_='value').getText()
        pattern = re.compile(r'([A-ZА-Я]+\.\d+\.\d+-\d+)')
        match  = pattern.search(ispolnenie)
        if match:
            return match.group(1)
        else:
            return None

    def parse_name_second(self, soup):
            title_text = soup.title.text
            match = re.search(".*?для", title_text)

            if match:
                name_org = match.group()[20:-4].strip()
                if name_org == "Рубикон-А":
                    name_org = name_org.replace("-А", "")

                return name_org

            return None
    
    def parse_serial_number(self, soup):
        serial_number = soup.find('div', class_='cf_51').find('div', class_='value').getText()
        # регулярное выражение совмещающее патерн серийников
        pattern = re.compile(r'2-(\d+) - 2-(\d+)|(\d+)')
        matches = pattern.findall(serial_number)

        result_list = []

        for match in matches:
            if match[0]:  #если есть совпадение для паттерна "2-(\d+) - 2-(\d+)"
                serial_number1 = int(match[0])
                serial_number2 = int(match[1])
                result_list.extend([f"2-{number:04d}" for number in range(serial_number1, serial_number2 + 1)])
            elif match[2]:
                result_list.append(match[2])

        return result_list

    def parse_date_start(self, soup):
        locale.setlocale(locale.LC_TIME, '')  # Устанавливаем локаль для русского языка
        month_forms = {
            1: 'Января',
            2: 'Февраля',
            3: 'Марта',
            4: 'Апреля',
            5: 'Мая',
            6: 'Июня',
            7: 'Июля',
            8: 'Августа',
            9: 'Сентября',
            10: 'Октября',
            11: 'Ноября',
            12: 'Декабря'
        }
        
        # Получаем строку с датой
        data_start = soup.find('div', class_='cf_46').find('div', class_='value').getText()
        date_start = datetime.strptime(data_start, "%d.%m.%Y")

        # Получаем форму месяца из словаря
        month_name = month_forms.get(date_start.month, '')

        formatted_date = date_start.strftime("%d %B").replace(date_start.strftime("%B"), month_name)
        return formatted_date.capitalize()
    
    def parse_date_end(self, soup):
        locale.setlocale(locale.LC_TIME, '')  # Устанавливаем локаль для русского языка
        month_forms = {
            1: 'Января',
            2: 'Февраля',
            3: 'Марта',
            4: 'Апреля',
            5: 'Мая',
            6: 'Июня',
            7: 'Июля',
            8: 'Августа',
            9: 'Сентября',
            10: 'Октября',
            11: 'Ноября',
            12: 'Декабря'
        }
        
        # Получаем строку с датой
        date_end_str = soup.find('div', class_='cf_48').find('div', class_='value').getText()
        date_end = datetime.strptime(date_end_str, "%d.%m.%Y")

        # Получаем форму месяца из словаря
        month_name = month_forms.get(date_end.month, '')

        formatted_date = date_end.strftime("%d %B").replace(date_end.strftime("%B"), month_name)
        return formatted_date.capitalize()
    
if __name__ == "__main__":
    chromedriver_path = "chromedriver.exe"  
    username = "a.atochkin"
    password = "Ha0Pyb2Xow2"
    excel_filename = "Тест2 рубик.xlsx"
    
    
    scraper = RubiScraper(chromedriver_path, username, password, excel_filename)

    scraper.login()

    urls_to_scrape = ['https://tasks.etecs.ru/issues/20741','https://tasks.etecs.ru/issues/20537', 'https://tasks.etecs.ru/issues/20740']

    data_list = scraper.scrape_data(urls_to_scrape)

    print("ДАнные:")
    for data in data_list:
        print(data)


    '''def update_excel(self, data_list):
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

            writer.close()'''