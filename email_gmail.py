import imaplib
import email
import datetime
from bs4 import BeautifulSoup 
import csv
import re
from dotenv import dotenv_values
from pypdf import PdfReader
from unicodedata import normalize
import io
import pymorphy2
import locale

CSV_LOCATIONS = {'check_info':'/home/rg/Documents/Study/PET_projects/Vkusvill/check_info.csv',
                 'items_data': '/home/rg/Documents/Study/PET_projects/Vkusvill/data.csv'}
INCREMENT_FILE = '/home/rg/Documents/Study/PET_projects/Vkusvill/increment.txt'


def get_increment(filename=INCREMENT_FILE):
    with open(filename, 'r') as f:
        num = f.readline()
        if num: return int(num)
        else: return 0
def set_increment(num=None, filename=INCREMENT_FILE):
    if not num:
        num = get_increment(filename) + 1
    with open(filename, 'w') as f:
        f.write(str(num))

class IMAPHandler:
    def __init__(self, username, password, mailbox):
        self.mailbox = mailbox
        self.imap = imaplib.IMAP4_SSL("imap.gmail.com")
        try:
            self.imap.login(username, password)
        except imaplib.IMAP4.error:
            print("Ошибка входа. Проверьте имя пользователя и пароль.")
            raise
        self.imap.select(self.mailbox)

    def get_message(self, num):
        """Получить письмо по номеру."""
        self.imap.literal = u"ВКУСВИЛЛ".encode("utf-8")
        status, messages = self.imap.search('UTF-8', 'OR (FROM "noreply-cloudkassir@cp.ru") SUBJECT')
        email_ids = messages[0].split()
        if num > len(email_ids) - 1:
            return None  # Возвращаем None, если письма нет
        res, msg = self.imap.fetch(email_ids[num], "(RFC822)")
        for response_part in msg:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                return msg
        return None

    def close(self):
        """Закрыть соединение."""
        self.imap.logout()

class Message:
    def __init__(self, username, password, mailbox):
        self.imap_handler = IMAPHandler(username, password, mailbox)
        self.msg_types = {
            'noreply@ofd.ru': Check_ofd,
            'echeck@1-ofd.ru': Check_1_ofd,
            'noreply-cloudkassir@cp.ru': CheckPDF
        }

    def get_msg(self, num):
        """Получить и обработать письмо."""
        msg = self.imap_handler.get_message(num)
        if not msg:
            return Check()

        msg_from = re.search(r"[\w.-]+@[\w.-]+", msg['from']).group(0)
        msg_body = None

        if msg.is_multipart():
            for part in msg.walk():
                if "text/" in part.get_content_type() and not msg_body:
                    msg_body = part.get_payload(decode=True)
                if "application/pdf" in part.get_content_type():
                    msg_body = part.get_payload(decode=True)
        else:
            msg_body = msg.get_payload(decode=True).decode()

        return self.msg_types[msg_from](msg_body)

class Check:
    HEADERS = {
        'items_data': ['id', 'msg_type', 'product_name', 'price', 'qty', 'amount', 'uom'],
        'check_info':['id', 'msg_type', 'address1', 'address2', 'date', 'cashier', 'total']
        }
    CSV_PARAMS = {"delimiter":";",
                  "quotechar":"|", 
                  "quoting":csv.QUOTE_MINIMAL}
    # Константа для индекса даты в check_info
    CHECK_DATE_INDEX = 2

    def __init__(self, msg_type='no data'):
        self.msg_type = msg_type
        self.check_info = []
        self.items_data = []
        self.parsed = False
    @staticmethod 
    def parse_russian_datetime(date_string):
        """
        Функция для парсинга дат в формате "дд месяц гггг г. в чч:мм"
        Поддерживает разные падежи месяцев
        
        Примеры входных строк:
        - "06 марта 2025 г. в 13:59"
        - "19 ноябрь 2024 г. в 11:26"
        """
        locale.setlocale(locale.LC_TIME, 'ru_RU.UTF-8')
        morph = pymorphy2.MorphAnalyzer()
        # Регулярное выражение для извлечения компонентов даты
        pattern = r'(\d{1,2})\s+([а-яА-Я]+)\s+(\d{4})\s+г\.\s+в\s+(\d{1,2}:\d{2})'
        match = re.match(pattern, date_string)
        if not match:
            raise ValueError("Неверный формат даты")
        day, month_ru, year, time = match.groups()
        # Приводим месяц к родительному падежу
        month = morph.parse(month_ru)[0].inflect({'gent'}).word.title()
        # Формируем строку для парсинга
        date_str = f"{day} {month} {year} {time}"
        # Парсим в datetime объект
        try:
            dt = datetime.datetime.strptime(date_str, '%d %B %Y %H:%M')
        except ValueError as e:
            raise ValueError(f"Ошибка при парсинге даты: {e}")
        return dt


    def parse(self):
        pass

    def write_to_csv(self, data_type, key, csv_location, headers_required=False):
        # to fix: add error handling
        data = self.__getattribute__(data_type)
        if data:
            with open(csv_location, 'a', newline='') as csvfile:
                spamwriter = csv.writer(csvfile, **self.CSV_PARAMS)
                if headers_required:
                    spamwriter.writerow(self.HEADERS[data_type])
                if isinstance(data[0], list):
                    for row in data:
                        spamwriter.writerow(key + row)
                else:
                    spamwriter.writerow(key + data)

    def print_status(self, msg_num):
        print(f"Msg #{msg_num} was loaded. Type: {self.msg_type}. Item count: {len(self.items_data)}. Check date is {self.check_info[2]}")

    
class Check_ofd(Check):
    def __init__(self, msg_body):
        super().__init__(msg_type = 'ofd')
        self.body_list = list(BeautifulSoup(msg_body, 'html.parser').stripped_strings)

    def parse(self):
        raw_check_info = []
        raw_items_data = []
        info, goods = None, None
        last_item_field = 0
        k = 0
        row = []
        for i, tag in enumerate(self.body_list):
            match tag:
                case 'Кассовый чек / Приход':
                    info = True
                    continue
                case 'check.ofd.ru':
                    info = False
                    goods = True
                    continue
                case 'ИТОГ': 
                    info = True
                    goods = False
            if info:
                raw_check_info.append(tag)
            elif goods:
                if k<4: 
                    row.append(tag)
                    k+=1
                elif tag == 'Мера кол-ва предмета расчета': 
                    last_item_field  = i
                    k=0
                if i == last_item_field + 1: 
                    k = 0
                    raw_items_data.append(row)
                    row = []
        for line in raw_items_data:
            try:
                if 'X' in line[0]:
                    description = 'N/A'
                    price = float(line[0].split(' X ')[1])
                    qty = float(line[0].replace(',','.').split('X')[0])
                else:
                    description = line[0]
                    price = float(line[1].split(' X ')[1])
                    qty = float(line[1].replace(',','.').split('X')[0])
                payment = float(line[2].split('=')[1])
                unit_measure = line[4].split('.')[0]
            except IndexError:
                description, price, qty, payment, unit_measure = 'N/A', 'N/A', 'N/A', 'N/A', 'N/A'
            self.items_data.append([description, price, qty, payment, unit_measure])
        if not self.items_data: self.items_data = [None,None,None,None,None]
        address, check_date, check_cashier, total, check_num, shift_num, check_inn, check_rn, check_fd, check_fn, check_fpd  = None, None, None, None, None, None, None, None, None, None, None
        address_2 = 'N/A'
        position = -1
        field_name = None
        for i, item in enumerate(raw_check_info):
            if item in ('#', 'НОМЕР СМЕНЫ', 'МЕСТО РАСЧЁТОВ', 'АДРЕС РАСЧЁТОВ', 'ДАТА ВЫДАЧИ', 'ДОКУМЕНТ В СМЕНЕ', 'КАССИР','ИТОГ'):
                field_name = item
                position = i+1
            elif re.fullmatch('^.+ \d+$', item, flags=0):
                field_name, value = item.split(' ')[:2]
            if i == position:
                match field_name:
                    case '#': check_fd = item
                    case 'НОМЕР СМЕНЫ': shift_num = item
                    case 'МЕСТО РАСЧЁТОВ': address = item
                    case 'АДРЕС РАСЧЁТОВ': address_2 = item
                    case 'ДАТА ВЫДАЧИ': check_date = datetime.datetime.strptime(item, '%d.%m.%y %H:%M')
                    case 'ДОКУМЕНТ В СМЕНЕ': check_num = item
                    case 'КАССИР': check_cashier = item
                    case 'РН': check_rn = value
                    case 'ИНН': check_inn = value
                    case 'ФН': check_fn = value
                    case 'ФПД': check_fpd = value
                    case 'ИТОГ': total = item
        self.check_info = [address, address_2, check_date, check_cashier, total]
        self.parsed = True

class Check_1_ofd(Check):
    def __init__(self, msg_body):
        super().__init__(msg_type='1-ofd')
        self.body_list = list(BeautifulSoup(msg_body, 'html.parser').stripped_strings)
        
    def parse(self):
        k = 0
        row = []
        goods, info = None, None
        for i, tag in enumerate(self.body_list):
            match tag:
                case '№':
                    goods = True
                    info = False
                    continue
                case 'АО "Вкусвилл"': 
                    info = True
                    continue
                case 'ИТОГО:': 
                    info = True
                    goods = False
            if info:
                self.check_info.append(tag) 
            elif goods:
                if re.match('^\d+\.$', tag):
                    k = 0
                    row = []
                    self.items_data.append(row)
                if k < 5: 
                    row.append(tag)
                    k+=1
        clean_items_data = []
        for line in self.items_data:
            description = ','.join(line[1].split(',')[:-1])
            price = float(line[2].replace(',','.'))
            qty = float(line[3].replace(',','.'))
            payment = float(line[4].replace(',','.'))
            unit_measure = line[1].split(',')[-1]
            clean_items_data.append([description, price, qty, payment, unit_measure])
        address = self.check_info[2]
        total = self.check_info[9]
        check_date = datetime.datetime.strptime(self.check_info[6], '%d.%m.%Y %H:%M')
        check_cashier = self.check_info[7]
        clean_check_info = [address, 'N/A', check_date, check_cashier, total]
        self.check_info = clean_check_info[:]
        self.items_data = clean_items_data[:]
        self.parsed = True

class CheckPDF(Check):
    def __init__(self, msg_body):
        super().__init__(msg_type='pdf')
        # Создаем байтовый поток вместо записи на диск
        self.reader = PdfReader(io.BytesIO(msg_body))

    def parse(self):
        for page in self.reader.pages:
            text = page.extract_text(extraction_mode='layout')
            # Регулярное выражение для извлечения полей и значений
            pattern_header = re.compile(r"(Дата выдачи|Место осуществления расчета|Адрес осуществления расчетов|ИТОГ)\s+(.*)")
            pattern_fields = re.compile(r"(?:\d+ +)(.+?)\s+(\d+,\d+)\s+(\d+,\d+)\s+(\d+,\d+)", re.MULTILINE)
            # Поиск совпадений
            matches_header = re.findall(pattern_header, text)
            matches_fields = re.findall(pattern_fields, text)
            for field, value in matches_header:
                match field:
                    case 'Дата выдачи': check_date = Check.parse_russian_datetime(value.strip())
                    case 'Место осуществления расчета': address = value.strip()
                    case 'Адрес осуществления расчетов': address_2 = value.strip()
                    case 'ИТОГ': total = normalize('NFKD', value.strip()).replace(',','.').replace(' ','')
            for product_name, price, quantity, amount in matches_fields:
                self.items_data.append([product_name.strip(), price.strip().replace(',','.'), quantity.strip().replace(',','.'), amount.strip().replace(',','.'), 'N/A'])
        else:
            self.check_info = [address, address_2, check_date, 'N/A', total]
            self.parsed = True

if __name__ == "__main__":
    config = dotenv_values('Vkusvill/.env')
    msg = Message(
        username=config['GMAIL_USERNAME'],
        password=config['GMAIL_PASSWORD'],
        mailbox=config['MAILBOX']
    )
    try:
        for i in range(100):
            latest_loaded_id = get_increment()
            new_check = msg.get_msg(latest_loaded_id)
            new_check.parse()
            if new_check.parsed:
                for data_type in ('check_info', 'items_data'):
                    new_check.write_to_csv(
                        data_type=data_type,
                        key=[latest_loaded_id, new_check.msg_type],
                        csv_location=CSV_LOCATIONS[data_type],
                        headers_required=latest_loaded_id == 0
                    )
                new_check.print_status(latest_loaded_id)
                set_increment()
            else:
                print('No data. Break.')
                break
    finally:
        msg.imap_handler.close()  # Закрыть соединение с IMAP
