import imaplib
import email
import datetime
from bs4 import BeautifulSoup 
import csv
import re
from dotenv import dotenv_values

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

class Message(imaplib.IMAP4_SSL):
    def __init__(self, username, password, mailbox):
        self.mailbox = mailbox
        super().__init__("imap.gmail.com")
        try:
            self.login(username, password)
        except imaplib.IMAP4.error:
            print("Ошибка входа. Проверьте имя пользователя и пароль.")
            return
        self.select(self.mailbox)
        self.msg_types = {'noreply@ofd.ru':Check_ofd, 'echeck@1-ofd.ru':Check_1_ofd}

    def get_msg(self, num):
        self.literal = u"ВКУСВИЛЛ".encode("utf-8")
        status, messages = self.search('UTF-8', 'SUBJECT')
        email_ids = messages[0].split()
        if num > len(email_ids)-1: 
            return Check()
        res, msg = self.fetch(email_ids[num], "(RFC822)")
        # Получение содержимого письма
        for response_part in msg:
            if isinstance(response_part, tuple):
                # Парсинг письма
                msg = email.message_from_bytes(response_part[1])
                msg_from = re.search('^.*\<(\S*\@\S*)\>\s*$', msg['from']).group(1)
                if msg.is_multipart():
                    for part in msg.walk():
                        if "text/" in part.get_content_type():
                            return self.msg_types[msg_from](msg_type=msg_from, msg_body=part.get_payload(decode=True))
                else:
                    return self.msg_types[msg_from](msg_type=msg_from, msg_body=msg.get_payload(decode=True).decode())

class Check:
    HEADERS = {
        'items_data': ['id', 'msg_type', 'product_name', 'price', 'qty', 'amount', 'uom'],
        'check_info':['id', 'msg_type', 'address1', 'address2', 'date', 'cashier', 'total']
        }
    CSV_PARAMS = {"delimiter":";",
                  "quotechar":"|", 
                  "quoting":csv.QUOTE_MINIMAL}
    def __init__(self, msg_type='no data', msg_body=''):
        self.msg_type = msg_type
        self.body_list = list(BeautifulSoup(msg_body, 'html.parser').stripped_strings)
        self.check_info = []
        self.items_data = []
        self.parsed = False

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
        print(f"Msg #{msg_num} was loaded. Item count: {len(self.items_data)}. Check date is {self.check_info[2]}")


    
class Check_ofd(Check):
    def _init__(self, msg_body):
        super().__init__(msg_type = 'ofd', msg_body = msg_body)

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
    def _init__(self, msg_body):
        super().__init__(msg_type='1-ofd', msg_body=msg_body)

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

if __name__ == "__main__":
    config = dotenv_values('Vkusvill/.env')
    msg = Message(username = config['GMAIL_USERNAME'],
                  password=config['GMAIL_PASSWORD'],
                  mailbox=config['MAILBOX'])
    for i in range(1000):
        latest_loaded_id = get_increment()
        new_check = msg.get_msg(latest_loaded_id)            
        new_check.parse()
        if new_check.parsed:
            for data_type in ('check_info', 'items_data'):
                new_check.write_to_csv(data_type=data_type,
                                       key=[latest_loaded_id, new_check.msg_type],
                                       csv_location=CSV_LOCATIONS[data_type],
                                       headers_required=latest_loaded_id>0
                )
            new_check.print_status(latest_loaded_id)
            set_increment()
        else:
            print('No data. Break.')
            break