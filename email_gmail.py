import imaplib
import email
from email.header import decode_header
import datetime
from bs4 import BeautifulSoup 
import csv
import re

CHECK_INFO_LOCATION = '/home/rg/Documents/Study/PET_projects/Vkusvill/check_info.csv'
DATA_LOCATION = '/home/rg/Documents/Study/PET_projects/Vkusvill/data.csv'

def get_increment(filename=r'/home/rg/Documents/Study/PET_projects/Vkusvill/increment.txt'):
    with open(filename, 'r') as f:
        num = f.readline()
        if num: return int(num)
        else: return 0
def set_increment(num=None, filename=r'/home/rg/Documents/Study/PET_projects/Vkusvill/increment.txt'):
    if not num:
        num = get_increment(filename) + 1
    with open(filename, 'w') as f:
        f.write(str(num))

def gmail_connect(username='', password=''):
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    try:
        imap.login(username, password)
    except imaplib.IMAP4.error:
        print("Ошибка входа. Проверьте имя пользователя и пароль.")
        return
    mailbox = "Money/&BCcENQQ6BDg-"
    imap.select(mailbox) 
    return imap

def get_msg(imap, num):
    imap.literal = u"ВКУСВИЛЛ".encode("utf-8")
    status, messages = imap.search('UTF-8', 'SUBJECT')
    email_ids = messages[0].split()
    if num > len(email_ids)-1: 
        return 'no data', None
    res, msg = imap.fetch(email_ids[num], "(RFC822)")
    # Получение содержимого письма
    for response_part in msg:
        if isinstance(response_part, tuple):
            # Парсинг письма
            msg = email.message_from_bytes(response_part[1])
            msg_from = re.search('^.*\<(\S*\@\S*)\>\s*$', msg['from']).group(1)
            if msg.is_multipart():
                for part in msg.walk():
                    if "text/" in part.get_content_type():
                        return msg_from, part.get_payload(decode=True)                       
            else:
                return msg_from, msg.get_payload(decode=True).decode()
def check_type(msg_from):
    if msg_from == 'echeck@1-ofd.ru': return '1-ofd'
    elif msg_from == 'noreply@ofd.ru': return 'ofd'
    elif msg_from == 'no data': return 'no data'
    else: return 'unknown'

def parse(msg_body, msg_type):
    if msg_body: soup = BeautifulSoup(msg_body, 'html.parser')
    else: return
    match msg_type:
        case 'echeck@1-ofd.ru': return parse_1_ofd(list(soup.stripped_strings))
        case 'noreply@ofd.ru': return parse_ofd(list(soup.stripped_strings))
        case _: return None


def parse_1_ofd(body_list):
    check_info = []
    items_data = []
    k = 0
    row = []
    goods, info = None, None
    for i, tag in enumerate(body_list):
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
            check_info.append(tag) 
        elif goods:
            if re.match('^\d+\.$', tag):
                k = 0
                row = []
                items_data.append(row)
            if k < 5: 
                row.append(tag)
                k+=1
    clean_items_data = []
    for line in items_data:
        description = ','.join(line[1].split(',')[:-1])
        price = float(line[2].replace(',','.'))
        qty = float(line[3].replace(',','.'))
        payment = float(line[4].replace(',','.'))
        unit_measure = line[1].split(',')[-1]
        clean_items_data.append([description, price, qty, payment, unit_measure])
    address = check_info[2]
    total = check_info[9]
    check_date = datetime.datetime.strptime(check_info[6], '%d.%m.%Y %H:%M')
    check_cashier = check_info[7]
    clean_check_info = [address, 'N/A', check_date, check_cashier, total]
    return {
            'check_info':clean_check_info,
            'items_data':clean_items_data
            }

def parse_ofd(body_list):
    check_info = []
    items_data = []
    info, goods = None, None
    last_item_field = 0
    k = 0
    row = []
    for i, tag in enumerate(body_list):
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
            check_info.append(tag)
        elif goods:
            if k<4: 
                row.append(tag)
                k+=1
            elif tag == 'Мера кол-ва предмета расчета': 
                last_item_field  = i
                k=0
            if i == last_item_field + 1: 
                k = 0
                items_data.append(row)
                row = []
    clean_items_data = []
    for line in items_data:
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
        clean_items_data.append([description, price, qty, payment, unit_measure])
    if not clean_items_data: clean_items_data = [None,None,None,None,None]
    address, check_date, check_cashier, total, check_num, shift_num, check_inn, check_rn, check_fd, check_fn, check_fpd  = None, None, None, None, None, None, None, None, None, None, None
    address_2 = 'N/A'
    position = -1
    field_name = None
    for i, item in enumerate(check_info):
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
    clean_check_info = [address, address_2, check_date, check_cashier, total]#, check_num, shift_num, check_inn, check_rn, check_fd, check_fn, check_fpd]
    return {
            'check_info':clean_check_info,
            'items_data':clean_items_data
            }
def write_csv(data, key, csv_location):
        if data:
            with open(csv_location, 'a', newline='') as csvfile:
                spamwriter = csv.writer(csvfile, delimiter=';',
                                    quotechar='|', quoting=csv.QUOTE_MINIMAL)
                if isinstance(data[0], list):
                    for row in data:
                        spamwriter.writerow(key + row)
                else:
                    spamwriter.writerow(key + data)

if __name__ == "__main__":
    imap = gmail_connect()
    for i in range(1000):
        latest_loaded_id = get_increment()
        msg_type, msg_body = get_msg(imap, latest_loaded_id)
        data = parse(msg_body, msg_type)
        if data:
            write_csv(data['check_info'], [latest_loaded_id, msg_type], CHECK_INFO_LOCATION)
            write_csv(data['items_data'], [latest_loaded_id, msg_type], DATA_LOCATION)
            print(f"Msg #{latest_loaded_id} was loaded. Item count: {len(data['items_data'])}. Check date is {data['check_info'][2]}")
            set_increment()
        else:
            print('No data. Break.')
            break     
    imap.close()
    imap.logout()