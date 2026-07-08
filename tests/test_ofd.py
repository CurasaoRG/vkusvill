import pytest
from pathlib import Path
from email import policy
from email.parser import BytesParser
from vkusvill.parsers.ofd import OfdParser
from datetime import datetime 
from decimal import Decimal

# Путь к тестовому файлу
TEST_EML_PATH = Path(__file__).parent / "files" / "ofd_test.eml"

@pytest.fixture
def raw_email():
    """Фикстура для чтения содержимого тестового файла."""
    with open(TEST_EML_PATH, "rb") as f:
        return f.read()

@pytest.fixture
def decoded_body(raw_email):
    """Фикстура для декодирования тела письма."""
    # Парсим письмо с помощью email
    msg = BytesParser(policy=policy.default).parsebytes(raw_email)
    
    # Ищем текстовую часть письма
    body = None
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            charset = part.get_content_charset() or "utf-8"
            body = part.get_payload(decode=True).decode(charset)
            break
    
    if not body:
        raise ValueError("Не удалось найти HTML-часть письма")
    
    return body

@pytest.fixture
def parser():
    """Фикстура для создания экземпляра OfdParser."""
    return OfdParser()

def test_parse_ofd(parser, decoded_body):
    """
    Тест парсинга письма от OFD.
    Проверяет корректность извлечения данных о чеке и товарах.
    """
    # Парсинг данных
    check = parser.parse(decoded_body)

    # Проверка заголовка чека
    assert check.msg_type == "ofd"
    assert check.address1 == "www.vkusvill.ru"
    # assert check.address2 == "109316, г. Москва, вн.тер.г. муниципальный округ Текстильщики, проезд Остаповский, д. 22 стр. 16"
    assert check.date == datetime(2025, 7, 27, 11, 6)
    assert check.cashier == "Кассир"
    assert check.total == Decimal("3633.23")

    # Проверка товаров
    expected_items = [
        {
            "product_name": "Йогурт греческий из цельного молока",
            "price": Decimal("50.94"),
            "qty": Decimal("1.00"),
            "amount": Decimal("50.94"),
            "uom": "шт",
        },
        {
            "product_name": "Кинза, 50 г",
            "price": Decimal("90.67"),
            "qty": Decimal("1.00"),
            "amount": Decimal("90.67"),
            "uom": "шт",
        },        
        {
            "product_name": "Айран 2%",
            "price": Decimal("77.71"),
            "qty": Decimal("1.00"),
            "amount": Decimal("77.71"),
            "uom": "шт",
        },
        {
            "product_name": "Сельдерей черешковый стебли, 250 г",
            "price": Decimal("160.61"),
            "qty": Decimal("1.00"),
            "amount": Decimal("160.61"),
            "uom": "шт",
        },
        {
            "product_name": "Облепиха дикорастущая зам, 300 г",
            "price": Decimal("172.71"),
            "qty": Decimal("2.00"),
            "amount": Decimal("345.42"),
            "uom": "шт",
        },
        {
            "product_name": "Хлеб ржаной 'Родной из детства', половинка. Пекарня",
            "price": Decimal("98.44"),
            "qty": Decimal("1.00"),
            "amount": Decimal("98.44"),
            "uom": "шт",
        },        
        {
            "product_name": "Йогурт греческий из цельного молока",
            "price": Decimal("50.94"),
            "qty": Decimal("1.00"),
            "amount": Decimal("50.94"),
            "uom": "шт",
        },        
        {
            "product_name": "Петрушка, 50 г",
            "price": Decimal("63.03"),
            "qty": Decimal("1.00"),
            "amount": Decimal("63.03"),
            "uom": "шт",
        },        
        {
            "product_name": "Сырок творожный с лесными ягодами и злаками 15%",
            "price": Decimal("51.80"),
            "qty": Decimal("1.00"),
            "amount": Decimal("51.80"),
            "uom": "шт",
        },        
        {
            "product_name": "Квас 'Овсяный' нефильтр. непастеризованный, 1 л",
            "price": Decimal("142.48"),
            "qty": Decimal("1.00"),
            "amount": Decimal("142.48"),
            "uom": "шт",
        },
        {
            "product_name": "Сырок творожный с лесными ягодами и злаками 15%",
            "price": Decimal("51.80"),
            "qty": Decimal("1.00"),
            "amount": Decimal("51.80"),
            "uom": "шт",
        },
        {
            "product_name": "Молоко 3,2% в бутылке, 900 мл",
            "price": Decimal("88.07"),
            "qty": Decimal("1.00"),
            "amount": Decimal("88.07"),
            "uom": "шт",
        }, 
        {
            "product_name": "Йогурт греческий из цельного молока",
            "price": Decimal("50.94"),
            "qty": Decimal("1.00"),
            "amount": Decimal("50.94"),
            "uom": "шт",
        }, 
        {
            "product_name": "Чеснок,100 г",
            "price": Decimal("89.81"),
            "qty": Decimal("1.00"),
            "amount": Decimal("89.81"),
            "uom": "шт",
        }, 
        {
            "product_name": "Йогурт греческий из цельного молока",
            "price": Decimal("50.94"),
            "qty": Decimal("1.00"),
            "amount": Decimal("50.94"),
            "uom": "шт",
        }, 
        {
            "product_name": "Сливки 10%, 450 мл",
            "price": Decimal("138.16"),
            "qty": Decimal("1.00"),
            "amount": Decimal("138.16"),
            "uom": "шт",
        }, 
        {
            "product_name": "Диски ватные, 120 шт",
            "price": Decimal("105.35"),
            "qty": Decimal("1.00"),
            "amount": Decimal("105.35"),
            "uom": "шт",
        }, 
        {
            "product_name": "Пакет-майка 'ВкусВилл' малый",
            "price": Decimal("7.09"),
            "qty": Decimal("1.00"),
            "amount": Decimal("7.09"),
            "uom": "шт",
        },     
        {
            "product_name": "Сервелат Зернистый варено-копченый",
            "price": Decimal("373.05"),
            "qty": Decimal("1.00"),
            "amount": Decimal("373.05"),
            "uom": "шт",
        },     
        {
            "product_name": "Укроп, 100 г",
            "price": Decimal("80.30"),
            "qty": Decimal("1.00"),
            "amount": Decimal("80.30"),
            "uom": "шт",
        },
        {
            "product_name": "Кефир 2,5% в бутылке, 900 г",
            "price": Decimal("93.26"),
            "qty": Decimal("1.00"),
            "amount": Decimal("93.26"),
            "uom": "шт",
        },     
        {
            "product_name": "Изделие х.б 'Чиабатта', Пекарня",
            "price": Decimal("90.67"),
            "qty": Decimal("1.00"),
            "amount": Decimal("90.67"),
            "uom": "шт",
        },     
        {
            "product_name": "Лук репчатый",
            "price": Decimal("64.75"),
            "qty": Decimal("0.726"),
            "amount": Decimal("47.01"),
            "uom": "кг",
        },     
        {
            "product_name": "Огурцы короткоплодные",
            "price": Decimal("155.43"),
            "qty": Decimal("0.908"),
            "amount": Decimal("141.13"),
            "uom": "кг",
        },     
        {
            "product_name": "Шейка свиная без кости",
            "price": Decimal("544.03"),
            "qty": Decimal("0.979"),
            "amount": Decimal("532.61"),
            "uom": "кг",
        },     
        {
            "product_name": "Морковь мытая",
            "price": Decimal("130.39"),
            "qty": Decimal("1.115"),
            "amount": Decimal("145.38"),
            "uom": "кг",
        },     
        {
            "product_name": "Свекла молодая",
            "price": Decimal("63.03"),
            "qty": Decimal("0.864"),
            "amount": Decimal("54.46"),
            "uom": "кг",
        },     
        {
            "product_name": "Картофель молодой Египет",
            "price": Decimal("86.34"),
            "qty": Decimal("2.516"),
            "amount": Decimal("217.23"),
            "uom": "кг",
        },     
        {
            "product_name": "Картофель молодой",
            "price": Decimal("58.71"),
            "qty": Decimal("2.096"),
            "amount": Decimal("123.06"),
            "uom": "кг",
        },     
        {
            "product_name": "Пакет-майка 'ВкусВилл' малый",
            "price": Decimal("6.95"),
            "qty": Decimal("1.00"),
            "amount": Decimal("6.95"),
            "uom": "шт",
        },     
                {
            "product_name": "Пакет-майка 'ВкусВилл' малый",
            "price": Decimal("6.96"),
            "qty": Decimal("2.00"),
            "amount": Decimal("13.92"),
            "uom": "шт",
        },
    ]

    # Сравнение товаров
    for idx, item in enumerate(check.items):
        assert item.product_name == expected_items[idx]["product_name"]
        assert item.price == expected_items[idx]["price"]
        assert item.qty == expected_items[idx]["qty"]
        assert item.amount == expected_items[idx]["amount"]
        assert item.uom == expected_items[idx]["uom"]

    # Дополнительные проверки
    assert len(check.items) == len(expected_items)