import pytest
from pathlib import Path
from decimal import Decimal
from vkusvill.parsers.ofd1 import Ofd1Parser
from datetime import datetime
from quopri import decodestring

# Путь к тестовому файлу
TEST_EML_PATH = Path(__file__).parent / "files" / "ofd1_test_part.eml"

@pytest.fixture
def raw_body():
    """Фикстура для чтения содержимого тестового файла."""
    with open(TEST_EML_PATH, "rb") as f:
        return f.read()

@pytest.fixture
def parser():
    """Фикстура для создания экземпляра Ofd1Parser."""
    return Ofd1Parser()

def test_parse_ofd1(parser, raw_body):
    """
    Тест парсинга письма от OFD1.
    Проверяет корректность извлечения данных о чеке и товарах.
    """
    # Парсинг данных
    check = parser.parse(decodestring(raw_body))

    # Проверка заголовка чека
    assert check.msg_type == "1-ofd"
    assert check.address1 == "115569, г Москва, Орехово-Борисово Северное р-н, ул Маршала Захарова, д 2"
    assert check.address2 == ""
    assert check.date == datetime(2025, 7, 25, 18, 43)
    assert check.cashier == "АУР"
    assert check.total == Decimal("1155.18")

    # Проверка товаров
    expected_items = [
        {
            "product_name": "Гедза со свининой, зам.",
            "price": Decimal("370.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("370.00"),
            "uom": "шт",
        },
        {
            "product_name": "[M] Сметана 15%, 250 г",
            "price": Decimal("106.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("106.00"),
            "uom": "шт",
        },
        {
            "product_name": "[M] Сливки ультрапастеризованные 10%, 208 г",
            "price": Decimal("81.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("81.00"),
            "uom": "шт",
        },
        {
            "product_name": "Хлеб \"Хуторской\" с семечками на ржаном солоде, пекарня",
            "price": Decimal("150.00"),
            "qty": Decimal("2.00"),
            "amount": Decimal("300.00"),
            "uom": "шт",
        },
        {
            "product_name": "Цукини",
            "price": Decimal("175.00"),
            "qty": Decimal("0.755"),
            "amount": Decimal("132.13"),
            "uom": "кг",
        },
        {
            "product_name": "Нектарин желтый Узбекистан",
            "price": Decimal("410.00"),
            "qty": Decimal("0.405"),
            "amount": Decimal("166.05"),
            "uom": "кг",
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