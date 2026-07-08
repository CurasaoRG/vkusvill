import pytest
from pathlib import Path
from decimal import Decimal
from datetime import datetime
from vkusvill.parsers.pdf import PdfParser

# Путь к тестовому PDF-файлу
TEST_PDF_PATH = Path(__file__).parent / "files" / "pdf_test.pdf"

@pytest.fixture
def pdf_parser():
    """Фикстура для создания экземпляра PdfParser."""
    return PdfParser()

def test_parse_pdf(pdf_parser):
    """
    Тест парсинга PDF-чека.
    Проверяет, что данные извлекаются корректно.
    """
    # Чтение тестового PDF-файла
    with open(TEST_PDF_PATH, "rb") as f:
        raw_body = f.read()

    # Парсинг PDF
    check = pdf_parser.parse(raw_body)

    # Проверка заголовка чека
    assert check.msg_type == "pdf"
    assert check.address1 == "https://www.tinkoff.ru/gorod/grocery/vkusvill"
    assert check.address2 == "117342, Москва, ул. Бутлерова, 17Б"
    assert check.date == datetime(2024, 11, 19, 11, 26)
    assert check.cashier == "N/A"
    assert check.total == Decimal("1512.64")

    # Проверка товаров
    expected_items = [
        {
            "product_name": "Яйцо куриное С0",
            "price": Decimal("158.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("158.00"),
            "uom": "шт",
        },
        {
            "product_name": "Пергамент с силиконовым покрытием",
            "price": Decimal("135.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("135.00"),
            "uom": "шт",
        },
        {
            "product_name": "Сливки 10%, 450 мл",
            "price": Decimal("150.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("150.00"),
            "uom": "шт",
        },
        {
            "product_name": "Наполнитель для кошачьего туалета минеральный, 5 л",
            "price": Decimal("176.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("176.00"),
            "uom": "шт",
        },
        {
            "product_name": "Фольга для запекания 29см*10м",
            "price": Decimal("140.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("140.00"),
            "uom": "шт",
        },
        {
            "product_name": "Говядина для тушения, 400 г",
            "price": Decimal("356.80"),
            "qty": Decimal("1.00"),
            "amount": Decimal("356.80"),
            "uom": "шт",
        },
        {
            "product_name": "Огурцы гладкие",
            "price": Decimal("280.00"),
            "qty": Decimal("0.43"),
            "amount": Decimal("119.84"),
            "uom": "шт",
        },
        {
            "product_name": 'Хлеб бездрожжевой формовой славянский. Пекарня',
            "price": Decimal("85.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("85.00"),
            "uom": "шт",
        },
        {
            "product_name": 'Хлеб "Хуторской" с семечками на ржаном солоде, пекарня',
            "price": Decimal("150.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("150.00"),
            "uom": "шт",
        },
        {
            "product_name": "Зелень сушеная \"Розмарин\"",
            "price": Decimal("34.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("34.00"),
            "uom": "шт",
        },
        {
            "product_name": "Пакет-майка \"ВкусВилл\" малый",
            "price": Decimal("8.00"),
            "qty": Decimal("1.00"),
            "amount": Decimal("8.00"),
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