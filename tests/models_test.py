import unittest
from datetime import datetime
from decimal import Decimal
from ..vkusvill.models import Item, Check


class TestItem(unittest.TestCase):
    def setUp(self):
        self.valid_csv_row = ["Product 1", "10.50", "2", "21.00", "кг"]
        self.invalid_csv_row_price = ["Product 2", "invalid", "1", "10.00", "шт"]
        self.invalid_csv_row_missing = ["Product 3", "15.00"]
        
        self.valid_dict = {
            "product_name": "Product 1",
            "price": "10.50",
            "qty": "2",
            "amount": "21.00",
            "uom": "кг"
        }
        self.invalid_dict_missing = {
            "product_name": "Product 2",
            "price": "15.00",
            "qty": "1"
        }

    def test_from_csv_row_valid(self):
        """Тест создания Item из валидной CSV строки"""
        item = Item.from_csv_row(self.valid_csv_row)
        self.assertEqual(item.product_name, "Product 1")
        self.assertEqual(item.price, Decimal("10.50"))
        self.assertEqual(item.qty, Decimal("2"))
        self.assertEqual(item.amount, Decimal("21.00"))
        self.assertEqual(item.uom, "кг")

    def test_from_csv_row_invalid_price(self):
        """Тест создания Item с невалидной ценой"""
        with self.assertRaises(ValueError):
            Item.from_csv_row(self.invalid_csv_row_price)

    def test_from_csv_row_missing_fields(self):
        """Тест создания Item с недостающими полями"""
        with self.assertRaises(ValueError):
            Item.from_csv_row(self.invalid_csv_row_missing)

    def test_to_csv_row(self):
        """Тест преобразования Item в CSV строку"""
        item = Item.from_csv_row(self.valid_csv_row)
        csv_row = item.to_csv_row()
        self.assertEqual(csv_row, self.valid_csv_row)

    def test_to_dict(self):
        """Тест преобразования Item в словарь"""
        item = Item.from_csv_row(self.valid_csv_row)
        item_dict = item.to_dict()
        self.assertEqual(item_dict, self.valid_dict)

    def test_from_dict_valid(self):
        """Тест создания Item из валидного словаря"""
        item = Item.from_dict(self.valid_dict)
        self.assertEqual(item.product_name, "Product 1")
        self.assertEqual(item.price, Decimal("10.50"))
        self.assertEqual(item.qty, Decimal("2"))
        self.assertEqual(item.amount, Decimal("21.00"))
        self.assertEqual(item.uom, "кг")

    def test_from_dict_missing_fields(self):
        """Тест создания Item из словаря с недостающими полями"""
        item = Item.from_dict(self.invalid_dict_missing)
        self.assertEqual(item.product_name, "Product 2")
        self.assertEqual(item.price, Decimal("15.00"))
        self.assertEqual(item.qty, Decimal("1"))
        self.assertEqual(item.amount, Decimal("0.0"))  # Значение по умолчанию
        self.assertEqual(item.uom, "N/A")  # Значение по умолчанию


class TestCheck(unittest.TestCase):
    def setUp(self):
        self.items = [
            Item("Product 1", Decimal("10.50"), Decimal("2"), Decimal("21.00"), "кг"),
            Item("Product 2", Decimal("5.00"), Decimal("3"), Decimal("15.00"), "шт")
        ]
        self.check = Check(
            msg_type="SALE",
            address1="Store 1",
            address2="Moscow",
            date=datetime(2023, 1, 1, 12, 0),
            cashier="Ivanov",
            total=Decimal("36.00"),
            items=self.items
        )

    def test_check_info_row(self):
        """Тест формирования строки информации о чеке"""
        expected_row = [
            "1",
            "SALE",
            "Store 1",
            "Moscow",
            "2023-01-01T12:00",
            "Ivanov",
            "36.00"
        ]
        self.assertEqual(self.check.check_info_row(1), expected_row)

    def test_check_total_calculation(self):
        """Тест правильности расчета общей суммы чека"""
        # Проверяем, что сумма чека равна сумме всех items
        calculated_total = sum(item.amount for item in self.check.items)
        self.assertEqual(self.check.total, calculated_total)


if __name__ == "__main__":
    unittest.main()