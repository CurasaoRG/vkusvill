from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
import logging

logging.basicConfig(level=logging.INFO)

@dataclass
class Item:
    """
    Представляет товар в чеке.
    
    Attributes:
        product_name (str): Название товара.
        price (Decimal): Цена за единицу.
        qty (Decimal): Количество.
        amount (Decimal): Общая сумма.
        uom (str): Единица измерения.
    """
    product_name: str = "N/A"
    price: Decimal = 0.0
    qty: Decimal = 0.0
    amount: Decimal = 0.0
    uom: str = "N/A"

    @staticmethod
    def from_csv_row(row: list[str]) -> Item:
        """
        Создает объект Item из строки CSV.
        """
        try:
            return Item(
                product_name=row[0],
                price=Decimal(row[1]),
                qty=Decimal(row[2]),
                amount=Decimal(row[3]),
                uom=row[4],
            )
        except (InvalidOperation, IndexError) as e:
            logging.error(f"Ошибка при создании Item из строки CSV: {row}. Ошибка: {e}")
            raise ValueError(f"Некорректные данные: {row}")

    def to_csv_row(self) -> list[str]:
        """Преобразует объект Item в строку CSV."""
        return [
            self.product_name,
            str(self.price),
            str(self.qty),
            str(self.amount),
            self.uom,
        ]

    def to_dict(self) -> dict:
        """Преобразует объект Item в словарь."""
        return {
            "product_name": self.product_name,
            "price": str(self.price),
            "qty": str(self.qty),
            "amount": str(self.amount),
            "uom": self.uom,
        }

    @staticmethod
    def from_dict(data: dict) -> Item:
        """Создает объект Item из словаря."""
        return Item(
            product_name=data.get("product_name", "N/A"),
            price = Decimal(data.get("price", 0.0)),
            qty = Decimal(data.get("qty", 0.0)),
            amount = Decimal(data.get("amount", 0.0)),
            uom=data.get("uom", "N/A"),
        )


@dataclass
class Check:
    """
    Представляет чек.
    
    Attributes:
        msg_type (str): Тип сообщения.
        address1 (str): Первый адрес.
        address2 (str): Второй адрес.
        date (datetime): Дата чека.
        cashier (str): Имя кассира.
        total (Decimal): Общая сумма.
        items (list[Item]): Список товаров.
    """
    msg_type: str
    address1: str
    address2: str
    date: datetime
    cashier: str
    total: Decimal
    items: list[Item]

    def check_info_row(self, check_id: int) -> list[str]:
        """Формирует строку CSV для информации о чеке."""
        return [
            str(check_id),
            self.msg_type,
            self.address1,
            self.address2,
            self.date.isoformat(timespec="minutes"),
            self.cashier,
            str(self.total),
        ]