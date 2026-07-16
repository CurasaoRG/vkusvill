from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from src.models import Item
import logging

logging.basicConfig(level=logging.INFO)

@dataclass
class Check:
    """
    Представляет чек.
    
    Attributes:
        msg_type (str): Тип сообщения.
        company (str): Магазин
        address1 (str): Первый адрес.
        address2 (str): Второй адрес.
        date (datetime): Дата чека.
        cashier (str): Имя кассира.
        total (Decimal): Общая сумма.
        items (list[Item]): Список товаров.
    """
    msg_type: str
    company: str
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
            self.company,
            self.address1,
            self.address2,
            self.date.isoformat(timespec="minutes"),
            self.cashier,
            str(self.total),
        ]