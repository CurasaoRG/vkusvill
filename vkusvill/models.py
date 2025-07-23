from __future__ import annotations
from dataclasses import dataclass, asdict
from datetime import datetime
from decimal import Decimal

@dataclass
class Item:
    product_name: str
    price: Decimal
    qty: Decimal
    amount: Decimal
    uom: str

    @staticmethod
    def from_csv_row(row: list[str]) -> "Item":
        return Item(
            product_name=row[0],
            price=Decimal(row[1]),
            qty=Decimal(row[2]),
            amount=Decimal(row[3]),
            uom=row[4],
        )

    def to_csv_row(self) -> list[str]:
        return [
            self.product_name,
            str(self.price),
            str(self.qty),
            str(self.amount),
            self.uom,
        ]

@dataclass
class Check:
    msg_type: str
    address1: str
    address2: str
    date: datetime
    cashier: str
    total: Decimal
    items: list[Item]

    def check_info_row(self, check_id: int) -> list[str]:
        return [
            str(check_id),
            self.msg_type,
            self.address1,
            self.address2,
            self.date.isoformat(timespec="minutes"),
            self.cashier,
            str(self.total),
        ]