# vkusvill/parsers/ofd1.py
from __future__ import annotations

import re
from decimal import Decimal
from datetime import datetime

from bs4 import BeautifulSoup

from .base import BaseParser
from ..models import Check, Item


class Ofd1Parser(BaseParser):
    """
    Парсер для писем от echeck@1-ofd.ru (HTML, тип '1-ofd').
    Возвращает объект Check с полями и списком Item-ов.
    """

    def parse(self, raw_body: bytes | str) -> Check:
        html = raw_body.decode() if isinstance(raw_body, bytes) else raw_body
        soup = BeautifulSoup(html, "html.parser")
        strings = list(soup.stripped_strings)

        # ---------- Парсинг позиций ----------
        items: list[Item] = []
        goods_section = False
        row: list[str] = []
        k = 0

        for text in strings:
            if text == "№":
                goods_section = True
                continue
            if text == "АО \"Вкусвилл\"":
                goods_section = False
                continue

            if goods_section:
                if re.match(r"^\d+\.$", text):  # "1.", "2.", ...
                    if row:  # сохраняем предыдущий
                        items.append(self._build_item(row))
                    row = [text.rstrip(".")]
                    k = 1
                elif 0 < k < 5:
                    row.append(text)
                    k += 1
        if row:
            items.append(self._build_item(row))

        # ---------- Парсинг заголовка ----------
        header_map = self._extract_header(strings)
        return Check(
            msg_type="1-ofd",
            address1=header_map["address"],
            address2="",
            date=header_map["date"],
            cashier=header_map["cashier"],
            total=header_map["total"],
            items=items,
        )

    # ---------- Внутренние хелперы ----------
    @staticmethod
    def _build_item(row: list[str]) -> Item:
        try:
            description = ",".join(row[1].split(",")[:-1]).strip()
            price = Decimal(row[2].replace(",", "."))
            qty = Decimal(row[3].replace(",", "."))
            amount = Decimal(row[4].replace(",", "."))
            uom = row[1].split(",")[-1].strip()
        except (IndexError, ValueError):
            description, price, qty, amount, uom = "N/A", Decimal("0"), Decimal("0"), Decimal("0"), "шт"
        return Item(product_name=description, price=price, qty=qty, amount=amount, uom=uom)

    @staticmethod
    def _extract_header(strings: list[str]) -> dict:
        # Пример порядка строк:
        # 0: '№', 1: 'АО "Вкусвилл"', 2: 'Адрес: г. Москва, ул. Льва Толстого, 16',
        # 3: 'ИНН: ...', 4: 'Кассир: Иванов И.И.', 5: 'Дата: 19.07.2024 14:32',
        # 6: '...', 7: 'ИТОГО: 123,45'
        address = next(
            (s.replace("Адрес: ", "").strip() for s in strings if s.startswith("Адрес:")),
            "",
        )
        date_raw = next(
            (s.replace("Дата: ", "").strip() for s in strings if s.startswith("Дата:")),
            "",
        )
        cashier = next(
            (s.replace("Кассир: ", "").strip() for s in strings if s.startswith("Кассир:")),
            "",
        )
        total_raw = next(
            (s.replace("ИТОГО: ", "").strip() for s in strings if s.startswith("ИТОГО:")),
            "",
        )

        date_obj = datetime.strptime(date_raw, "%d.%m.%Y %H:%M")
        total = Decimal(total_raw.replace(",", "."))
        return {"address": address, "date": date_obj, "cashier": cashier, "total": total}