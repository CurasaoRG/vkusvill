from __future__ import annotations

import re
from decimal import Decimal
from datetime import datetime

from bs4 import BeautifulSoup

from src.parsers.base import BaseParser
from src.models import Check, Item
from src.utils import parse_decimal


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
        header_strings: list[str] = []
        items: list[Item] = []
        goods_section = False
        info_section = None
        row: list[str] = []
        k = 0

        for text in strings:
            if text == "№":
                info_section = False
                goods_section = True
                continue
            if text.startswith("ИНН"):
                info_section = True
                goods_section = False
                header_strings.append(prev_text)
                header_strings.append(text)
                continue
            if text == 'ИТОГО:': 
                info_section = True
                goods_section = False
            if info_section:
                header_strings.append(text) 
            elif goods_section:
                if re.match(r"^\d+\.$", text):  
                    if row:
                        items.append(self._build_item(row))
                    row = [text.rstrip(".")]
                    k = 1
                elif 0 < k < 5:
                    row.append(text)
                    k += 1
            prev_text = text[:]
        if row:
            items.append(self._build_item(row))

        # ---------- Парсинг заголовка ----------
        header_map = self._extract_header(header_strings)
        return Check(
            msg_type="1-ofd",
            company=header_map["company"],
            address1=header_map["address1"],
            address2=header_map["address2"],
            date=header_map["date"],
            cashier=header_map["cashier"],
            total=header_map["total"],
            items=items,
        )

    # ---------- Внутренние хелперы ----------
    @staticmethod
    def _build_item(row: list[str]) -> Item:
        try:
            description = row[1].strip()
            price = parse_decimal(row[2])
            qty = parse_decimal(row[3])
            amount = parse_decimal(row[4])
            uom = ''
        except (IndexError, ValueError):
            description, price, qty, amount, uom = "N/A", Decimal("0"), Decimal("0"), Decimal("0"), "шт"
        return Item(product_name=description, price=price, qty=qty, amount=amount, uom=uom)

    @staticmethod
    def _extract_header(strings: list[str]) -> dict:
        company = strings[0]
        inn = strings[1].split(':')[-1].strip()
        address1 = strings[2]
        address2 = strings[3]
        total = parse_decimal(strings[11])
        check_date = datetime.strptime(strings[7], '%d.%m.%Y %H:%M')
        cashier = strings[8].split(':')[-1].strip()
        return {
            "company":company,
            "address1": address1, 
            "address2": address2, 
            "date": check_date, 
            "cashier": cashier, 
            "total": total}