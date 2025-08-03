# vkusvill/parsers/ofd1.py
from __future__ import annotations

import re
from decimal import Decimal
from datetime import datetime

from bs4 import BeautifulSoup

from .base import BaseParser
from ..models import Check, Item
from ..utils import parse_decimal


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
            if text == "АО \"Вкусвилл\"":
                info_section = True
                goods_section = False
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
        if row:
            items.append(self._build_item(row))

        # ---------- Парсинг заголовка ----------
        header_map = self._extract_header(header_strings)
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
            price = parse_decimal(row[2])
            qty = parse_decimal(row[3])
            amount = parse_decimal(row[4])
            uom = row[1].split(",")[-1].strip()
        except (IndexError, ValueError):
            description, price, qty, amount, uom = "N/A", Decimal("0"), Decimal("0"), Decimal("0"), "шт"
        return Item(product_name=description, price=price, qty=qty, amount=amount, uom=uom)

    @staticmethod
    def _extract_header(strings: list[str]) -> dict:
        address = strings[2]
        total = parse_decimal(strings[9])
        check_date = datetime.strptime(strings[6], '%d.%m.%Y %H:%M')
        cashier = strings[7].split(':')[-1].strip()
        return {"address": address, "date": check_date, "cashier": cashier, "total": total}