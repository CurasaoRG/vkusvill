# vkusvill/parsers/ofd.py
from __future__ import annotations

import re
from decimal import Decimal
from datetime import datetime

from bs4 import BeautifulSoup

from .base import BaseParser
from ..models import Check, Item


class OfdParser(BaseParser):
    """
    Парсер для писем от check.ofd.ru (HTML-версия, тип 'ofd').
    Возвращает объект Check с полями и списком Item-ов.
    """

    def parse(self, raw_body: bytes | str) -> Check:
        html = raw_body.decode() if isinstance(raw_body, bytes) else raw_body
        soup = BeautifulSoup(html, "html.parser")
        strings = list(soup.stripped_strings)
        # ---------- Парсинг позиций ----------
        items: list[Item] = []
        info_section = None
        goods_section = False
        row: list[str] = []
        header_strings: list[str] = []
        k = 0
        last_item_field = 0
        for idx, text in enumerate(strings):
            if text in ('Кассовый чек / Приход', 'Кассовый чек / Возврат прихода'):
                info_section = True
                continue
            if text == "check.ofd.ru":
                info_section = False
                goods_section = True
                continue
            if text == "ИТОГ":
                info_section = True
                goods_section = False
            if info_section:
                header_strings.append(text)
            if goods_section:
                if k < 4:
                    row.append(text)
                    k += 1
                elif text == "Мера кол-ва предмета расчета":
                    last_item_field = idx
                    k = 0
                if idx == last_item_field + 1:
                    k = 0
                    items.append(self._build_item(row))
                    row = []

        # ---------- Парсинг заголовка ----------
        header_map = self._extract_header(header_strings)
        return Check(
            msg_type="ofd",
            address1=header_map.get("address1", ""),
            address2=header_map.get("address2", ""),
            date=header_map["date"],
            cashier=header_map.get("cashier", ""),
            total=header_map["total"],
            items=items,
        )

    # ---------- Внутренние хелперы ----------
    @staticmethod
    def _build_item(row: list[str]) -> Item:
        try:
            if "X" in row[0]:
                name = "N/A"
                price = Decimal(row[0].split(" X ")[1].replace(",", "."))
                qty = Decimal(row[0].split("X")[0].replace(",", "."))
                amount = Decimal(row[3].split("=")[1].replace(",", "."))
                uom = row[4].split(".")[0]
            else:
                name = row[0]
                price = Decimal(row[1].split(" X ")[1].replace(",", "."))
                qty = Decimal(row[1].split("X")[0].replace(",", "."))
                amount = Decimal(row[3].split("=")[1].replace(",", "."))
                uom = row[4].split(".")[0]
        except (IndexError, ValueError):
            name, price, qty, amount, uom = "N/A", Decimal("0"), Decimal("0"), Decimal("0"), "шт"
        return Item(product_name=name, price=price, qty=qty, amount=amount, uom=uom)

    @staticmethod
    def _extract_header(strings: list[str]) -> dict:
        fields = {
            "#": "check_fd",
            "НОМЕР СМЕНЫ": "shift_num",
            "МЕСТО РАСЧЁТОВ": "address1",
            "АДРЕС РАСЧЁТОВ": "address2",
            "ДАТА ВЫДАЧИ": "date_raw",
            "КАССИР": "cashier",
            "ИТОГ": "total_raw",
        }
        position = -1
        field_name = None
        data = {}
        for idx, text in enumerate(strings):
            if text in fields.keys():
                field_name = fields[text]
                position = idx + 1
            elif re.fullmatch(r"^.+ \d+$", text):
                key, val = text.rsplit(" ", 1)
                if key.strip() in fields.values():
                    data[key.strip()] = val
            elif idx == position and field_name:
                data[field_name] = text

        # Преобразуем дату и сумму
        date_obj = datetime.strptime(data.pop("date_raw"), "%d.%m.%y %H:%M")
        total = Decimal(data.pop("total_raw").replace(",", "."))

        return {"date": date_obj, "total": total, **data}