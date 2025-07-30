# vkusvill/parsers/pdf.py
from __future__ import annotations

import re
from decimal import Decimal
from datetime import datetime
from io import BytesIO
from typing import List

from pypdf import PdfReader
from unicodedata import normalize

from .base import BaseParser
from ..models import Check, Item


class PdfParser(BaseParser):
    """
    Парсер PDF-чека из писем noreply-cloudkassir@cp.ru (тип 'pdf').
    Умеет работать с датами в любых падежах, например:
      «15 марта 2025 г. в 14:32»  (родительный)
      «07 ноябрь 2024 г. в 11:26» (именительный)
    """

    # Маппинг месяцев «в любом падеже» → номер месяца
    MONTH_MAP = {
        "январь": 1, "января": 1,
        "февраль": 2, "февраля": 2,
        "март": 3, "марта": 3,
        "апрель": 4, "апреля": 4,
        "май": 5, "мая": 5,
        "июнь": 6, "июня": 6,
        "июль": 7, "июля": 7,
        "август": 8, "августа": 8,
        "сентябрь": 9, "сентября": 9,
        "октябрь": 10, "октября": 10,
        "ноябрь": 11, "ноября": 11,
        "декабрь": 12, "декабря": 12,
    }

    def parse(self, raw_body: bytes | BytesIO) -> Check:
        reader = PdfReader(BytesIO(raw_body) if isinstance(raw_body, bytes) else raw_body)

        # ---------- Регулярки ----------
        # 1. Дата выдачи: «12 марта 2025 г. в 14:32» или «07 ноябрь 2024 г. в 11:26»
        date_re = re.compile(
            r"Дата выдачи\s+(\d{1,2})\s+([а-яё]+)\s+(\d{4})\s+г\.\s+в\s+(\d{1,2}:\d{2})",
            re.I,
        )
        place_re = re.compile(r"Место осуществления расчета\s+(.+)", re.I)
        addr_re = re.compile(r"Адрес осуществления расчетов\s+([^\n]+)",  flags=re.I,)
        item_re = re.compile(
            r"(?:\d+ +)(.+?)\s+(\d+,\d+)\s+(\d+,\d+)\s+(\d+,\d+)", re.M
        )

#         item_re = re.compile(
#     r"^\s*(\d+)\s+(.+?)\s+(\d+(?:\s?\d+)*(?:,\d{2}))\s+(\d+(?:\s?\d+)*(?:,\d{2}))\s+(\d+(?:\s?\d+)*(?:,\d{2}))",
#     re.MULTILINE | re.DOTALL,
# )
        text = ""
        for page in reader.pages:
            text += page.extract_text(extraction_mode="layout") + "\n"

        # ---------- Дата (устойчивая к падежу) ----------
        m = date_re.search(text)
        if not m:
            raise ValueError("Не удалось найти дату выдачи")
        day, month_word, year, time = m.groups()
        month_num = self.MONTH_MAP.get(month_word.lower())
        if month_num is None:
            raise ValueError(f"Неизвестный месяц: {month_word}")

        date = datetime(
            int(year),
            month_num,
            int(day),
            *map(int, time.split(":")),
        )

        # ---------- Адрес ----------
        place = place_re.search(text)
        address1 = place.group(1).strip() if place else ""
        # --- Адрес ---
        addr = addr_re.search(text)
        address2 = addr.group(1).strip() if addr else ""
        # Очищаем от лишних пробелов и переносов
        address2 = " ".join(address2.split())

        # ---------- Итог ----------
        total_line = re.search(r"ИТОГ\s+(\d[\d\s]*,\d{2})", text, flags=re.I)
        if not total_line:
            raise ValueError("Не удалось найти итог")
        total = Decimal(
            normalize("NFKD", total_line.group(1)).replace(" ", "").replace(",", ".")
        )

        # ---------- Позиции ----------
        items: List[Item] = []
        for name, price_s, qty_s, amount_s in item_re.findall(text):
            items.append(
                Item(
                    product_name=name.strip(),
                    price=Decimal(price_s.replace(",", ".")),
                    qty=Decimal(qty_s.replace(",", ".")),
                    amount=Decimal(amount_s.replace(",", ".")),
                    uom="шт",  # в PDF ЕИ не указана
                )
            )
            print(name, price_s, qty_s, amount_s)
        # items: list[Item] = []
        # for num, name, price_s, qty_s, amount_s in item_re.findall(text):
        #     # собираем многострочное имя в одну строку
        #     name = re.sub(r"\s+", " ", name.strip())
        #     items.append(
        #         Item(
        #             product_name=name,
        #             price=Decimal(price_s.replace(" ", "").replace(",", ".")),
        #             qty=Decimal(qty_s.replace(" ", "").replace(",", ".")),
        #             amount=Decimal(amount_s.replace(" ", "").replace(",", ".")),
        #             uom="шт",
        #         )
        #     )
            # print(num, name, price_s, qty_s, amount_s)


        return Check(
            msg_type="pdf",
            address1=address1,
            address2=address2,
            date=date,
            cashier="N/A",
            total=total,
            items=items,
        )