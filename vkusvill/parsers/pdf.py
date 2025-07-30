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
    Использует PyPDF c layout-режимом для устойчивого извлечения текста.
    """

    MONTHS_RU = {
        "января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5, "июня": 6,
        "июля": 7, "августа": 8, "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12
    }

    def parse(self, raw_body: bytes | BytesIO) -> Check:
        reader = PdfReader(BytesIO(raw_body) if isinstance(raw_body, bytes) else raw_body)

        # ---------- Регулярки ----------
        # 1. Дата выдачи: «12 марта 2024 г. в 14:32»
        date_re = re.compile(
            r"Дата выдачи\s+(\d{1,2})\s+([а-яё]+)\s+(\d{4})\s+г\.\s+в\s+(\d{1,2}:\d{2})",
            re.I,
        )
        # 2. Место расчёта (первая строка)
        place_re = re.compile(r"Место осуществления расчета\s+(.+)", re.I)
        # 3. Адрес (многострочный, берём до ключевого слова «ИТОГ»)
        addr_re = re.compile(r"Адрес осуществления расчетов\s+(.+?)(?=\s*ИТОГ)", flags=re.S)

        item_re = re.compile(r"(?:\d+ +)(.+?)\s+(\d+,\d+)\s+(\d+,\d+)\s+(\d+,\d+)", re.MULTILINE)
        text = ""
        for page in reader.pages:
            text += page.extract_text(extraction_mode="layout") + "\n"

        # ---------- Дата ----------
        m = date_re.search(text)
        if not m:
            raise ValueError("Не удалось найти дату выдачи")
        day, month_ru, year, time = m.groups()
        date = datetime(
            int(year),
            self.MONTHS_RU[month_ru.lower()],
            int(day),
            *map(int, time.split(":")),
        )

        # ---------- Адрес ----------
        place = place_re.search(text)
        address1 = place.group(1).strip() if place else ""
        addr = addr_re.search(text)
        address2 = addr.group(1).replace("\n", " ").strip() if addr else ""

        # ---------- Итог ----------
        total_line = re.search(r"ИТОГ\s+(\d[\d\s]*,\d{2})", text, flags=re.I)
        if not total_line:
            raise ValueError("Не удалось найти итог")
        total = Decimal(normalize("NFKD", total_line.group(1)).replace(" ", "").replace(",", "."))

        # ---------- Позиции ----------
        items: List[Item] = []
        for name, price_s, qty_s, amount_s in item_re.findall(text):
            items.append(
                Item(
                    product_name=name.strip(),
                    price=Decimal(price_s.replace(",", ".")),
                    qty=Decimal(qty_s.replace(",", ".")),
                    amount=Decimal(amount_s.replace(",", ".")),
                    uom="шт",  # PDF не содержит ЕИ
                )
            )

        return Check(
            msg_type="pdf",
            address1=address1,
            address2=address2,
            date=date,
            cashier="N/A",  # PDF не показывает кассира
            total=total,
            items=items,
        )