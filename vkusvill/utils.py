from __future__ import annotations

import re
from decimal import Decimal
from datetime import datetime

NON_BREAKING_SPACE = '\xa0'


def parse_decimal(input_string: str) -> Decimal:
    for char in (NON_BREAKING_SPACE, ' ',):
        input_string = input_string.replace(char, '')
    return Decimal(input_string.replace(",", "."))