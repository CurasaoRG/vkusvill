import csv
from pathlib import Path
from typing import Iterable
from src.models import Check

class CsvRepository:
    CSV_PARAMS = dict(delimiter=";", quotechar="|", quoting=csv.QUOTE_MINIMAL)

    def __init__(self, root: Path):
        self.root = root
        self.check_info = root / "check_info.csv"
        self.items = root / "data.csv"
        self._ensure_headers()

    def _ensure_headers(self):
        for file, headers in (
            (self.check_info, ["id", "msg_type", "address1", "address2", "date", "cashier", "total"]),
            (self.items, ["id", "msg_type", "product_name", "price", "qty", "amount", "uom"]),
        ):
            if not file.exists() or file.stat().st_size == 0:
                file.write_text(";".join(headers) + "\n")

    def append_check(self, check_id: int, check: "Check"):
        with self.check_info.open("a", newline="") as f:
            writer = csv.writer(f, **self.CSV_PARAMS)
            writer.writerow(check.check_info_row(check_id))

        with self.items.open("a", newline="") as f:
            writer = csv.writer(f, **self.CSV_PARAMS)
            for item in check.items:
                writer.writerow([check_id, check.msg_type] + item.to_csv_row())