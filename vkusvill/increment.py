from pathlib import Path


class IncrementHandler:
    """
    Хранит/читает последний обработанный UID
    в текстовом файле. 0 == «еще не было».
    """

    def __init__(self, file_path: Path) -> None:
        self.file_path = file_path
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.file_path.exists():
            self.file_path.write_text("0")

    def get(self) -> int:
        """Возвращает сохранённый UID."""
        return int(self.file_path.read_text().strip() or "0")

    def set(self, value: int) -> None:
        """Записывает новый UID."""
        self.file_path.write_text(str(value))