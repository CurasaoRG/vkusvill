from abc import ABC, abstractmethod
from src.models import Check

class BaseParser(ABC):
    @abstractmethod
    def parse(self, raw_body: bytes | str) -> Check:
        ...