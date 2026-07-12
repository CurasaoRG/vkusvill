from abc import ABC, abstractmethod
from src.models import Check

class BaseRepository(ABC):
    @abstractmethod
    def append_check(self, check_id: int, check: Check):
        ...