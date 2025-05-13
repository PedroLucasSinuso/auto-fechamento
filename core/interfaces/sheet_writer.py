from abc import ABC, abstractmethod
from core.entities.report import Report
from datetime import datetime

class SheetWriter(ABC):
    @abstractmethod
    def write(self, sheet_path: str, report: Report, day: datetime, 
              collect: bool = False, change_requested: float = 0, 
              cash_fund: float = 0.0):
        pass