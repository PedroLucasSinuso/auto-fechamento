from abc import ABC, abstractmethod
from core.entities.report import Report
from io import BytesIO
from datetime import datetime

class ReportPresenter(ABC):
    @abstractmethod
    def show_report_summary(self, report: Report):
        pass
    
    @abstractmethod
    def show_comparison_result(self, is_match: bool):
        pass
    
    @abstractmethod
    def get_download_button(self, file_content: BytesIO, filename: str):
        pass
    
    @abstractmethod
    def prepare_download(self, wb, day: datetime) -> tuple:
        pass