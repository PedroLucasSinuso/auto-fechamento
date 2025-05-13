from dataclasses import dataclass
from core.entities.report import Report
from core.interfaces.sheet_writer import SheetWriter
from datetime import datetime

@dataclass
class SheetProcessor:
    sheet_writer: SheetWriter
    
    def process(self, sheet_path: str, report: Report, day: datetime, 
               collect: bool = False, change_requested: float = 0, 
               cash_fund: float = 0.0):
        #Processa a planilha com os dados do relatório
        return self.sheet_writer.write(
            sheet_path,
            report,
            day,
            collect,
            change_requested,
            cash_fund
        )