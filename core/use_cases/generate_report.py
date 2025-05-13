import re
from dataclasses import dataclass
from core.entities.report import Report
from core.interfaces.file_reader import FileReader
from infrastructure.config.settings import (
    PAYMENT_METHODS, 
    CASH_INFLOW_PATTERNS, 
    CASH_OUTFLOW_PATTERNS
)

@dataclass
class ReportGenerator:
    file_reader: FileReader
    
    def extract_value(self, line: str) -> float:
        #Extrai valor numérico de uma linha
        match = re.search(r'\d+,\d+', line)
        return float(match.group().replace(',', '.')) if match else 0.0
    
    def process_terminal(self, lines: list[str], index: int, report: Report):
        #Processa dados de um terminal
        terminal_match = re.search(r'\d+', lines[index])
        if not terminal_match:
            return
            
        terminal = terminal_match.group()
        venda_bruta = self.extract_value(lines[index + 2])
        
        if venda_bruta <= 0:
            return
            
        if terminal not in report.terminals:
            report.terminals.append(terminal)
            report.gross_sales[terminal] = venda_bruta
            report.exchanged_items += self.extract_value(lines[index+8])
            report.gross_add += self.extract_value(lines[index+4])
            report.gross_add += self.extract_value(lines[index+6])
            report.discounts += self.extract_value(lines[index+3])
            report.discounts += self.extract_value(lines[index+5])
    
    def process_financial_entries(self, line: str, report: Report):
        #Processa entradas financeiras especiais
        categories = {
            r'FRETE\s+B2C': 'shipping',
            r'OMNICHANNEL': 'omnichannel',
            r'\bCREDSYSTEM\b': 'credsystem'
        }
        
        for pattern, key in categories.items():
            if re.search(pattern, line):
                value = self.extract_value(line)
                setattr(report, key, getattr(report, key) + value)
    
    def process_payment_methods(self, line: str, report: Report):
        #Processa métodos de pagamento
        for pattern, label in PAYMENT_METHODS.items():
            if re.search(pattern, line):
                value = self.extract_value(line)
                if label == 'DINHEIRO':
                    value /= 2  # Caso especial para dinheiro
                report.payment_methods[label] = report.payment_methods.get(label, 0) + value
    
    def process_cash_flows(self, line: str, report: Report, category: str, patterns: list[str]):
        #Processa fluxos de caixa
        total_key = f'total_{category}'
        
        for pattern in patterns:
            if re.search(pattern, line):
                value = self.extract_value(line)
                getattr(report, category)[pattern] = value
                setattr(report, total_key, getattr(report, total_key) + value)
                
                if pattern == "PREMIAÇÃO CREDSYSTEM":
                    report.credsystem -= value
    
    def generate(self, file_path: str) -> Report:
        #Gera relatório consolidado
        lines = self.file_reader.read(file_path)
        report = Report()
        
        for i, line in enumerate(lines):
            if re.search(r'Terminal', line):
                self.process_terminal(lines, i, report)
                
            self.process_financial_entries(line, report)
            self.process_payment_methods(line, report)
            self.process_cash_flows(line, report, 'cash_inflow', CASH_INFLOW_PATTERNS)
            self.process_cash_flows(line, report, 'cash_outflow', CASH_OUTFLOW_PATTERNS)
        
        return report