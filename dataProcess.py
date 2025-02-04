import aspose.words as aw
import re
import streamlit as st

def extract_text_from_word(word_file):
    # Extracts and returns a list of non-empty lines from a Word document.
    doc = aw.Document(word_file)
    return [p.get_text().strip() for p in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True) if p.get_text().strip()]

def extract_value_from_line(line, pattern):
    # Extracts and converts a numeric value from a given line based on a regex pattern.
    match = re.search(pattern, line)
    return float(match.group().replace(',', '.')) if match else 0

def process_terminal_data(lines, index, report):
    # Processes terminal-related data and updates the report dictionary.
    terminal = re.search(r'\d+', lines[index]).group()
    if terminal not in report['Terminals']:
        report['Terminals'].append(terminal)
        report['Gross_Sales'][terminal] = extract_value_from_line(lines[index+2], r'\d+,\d+')
        report['Exchanged_Items'] += extract_value_from_line(lines[index+8], r'\d+,\d+')
        report['Gross_Add'] += extract_value_from_line(lines[index+4], r'\d+,\d+')
        report['Gross_Add'] += extract_value_from_line(lines[index+6], r'\d+,\d+')
        report['Discounts'] += extract_value_from_line(lines[index+3], r'\d+,\d+')
        report['Discounts'] += extract_value_from_line(lines[index+5], r'\d+,\d+')

def process_financial_entries(line, report):
    # Processes financial information like shipping, omnichannel, and credsystem.
    categories = {
        r'FRETE\s+B2C': 'Shipping',
        r'OMNICHANNEL': 'Omnichannel',
        r'\bCREDSYSTEM\b': 'Credsystem'
    }
    for pattern, key in categories.items():
        report[key] += extract_value_from_line(line, r'\d+,\d+') if re.search(pattern, line) else 0

def process_payment_methods(line, report):
    # Extracts payment method details and updates the report.
    payment_methods = {
        r'DINHEIRO': 'DINHEIRO',
        r'QR': 'QR',
        r'CARTAO\s+CREDITO\s+PDV': 'CARTAO CREDITO PDV',
        r'CARTAO\s+DEBITO\s+PDV': 'CARTAO DEBITO PDV',
        r'CARTAO\s+CREDITO$': 'CARTAO CREDITO',
        r'CARTAO\s+DEBITO$': 'CARTAO DEBITO'
    }
    for pattern, label in payment_methods.items():
        if re.search(pattern, line):
            value = extract_value_from_line(line, r'\d+,\d+')
            if label == 'DINHEIRO':
                value /= 2  # Special case for cash
            report['Payment_Methods'][label] = report['Payment_Methods'].get(label, 0) + value

def process_cash_flows(line, report, category, patterns):
    # Processes cash inflow or outflow entries.
    for pattern in patterns:
        if re.search(pattern, line):
            value = extract_value_from_line(line, r'\d+,\d+')
            report[category][pattern] = value
            report[f'Total_{category}'] += value

def genReport(word_file) -> dict:
    lines = extract_text_from_word(word_file)
    report = {
        'Terminals': [], 'Gross_Sales': {}, 'Gross_Add': 0, 'Discounts': 0,
        'Exchanged_Items': 0, 'Shipping': 0, 'Cash_Inflow': {}, 'Cash_Outflow': {},
        'Total_Cash_Inflow': 0, 'Total_Cash_Outflow': 0, 'Expenses': 0,
        'Credsystem': 0, 'Omnichannel': 0, 'Payment_Methods': {}
    }
    
    cash_inflow_patterns = [
        r"ENTRADA DE TROCO", r"ENTRADA FUNDO DE CAIXA", r"ENTRADA NO CAIXA", r"SALDO INICIAL", r"VENDA FRANQUIA"
    ]
    
    cash_outflow_patterns = [
        r"BAINHAS E MATERIAL DE COSTURA", r"COLETA PROSEGUR", r"CRM BONUS", r"CÓPIAS/XEROX", r"DEPÓSITO DE TROCO",
        r"DESCONTO UNIFORME", r"DESP - BRINDES", r"DESP-OUTROS MAT. DE CONSU", r"DESPESA MATERIAL COPA",
        r"DESPESA NOTA FISCAL", r"DIFERENCA DE TROCA", r"ESTACIONAMENTO", r"HIGIENE E LIMPEZA", r"LANCHE",
        r"MANUTENÇÃO E REPAROS", r"MATERIAL ESCRITÓRIO E PAPELARIA", r"MEDICINA DO TRABALHO", r"PASSAGENS/CONDUÇÕES",
        r"PREMIAÇÃO CREDSYSTEM", r"PRÊMIOS E GRATIFICAÇÕES", r"RETIRADA DO CAIXA", r"SEGURANÇA E VIGILÂNCIA",
        r"SUPLEMENTO DE INFORMATICA", r"TREINAMENTO", r"UNIFORME", r"VIAGEM", r"VITRINE/MATERIAL DE VITRINE", r"ÁGUA"
    ]
    
    for i, line in enumerate(lines):
        if re.search(r'Terminal', line):
            process_terminal_data(lines, i, report)
        process_financial_entries(line, report)
        process_payment_methods(line, report)
        process_cash_flows(line, report, 'Cash_Inflow', cash_inflow_patterns)
        process_cash_flows(line, report, 'Cash_Outflow', cash_outflow_patterns)
    
    return report

if __name__ == "__main__":
    print(genReport("output.doc"))
