import aspose.words as aw
import re
from config import DATE_PATTERN, PAYMENT_METHODS, CASH_INFLOW_PATTERNS, CASH_OUTFLOW_PATTERNS

# Função para extrair texto de um arquivo Word, filtrando linhas irrelevantes
def extract_text_from_word(word_file):
    doc = aw.Document(word_file)
    lines = [p.get_text().strip() for p in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True) if p.get_text().strip()]
    filtered_lines = [line for line in lines if "Filial" not in line and "Aspose.Words" not in line and not re.search(DATE_PATTERN, line)]
    return filtered_lines

# Função para extrair valores numéricos de uma linha com base em um padrão
def extract_value_from_line(line, pattern):
    match = re.search(pattern, line)
    return float(match.group().replace(',', '.')) if match else 0

# Processa os dados de um terminal específico e atualiza o relatório
def process_terminal_data(lines, index, report):
    terminal = re.search(r'\d+', lines[index]).group()
    venda_bruta = extract_value_from_line(lines[index + 2], r'\d+,\d+')
    if venda_bruta <= 0: return
    if terminal not in report['Terminals']:
        report['Terminals'].append(terminal)
        report['Gross_Sales'][terminal] = extract_value_from_line(lines[index+2], r'\d+,\d+')
        report['Exchanged_Items'] += extract_value_from_line(lines[index+8], r'\d+,\d+')
        report['Gross_Add'] += extract_value_from_line(lines[index+4], r'\d+,\d+')
        report['Gross_Add'] += extract_value_from_line(lines[index+6], r'\d+,\d+')
        report['Discounts'] += extract_value_from_line(lines[index+3], r'\d+,\d+')
        report['Discounts'] += extract_value_from_line(lines[index+5], r'\d+,\d+')

# Processa entradas financeiras específicas e atualiza o relatório
def process_financial_entries(line, report):
    categories = {
        r'FRETE\s+B2C': 'Shipping',
        r'OMNICHANNEL': 'Omnichannel',
        r'\bCREDSYSTEM\b': 'Credsystem'
    }
    for pattern, key in categories.items():
        report[key] += extract_value_from_line(line, r'\d+,\d+') if re.search(pattern, line) else 0

# Processa métodos de pagamento encontrados em uma linha e atualiza o relatório
def process_payment_methods(line, report):
    for pattern, label in PAYMENT_METHODS.items():
        if re.search(pattern, line):
            value = extract_value_from_line(line, r'\d+,\d+')
            if label == 'DINHEIRO':
                value /= 2  # Special case for cash
            report['Payment_Methods'][label] = report['Payment_Methods'].get(label, 0) + value

# Processa fluxos de caixa (entrada ou saída) e atualiza o relatório
def process_cash_flows(line, report, category, patterns):
    for pattern in patterns:
        if re.search(pattern, line):
            value = extract_value_from_line(line, r'\d+,\d+')
            report[category][pattern] = value
            report[f'Total_{category}'] += value
            if pattern == "PREMIAÇÃO CREDSYSTEM":
                report['Credsystem'] -= value

# Gera o relatório consolidado a partir do arquivo Word
def genReport(word_file) -> dict:
    lines = extract_text_from_word(word_file)
    # Inicializa o dicionário do relatório
    report = {
        'Terminals': [],
        'Gross_Sales': {},
        'Gross_Add': 0,
        'Discounts': 0,
        'Exchanged_Items': 0,
        'Shipping': 0,
        'Cash_Inflow': {},
        'Cash_Outflow': {},
        'Total_Cash_Inflow': 0,
        'Total_Cash_Outflow': 0,
        'Expenses': 0,
        'Credsystem': 0,
        'Omnichannel': 0,
        'Payment_Methods': {}
    }
    
    # Itera pelas linhas extraídas do arquivo e processa os dados
    for i, line in enumerate(lines):
        if re.search(r'Terminal', line):
            process_terminal_data(lines, i, report)
        process_financial_entries(line, report)
        process_payment_methods(line, report)
        process_cash_flows(line, report, 'Cash_Inflow', CASH_INFLOW_PATTERNS)
        process_cash_flows(line, report, 'Cash_Outflow', CASH_OUTFLOW_PATTERNS)
    
    return report