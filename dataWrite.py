from openpyxl import load_workbook
import re

# Função para iniciar um dia novo
def startNewDay(sheet, collect=False):
    # Carregando sheet data_only
    wbDataOnly = load_workbook(sheet, data_only=True)
    wsSaldoCaixaDataOnly = wbDataOnly["Saldo em Caixa"]

    # Carregando sheet com macros
    wb = load_workbook(sheet, keep_vba=True)
    wsSaldoCaixa = wb["Saldo em Caixa"]
    wsRelFechamento = wb["Rel. Fechamento de Caixa"]

    # Remover proteção das planilhas
    wsSaldoCaixa.protection.disable()
    wsRelFechamento.protection.disable()

    # Copiar valor de F29 para F12
    wsSaldoCaixa["F12"] = wsSaldoCaixaDataOnly["F29"].value
    wsSaldoCaixa["F14"] = wsSaldoCaixaDataOnly["F30"].value

    # Se tiver coleta
    if collect:
        wsSaldoCaixa["F16"] = wsSaldoCaixaDataOnly["F29"].value
        wsSaldoCaixa["F18"] = wsSaldoCaixaDataOnly["F30"].value
    else:
        wsSaldoCaixa["F16"] = 0
        wsSaldoCaixa["F18"] = 0

    # Zerar as células especificadas
    for cell_range in ["B13:B20", "B24:B25", "G13:G23", "G25:G26"]:
        for row in wsRelFechamento[cell_range]:
            for cell in row:
                cell.value = 0

    return wb

def sheetEdit(sheet, report, collect=False):
    wb = startNewDay(sheet, collect)
    wsRelFechamento = wb["Rel. Fechamento de Caixa"]

    # Inserir valores de terminais (001-006) nas células B13-B18
    for i in range(1, 7):
        terminal_key = f"{i:03}"
        cell = f"B{12 + i}"
        if terminal_key in report['Gross_Sales']:
            wsRelFechamento[cell] = report['Gross_Sales'][terminal_key]

    # Inserir valor de acréscimo total na célula B19
    wsRelFechamento["B19"] = report['Gross_Add']

    # Inserir valor de descontos total na célula B25
    wsRelFechamento["B25"] = report['Discounts']

    # Mapear outras informações para células específicas
    mappings = {
        'Exchanged_Items': "B24",
        'Shipping': "G23",
        'Expenses': "G25",
        'Credsystem': "G26",
        'Omnichannel': "G19",
    }
    for key, cell in mappings.items():
        if key in report:
            wsRelFechamento[cell] = report[key]
    
    # Inserir omnichannel em célula específica
    wsRelFechamento["G19"] = report['Omnichannel']

    # Inserir frete em célula específica
    wsRelFechamento["G23"] = report['Shipping']

    # Inserir métodos de pagamento em células específicas
    payment_mappings = {
        r'DINHEIRO': "G13",
        r'QR': "G18",
        r'CARTAO\s*CREDITO\s*PDV': "G14",
        r'CARTAO\s*DEBITO\s*PDV': "G15",
        r'CARTAO\s*CREDITO$': "G16",
        r'CARTAO\s*DEBITO$': "G17",
    }
    for pattern, cell in payment_mappings.items():
        if any(re.search(pattern, item) for item in report['Payment_Methods']):
            match_item = next(item for item in report['Payment_Methods'] if re.search(pattern, item))
            wsRelFechamento[cell] = report['Payment_Methods'][match_item]

    return wb

