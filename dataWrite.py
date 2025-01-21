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

# Função para editar a planilha "Rel. Fechamento de Caixa"
def sheetEdit(sheet, report, collect=False):
    wb = startNewDay(sheet, collect)
    wsRelFechamento = wb["Rel. Fechamento de Caixa"]

    # Inserir na célula B13 o valor do terminal 001 se ele existir
    if '001' in report['Gross_Sales']: wsRelFechamento["B13"] = report['Gross_Sales']['001']
    # Inserir na célula B14 o valor do terminal 002 se ele existir
    if '002' in report['Gross_Sales']: wsRelFechamento["B14"] = report['Gross_Sales']['002']
    # Inserir na célula B15 o valor do terminal 003 se ele existir
    if '003' in report['Gross_Sales']: wsRelFechamento["B15"] = report['Gross_Sales']['003']
    # Inserir na célula B16 o valor do terminal 004 se ele existir
    if '004' in report['Gross_Sales']: wsRelFechamento["B16"] = report['Gross_Sales']['004']
    # Inserir na célula B17 o valor do terminal 005 se ele existir
    if '005' in report['Gross_Sales']: wsRelFechamento["B17"] = report['Gross_Sales']['005']
    # Inserir na célula B18 o valor do terminal 006 se ele existir
    if '006' in report['Gross_Sales']: wsRelFechamento["B18"] = report['Gross_Sales']['006']
    # Inserir o valor total de trocas na célula B24 se ele existir
    if 'Exchanged_Items' in report: wsRelFechamento["B24"] = report['Exchanged_Items']
    # Inserir o valor total de frete na célula G23 se ele existir
    if 'Shipping' in report: wsRelFechamento["G23"] = report['Shipping']
    # Inserir o valor total de movimentações positivas na célula G26 se ele existir
    if 'movPositive' in report['Cash_Movement']: wsRelFechamento["G26"] = report['Cash_Movement']['movPositive']
    # Inserir o valor total de movimentações negativas na célula G25 se ele existir
    if 'movNegative' in report['Cash_Movement']: wsRelFechamento["G25"] = report['Cash_Movement']['movNegative']
    # Inserir o valor total do Omnichannel na célula G19 se ele existir
    if 'Omnichannel' in report: wsRelFechamento["G19"] = report['Omnichannel']
    #Ieserir o valor total de dinheiro na célula G13 se ele existir
    if any(re.search(r'DINHEIRO', item) for item in report['Payment_Methods']): 
        wsRelFechamento["G13"] = report['Payment_Methods'][next(item for item in report['Payment_Methods'] if re.search(r'DINHEIRO', item))]
    # Inserir o valor de QR na célula G18 se ele existir
    if any(re.search(r'QR', item) for item in report['Payment_Methods']):
        wsRelFechamento["G18"] = report['Payment_Methods'][next(item for item in report['Payment_Methods'] if re.search(r'QR', item))]
    # Inserir o valor de Cartão de Crédito PDV na célula G14 se ele existir
    if any(re.search(r'CARTAO\s*CREDITO\s*PDV', item) for item in report['Payment_Methods']): 
        wsRelFechamento["G14"] = report['Payment_Methods'][next(item for item in report['Payment_Methods'] if re.search(r'CARTAO\s*CREDITO\s*PDV', item))]
    # Inserir o valor de Cartão de Débito PDV na célula G15 se ele existir
    if any(re.search(r'CARTAO\s*DEBITO\s*PDV', item) for item in report['Payment_Methods']): 
        wsRelFechamento["G15"] = report['Payment_Methods'][next(item for item in report['Payment_Methods'] if re.search(r'CARTAO\s*DEBITO\s*PDV', item))]
    # Inserir o valor de Cartão de Crédito POS na célula G16 se ele existir
    if any(re.search(r'CARTAO\s*CREDITO$', item) for item in report['Payment_Methods']): 
        wsRelFechamento["G16"] = report['Payment_Methods'][next(item for item in report['Payment_Methods'] if re.search(r'CARTAO\s*CREDITO$', item))]
    # Inserir o valor de Cartão de Débito POS na célula G17 se ele existir
    if any(re.search(r'CARTAO\s*DEBITO$', item) for item in report['Payment_Methods']): 
        wsRelFechamento["G17"] = report['Payment_Methods'][next(item for item in report['Payment_Methods'] if re.search(r'CARTAO\s*DEBITO$', item))]
    
    return wb


