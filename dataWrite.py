from openpyxl import load_workbook
import re
import streamlit as st

# Função para iniciar um dia novo
def startNewDay(sheet, day, collect=False):
    try:
        # Carregando sheet data_only
        wbDataOnly = load_workbook(sheet, data_only=True)
        wsSaldoCaixaDataOnly = wbDataOnly["Saldo em Caixa"]

        # Carregar sheet com macros
        wb = load_workbook(sheet, keep_vba=True)
        wsSaldoCaixa = wb["Saldo em Caixa"]
        wsRelFechamento = wb["Rel. Fechamento de Caixa"]

        # Remover proteção das planilhas
        wsSaldoCaixa.protection.disable()
        wsRelFechamento.protection.disable()

        # Subir dinheiro em espécie e credsystem
        wsSaldoCaixa["F12"] = wsSaldoCaixaDataOnly["F29"].value
        wsSaldoCaixa["F14"] = wsSaldoCaixaDataOnly["F30"].value

        # Dia de coleta
        if collect:
            wsSaldoCaixa["F16"] = wsSaldoCaixaDataOnly["F29"].value
            wsSaldoCaixa["F18"] = wsSaldoCaixaDataOnly["F30"].value
        else:
            wsSaldoCaixa["F16"] = None
            wsSaldoCaixa["F18"] = None

        # Zerar as células especificadas
        for cell_range in ["B13:B20", "B24:B25", "G13:G23", "G25:G26"]:
            for row in wsRelFechamento[cell_range]:
                for cell in row:
                    cell.value = None

        # Zerar troco
        wsSaldoCaixa["F20"] = None
        wsSaldoCaixa["F22"] = None
        # Atualizar a data das células
        wsRelFechamento["B7"] = day.strftime('%d/%m/%Y')
        wsSaldoCaixa["G7"] = day.strftime('%d/%m/%Y')

        return wb

    except Exception as e:
        error_message = f"Erro ao iniciar um novo dia: {e}"
        print(error_message)
        st.warning(error_message)
        return None

# Função para editar a planilha
def sheetEdit(sheet, report, day, collect=False):
    wb = startNewDay(sheet, day, collect)
    if wb is None:
        error_message = "Erro ao iniciar um novo dia"
        print(error_message)
        st.warning(error_message)
        return None

    wsRelFechamento = wb["Rel. Fechamento de Caixa"]

    # Inserir valores de terminais (001-006) nas células B13-B18
    for i in range(1, 7):
        try:
            terminal_key = f"{i:03}"
            cell = f"B{12 + i}"
            if terminal_key in report['Gross_Sales']:
                wsRelFechamento[cell] = report['Gross_Sales'][terminal_key]
        except Exception as e:
            error_message = f"Erro ao inserir valor do terminal {terminal_key}: {e}"
            print(error_message)
            st.warning(error_message)

    # Inserir valor de acréscimo total na célula B19
    try:
        wsRelFechamento["B19"] = report['Gross_Add']
    except Exception as e:
        error_message = f"Erro ao inserir valor de acréscimo total: {e}"
        print(error_message)
        st.warning(error_message)

    # Inserir valor de desconto total na célula B25
    try:
        wsRelFechamento["B25"] = report['Discounts']
    except Exception as e:
        error_message = f"Erro ao inserir valor de desconto total: {e}"
        print(error_message)
        st.warning(error_message)

    # Mapear outras informações para células específicas
    mappings = {
        'Exchanged_Items': "B24",
        'Shipping': "G23",
        'Expenses': "G25",
        'Credsystem': "G26",
        'Omnichannel': "G19",
        'Total_Cash_Outflow': "G25",
    }
    for key, cell in mappings.items():
        try:
            if key in report:
                wsRelFechamento[cell] = report[key]
        except Exception as e:
            error_message = f"Erro ao mapear {key}: {e}"
            print(error_message)
            st.warning(error_message)
    
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
        try:
            if any(re.search(pattern, item) for item in report['Payment_Methods']):
                match_item = next(item for item in report['Payment_Methods'] if re.search(pattern, item))
                wsRelFechamento[cell] = report['Payment_Methods'][match_item]
        except Exception as e:
            error_message = f"Erro ao inserir método de pagamento {pattern}: {e}"
            print(error_message)
            st.warning(error_message)

    return wb