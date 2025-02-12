from openpyxl import load_workbook
import re
import streamlit as st

def load_workbooks(sheet):
    # Load the workbooks
    wbDataOnly = load_workbook(sheet, data_only=True)
    wb = load_workbook(sheet, keep_vba=True)
    return wbDataOnly, wb

def disable_protection(wsSaldoCaixa, wsRelFechamento):
    # Remove sheet protection
    wsSaldoCaixa.protection.disable()
    wsRelFechamento.protection.disable()

def update_cash_and_credsystem(wsSaldoCaixa, wsSaldoCaixaDataOnly, collect, changeRequested):
    # Update cash and credsystem values
    wsSaldoCaixa["F12"] = wsSaldoCaixaDataOnly["F29"].value
    wsSaldoCaixa["F14"] = wsSaldoCaixaDataOnly["F30"].value

    # Collection day
    if collect:
        wsSaldoCaixa["F16"] = wsSaldoCaixaDataOnly["F29"].value
        wsSaldoCaixa["F18"] = wsSaldoCaixaDataOnly["F30"].value
    else:
        wsSaldoCaixa["F16"] = None
        wsSaldoCaixa["F18"] = None

    # Change requested
    if changeRequested:
        wsSaldoCaixa["F20"] = changeRequested
        wsSaldoCaixa["F22"] = changeRequested
    else:
        wsSaldoCaixa["F20"] = None
        wsSaldoCaixa["F22"] = None

def clear_cells(wsRelFechamento, wsSaldoCaixa):
    # Clear specified cells
    for cell_range in ["B13:B20", "B24:B25", "G13:G23", "G25:G26"]:
        for row in wsRelFechamento[cell_range]:
            for cell in row:
                cell.value = None

def update_dates(wsRelFechamento, wsSaldoCaixa, day):
    # Update cell dates
    wsRelFechamento["B7"] = day.strftime('%d/%m/%Y')
    wsSaldoCaixa["G7"] = day.strftime('%d/%m/%Y')

def startNewDay(sheet, day, collect, changeRequested):
    try:
        wbDataOnly, wb = load_workbooks(sheet)
        wsSaldoCaixaDataOnly = wbDataOnly["Saldo em Caixa"]
        wsSaldoCaixa = wb["Saldo em Caixa"]
        wsRelFechamento = wb["Rel. Fechamento de Caixa"]

        disable_protection(wsSaldoCaixa, wsRelFechamento)
        update_cash_and_credsystem(wsSaldoCaixa, wsSaldoCaixaDataOnly, collect, changeRequested)
        clear_cells(wsRelFechamento, wsSaldoCaixa)
        update_dates(wsRelFechamento, wsSaldoCaixa, day)

        return wb

    except Exception as e:
        error_message = f"Error starting a new day: {e}"
        print(error_message)
        st.warning(error_message)
        return None

def insert_terminal_values(wsRelFechamento, report):
    # Insert terminal values (001-006) into cells B13-B18
    for i in range(1, 7):
        try:
            terminal_key = f"{i:03}"
            cell = f"B{12 + i}"
            if terminal_key in report['Gross_Sales']:
                wsRelFechamento[cell] = report['Gross_Sales'][terminal_key]
        except Exception as e:
            error_message = f"Error inserting terminal value {terminal_key}: {e}"
            print(error_message)
            st.warning(error_message)

def insert_total_values(wsRelFechamento, report):
    # Insert total addition value into cell B19
    try:
        wsRelFechamento["B19"] = report['Gross_Add']
    except Exception as e:
        error_message = f"Error inserting total addition value: {e}"
        print(error_message)
        st.warning(error_message)

    # Insert total discount value into cell B25
    try:
        wsRelFechamento["B25"] = report['Discounts']
    except Exception as e:
        error_message = f"Error inserting total discount value: {e}"
        print(error_message)
        st.warning(error_message)

def map_other_information(wsRelFechamento, report):
    # Map other information to specific cells
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
            error_message = f"Error mapping {key}: {e}"
            print(error_message)
            st.warning(error_message)

def insert_payment_methods(wsRelFechamento, report):
    # Insert payment methods into specific cells
    payment_mappings = {
        r'DINHEIRO': "G13",
        r'QR': "G18",
        r'CARTAO\s*CREDITO\s*PDV': "G14",
        r'CARTAO\s*DEBITO\s*PDV': "G15",
        r'CARTAO\s+DE\s+CREDITO': "G16",
        r'CARTAO\s+DE\s+DEBITO': "G17",
    }
    for pattern, cell in payment_mappings.items():
        try:
            if any(re.search(pattern, item) for item in report['Payment_Methods']):
                match_item = next(item for item in report['Payment_Methods'] if re.search(pattern, item))
                wsRelFechamento[cell] = report['Payment_Methods'][match_item]
        except Exception as e:
            error_message = f"Error inserting payment method {pattern}: {e}"
            print(error_message)
            st.warning(error_message)

def compare_totals(wsRelFechamento):
    # Helper function to handle NoneType values
    def get_value(cell):
        return cell.value if cell.value is not None else 0

    # Sum the gross total of terminals (B13 to B22)
    gross_total_terminals = sum(get_value(wsRelFechamento[f"B{row}"]) for row in range(13, 23))

    # Subtract the value (B24 + B25) from the gross total of terminals and assign the result to netSale
    netSale = round(gross_total_terminals - (get_value(wsRelFechamento["B24"]) + get_value(wsRelFechamento["B25"])), 2)

    # Sum the gross total of payments from cells G13 to G22
    gross_total_payments = sum(get_value(wsRelFechamento[f"G{row}"]) for row in range(13, 23))

    # Subtract the value of G23 from the gross total of payments and assign the result to netPay
    netPay = round(gross_total_payments - get_value(wsRelFechamento["G23"]), 2) 

    # Compare netPay and netSale
    if netPay == netSale:
        st.success("Bateu! :sunglasses:")
    else:
        st.error("Bateu não. Confere a tabela por favor :skull:")

def sheetEdit(sheet, report, day, collect=False, changeRequested=0):
    wb = startNewDay(sheet, day, collect, changeRequested)
    if wb is None:
        error_message = "Error starting a new day"
        print(error_message)
        st.warning(error_message)
        return None

    wsRelFechamento = wb["Rel. Fechamento de Caixa"]

    insert_terminal_values(wsRelFechamento, report)
    insert_total_values(wsRelFechamento, report)
    map_other_information(wsRelFechamento, report)
    insert_payment_methods(wsRelFechamento, report)
    compare_totals(wsRelFechamento)

    return wb