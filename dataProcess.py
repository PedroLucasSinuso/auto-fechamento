import aspose.words as aw
import re
import streamlit as st

def genReport(word_file) -> dict:
    
    doc = aw.Document(word_file)
    lines = [p.get_text().strip() for p in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True) if p.get_text().strip()]
    
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
    for i, line in enumerate(lines):
        try:
            if re.search(r'Terminal', line):
                # Inserir o número do terminal ao relatório se ele não existir
                terminal = re.search(r'\d+', line).group()
                if terminal not in report['Terminals']: 
                    report['Terminals'].append(terminal)
                    try:
                        # Inserir venda bruta do terminal ao relatório
                        gross_sales_match = re.search(r'\d+,\d+', lines[i+2])
                        if gross_sales_match:
                            report['Gross_Sales'][terminal] = float(gross_sales_match.group().replace(',', '.'))
                    except Exception as e:
                        error_message = f"Erro ao inserir venda bruta do terminal {terminal}: {e}"
                        print(error_message)
                        st.warning(error_message)
                    try:
                        # Inserir trocas do terminal ao relatório
                        exchanged_items_match = re.search(r'\d+,\d+', lines[i+8])
                        if exchanged_items_match:
                            report['Exchanged_Items'] += float(exchanged_items_match.group().replace(',', '.'))
                    except Exception as e:
                        error_message = f"Erro ao inserir trocas do terminal {terminal}: {e}"
                        print(error_message)
                        st.warning(error_message)
                    try:
                        # Inserir acréscimos do terminal ao relatório
                        gross_add_match_1 = re.search(r'\d+,\d+', lines[i+4])
                        gross_add_match_2 = re.search(r'\d+,\d+', lines[i+6])
                        if gross_add_match_1:
                            report['Gross_Add'] += float(gross_add_match_1.group().replace(',', '.'))
                        if gross_add_match_2:
                            report['Gross_Add'] += float(gross_add_match_2.group().replace(',', '.'))
                    except Exception as e:
                        error_message = f"Erro ao inserir acréscimos do terminal {terminal}: {e}"
                        print(error_message)
                        st.warning(error_message)
                    try:
                        # Inserir descontos do terminal ao relatório
                        discounts_match_1 = re.search(r'\d+,\d+', lines[i+3])
                        discounts_match_2 = re.search(r'\d+,\d+', lines[i+5])
                        if discounts_match_1:
                            report['Discounts'] += float(discounts_match_1.group().replace(',', '.'))
                        if discounts_match_2:
                            report['Discounts'] += float(discounts_match_2.group().replace(',', '.'))
                    except Exception as e:
                        error_message = f"Erro ao inserir descontos do terminal {terminal}: {e}"
                        print(error_message)
                        st.warning(error_message)
            
            try:
                # Inserir frete ao relatório se existir na linha
                if re.search(r'FRETE\s*B2C', line): 
                    shipping_match = re.search(r'\d+,\d+', line)
                    if shipping_match:
                        report['Shipping'] += float(shipping_match.group().replace(',', '.'))
            except Exception as e:
                error_message = f"Erro ao inserir frete: {e}"
                print(error_message)
                st.warning(error_message)

            try:
                # Inserir omnichannel ao relatório se existir na linha
                if re.search(r'OMNICHANNEL', line): 
                    omnichannel_match = re.search(r'\d+,\d+', line)
                    if omnichannel_match:
                        report['Omnichannel'] += float(omnichannel_match.group().replace(',', '.'))
            except Exception as e:
                error_message = f"Erro ao inserir omnichannel: {e}"
                print(error_message)
                st.warning(error_message)
            
            try:
                # Inserir credsystem ao relatório se existir na linha
                if re.search(r'CREDSYSTEM', line): 
                    credsystem_match = re.search(r'\d+,\d+', line)
                    if credsystem_match:
                        report['Credsystem'] += float(credsystem_match.group().replace(',', '.'))
            except Exception as e:
                error_message = f"Erro ao inserir credsystem: {e}"
                print(error_message)
                st.warning(error_message)

            try:
                # Identificar as linhas que contém métodos de pagamento e inserir ao relatório {"Método de pagamento":valor}
                if re.search(r'DINHEIRO|QR|CARTAO\s*CREDITO\s*PDV|CARTAO\s*DEBITO\s*PDV|CARTAO\s*CREDITO$|CARTAO\s*DEBITO$', line):
                    payType = re.search(r'DINHEIRO|QR|CARTAO\s*CREDITO\s*PDV|CARTAO\s*DEBITO\s*PDV|CARTAO\s*CREDITO$|CARTAO\s*DEBITO$', line).group()
                    payment_match = re.search(r'\d+,\d+', line)
                    if payment_match:
                        value = float(payment_match.group().replace(',', '.'))
                        # Se pagamento for dinheiro, dividir por 2
                        if payType == 'DINHEIRO': value /= 2
                        report['Payment_Methods'][payType] = report['Payment_Methods'].get(payType,0) + value
            except Exception as e:
                error_message = f"Erro ao inserir método de pagamento {payType}: {e}"
                print(error_message)
                st.warning(error_message)

            # Mapear entradas de caixa
            cashInflow = [
                "ENTRADA DE TROCO",
                "ENTRADA FUNDO DE CAIXA",
                "ENTRADA NO CAIXA",
                "SALDO INICIAL",
                "VENDA FRANQUIA"
            ]

            # Mapear saídas de caixa
            cashOutflow = [
                "BAINHAS E MATERIAL DE COSTURA",
                "COLETA PROSEGUR",
                "CRM BONUS",
                "CÓPIAS/XEROX",
                "DEPÓSITO DE TROCO",
                "DESCONTO UNIFORME",
                "DESP - BRINDES",
                "DESP-OUTROS MAT. DE CONSU",
                "DESPESA MATERIAL COPA",
                "DESPESA NOTA FISCAL",
                "DIFERENCA DE TROCA",
                "ESTACIONAMENTO",
                "HIGIENE E LIMPEZA",
                "LANCHE",
                "MANUTENÇÃO E REPAROS",
                "MATERIAL ESCRITÓRIO E PAPELARIA",
                "MEDICINA DO TRABALHO",
                "PASSAGENS/CONDUÇÕES",
                "PREMIAÇÃO CREDSYSTEM",
                "PRÊMIOS E GRATIFICAÇÕES",
                "RETIRADA DO CAIXA",
                "SEGURANÇA E VIGILÂNCIA",
                "SUPLEMENTO DE INFORMATICA",
                "TREINAMENTO",
                "UNIFORME",
                "VIAGEM",
                "VITRINE/MATERIAL DE VITRINE",
                "ÁGUA"
            ]
            
            # Inserir entradas de caixa ao relatório
            for pattern in cashInflow:
                try:
                    if re.search(pattern, line):
                        cash_inflow_match = re.search(r'\d+,\d+', line)
                        if cash_inflow_match:
                            value = float(cash_inflow_match.group().replace(',', '.'))
                            report['Cash_Inflow'][pattern] = value
                            report['Total_Cash_Inflow'] += value
                except Exception as e:
                    error_message = f"Erro ao inserir entrada de caixa {pattern}: {e}"
                    print(error_message)
                    st.warning(error_message)
            
            # Inserir saídas de caixa ao relatório
            for pattern in cashOutflow:
                try:
                    if re.search(pattern, line):
                        cash_outflow_match = re.search(r'\d+,\d+', line)
                        if cash_outflow_match:
                            value = float(cash_outflow_match.group().replace(',', '.'))
                            report['Cash_Outflow'][pattern] = value
                            report['Total_Cash_Outflow'] += value
                except Exception as e:
                    error_message = f"Erro ao inserir saída de caixa {pattern}: {e}"
                    print(error_message)
                    st.warning(error_message)

        except (AttributeError, ValueError) as e:
            error_message = f"Erro ao processar a linha: {line}"
            print(error_message)
            st.warning(error_message)
    
    return report

if __name__ == "__main__":
    print(genReport("output.doc"))