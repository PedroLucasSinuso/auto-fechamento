import aspose.words as aw
import re

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
        'Expenses': 0,
        'Credsystem': 0,
        'Omnichannel': 0,
        'Payment_Methods': {}
    }
    for i, line in enumerate(lines):
        
        if re.search(r'Terminal', line):
            # Inserir o número do terminal ao relatório se ele não existir
            terminal = re.search(r'\d+', line).group()
            if terminal not in report['Terminals']: 
                report['Terminals'].append(terminal)
                # Inserir venda bruta do terminal ao relatório
                report['Gross_Sales'][terminal] = float(re.search(r'\d+,\d+', lines[i+2]).group().replace(',', '.'))
                # Inserir trocas do terminal ao relatório
                report['Exchanged_Items'] += float(re.search(r'\d+,\d+', lines[i+8]).group().replace(',', '.'))
                # Inserir acréscimos do terminal ao relatório
                report['Gross_Add'] += float(re.search(r'\d+,\d+', lines[i+4]).group().replace(',', '.'))
                report['Gross_Add'] += float(re.search(r'\d+,\d+', lines[i+6]).group().replace(',', '.'))
                # Inserir descontos do terminal ao relatório
                report['Discounts'] -= float(re.search(r'\d+,\d+', lines[i+3]).group().replace(',', '.'))
                report['Discounts'] -= float(re.search(r'\d+,\d+', lines[i+5]).group().replace(',', '.'))
        
        
        # Inserir frete ao relatório se existir na linha
        if re.search(r'FRETE\s*B2C', line): 
            report['Shipping'] += float(re.search(r'\d+,\d+', line).group().replace(',', '.'))

        # Inserir credsystem ao relatório se existir na linha
        if re.search(r'CREDSYSTEM', line): 
            report['Credsystem'] += float(re.search(r'\d+,\d+', line).group().replace(',', '.'))
        
        # Inserir omnichannel ao relatório se existir na linha
        if re.search(r'OMNICHANNEL', line): 
            report['Omnichannel'] += float(re.search(r'\d+,\d+', line).group().replace(',', '.'))
        
        # Identificar as linhas que contém métodos de pagamento e inserir ao relatório {"Método de pagamento":valor}
        if re.search(r'DINHEIRO|QR|CARTAO\s*CREDITO\s*PDV|CARTAO\s*DEBITO\s*PDV|CARTAO\s*CREDITO$|CARTAO\s*DEBITO$', line):
            payType = re.search(r'DINHEIRO|QR|CARTAO\s*CREDITO\s*PDV|CARTAO\s*DEBITO\s*PDV|CARTAO\s*CREDITO$|CARTAO\s*DEBITO$', line).group()
            value = float(re.search(r'\d+,\d+', line).group().replace(',', '.'))
            # Se pagamento for dinheiro, dividir por 2
            if payType == 'DINHEIRO': value /= 2
            report['Payment_Methods'][payType] = report['Payment_Methods'].get(payType,0) + value
    
    return report

if __name__ == "__main__":
    print(genReport("output.doc"))