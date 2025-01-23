from bs4 import BeautifulSoup
import re

#Carregar o conteúdo HTML
def getSoup(htmlFile:bytes) -> BeautifulSoup:
    return BeautifulSoup(htmlFile,'html.parser')

#Identificar os terminais
def getTerminals(soup:BeautifulSoup) -> list:
    terminalTags = soup.find_all(string = re.compile(r'Terminal:\s*\d+'))
    return list({
        re.search(r'\d+',tag).group() 
        for tag in terminalTags 
        if re.search(r'\d+',tag)
    })

#Venda bruta por terminal
def grossSale(soup:BeautifulSoup) -> dict:
    gsTags = soup.find_all(string=re.compile(r'Venda\s*bruta'))
    return {
        re.search(r'\d+', tag.find_previous(string=re.compile(r'Terminal:\s*\d+'))).group(): 
        float(tag.find_next('div', {'class': 'font1'}).text.replace(',', '.'))
        for tag in gsTags
    }

# Soma do acréscimo no pagamento
def grossAdd(soup: BeautifulSoup) -> float:
    grossAddTags = soup.find_all(string=re.compile(r'Acréscimo\s*no\s*pagamento'))
    grossAdd = []
    for tag in grossAddTags:
        value_tag = tag.find_next('div', {'class': 'font1'})
        if value_tag:
            value = value_tag.text.replace(',', '.')
            try:
                grossAdd.append(float(value))
            except ValueError:
                continue
    return sum(grossAdd)

#Valor total das trocas
def exchangedItems(soup:BeautifulSoup) -> float:
    exchangedTags = soup.find_all(string=re.compile(r'Troca$'))
    exchanged = [
        float(tag.find_next('div', {'class': 'font1'}).text.replace(',', '.'))
        for tag in exchangedTags
        if re.match(r'\d+,\d+', tag.find_next('div', {'class': 'font1'}).text)
    ]
    return sum(exchanged)

#Valor do frete
#Obs: Maldita string - (negativo/hífen). Linx, eu te odeio por não usar codificação de gente
def shipping(soup:BeautifulSoup) -> float:
    shippingTags = soup.find_all(string=re.compile(r'FRETE'))
    shippingPrices = [
        float(tag.find_next("div",{'class':'font1'}).text.replace('‑', '-').replace(',', '.'))
        for tag in shippingTags
    ]
    return abs(sum(shippingPrices))

# Função para somar todas as movimentações de caixa negativas e retornar o módulo da soma
# Reescrever a função expenses
def expenses(soup: BeautifulSoup) -> float:
    movTags = soup.find_all(string=re.compile(r'Movimentação\s*de\s*caixa\s*−'))
    movNegative = [
        float(tag.find_next('div', {'class': 'font1'}).text.replace('−', '-').replace(',', '.'))
        for tag in movTags
        if float(tag.find_next('div', {'class': 'font1'}).text.replace('−', '-').replace(',', '.')) < 0
    ]
    return abs(sum(movNegative))

#Valor total do Credsystem
def credsystem(soup:BeautifulSoup) -> float:
    credyTags = soup.find_all(string=re.compile(r'CREDSYSTEM'))
    credyValues = [
        float(tag.find_next('div',{'class':'font1'}).text.replace(',', '.'))
        for tag in credyTags
    ]
    return sum(credyValues)

#Valor total do Omnichannel
def omnichannel(soup:BeautifulSoup) -> float:
    omniTags = soup.find_all(string=re.compile(r'OMNICHANNEL'))
    omniValues = [
        float(tag.find_next('div',{'class':'font1'}).text.replace(',', '.'))
        for tag in omniTags
    ]
    return sum(omniValues)

#Valores por tipo de pagamento
def paymentMethods(soup:BeautifulSoup) -> dict:
    paymentTags = soup.find_all(string=re.compile(r'CARTAO|DINHEIRO|QR'))
    payments = {}
    for tag in paymentTags:
        payType = tag.strip()
        value = float(tag.find_next('div',{'class':'font1'}).text.replace(',', '.'))
        payments[payType] = payments.get(payType,0) + value
    return payments
    
# Gerar relatório
def genReport(htmlFile:BeautifulSoup) -> dict:

    soup = getSoup(htmlFile)
    terminals = getTerminals(soup)
    gross = grossSale(soup)
    grossAddPrice = grossAdd(soup)
    exchangedPrice = exchangedItems(soup)
    shippingPrice = shipping(soup)
    expensesPrice = expenses(soup)
    credsystemPrice = credsystem(soup)
    omniPrice = omnichannel(soup)
    paymentTypes = paymentMethods(soup)
        
    relatorio = f"""
    Relatório de Terminais
    =====================
    Número de terminais: {len(terminals)}

    Venda bruta por terminal:
    """
    relatorio+="\n"
    for terminal, venda in gross.items():
        relatorio += f"- Terminal {terminal}: R$ {venda:.2f}\n"

    relatorio += f"""
    Número total de trocas: R$ {exchangedPrice:.2f}

    Frete total: R$ {shippingPrice:.2f}

    Valor total do Omnichannel: R$ {omniPrice:.2f}

    Totais por tipo de pagamento:
    """
    for tipo, valor in paymentTypes.items():
        relatorio += f"- {tipo}: R$ {valor:.2f}\n"

    print(relatorio)
    return {"Terminals":terminals,
            "Gross_Sales":gross,
            "Gross_Add":grossAddPrice,
            "Exchanged_Items":exchangedPrice,
            "Shipping":shippingPrice,
            "Expenses":expensesPrice,
            "Credsystem":credsystemPrice,
            "Omnichannel":omniPrice,
            "Payment_Methods":paymentTypes
        }

# Função para gerar o relatório em formato JSON
def genReportJson(htmlFile: bytes) -> str:
    import json
    report = genReport(htmlFile)
    return json.dumps(report, indent=4, ensure_ascii=False)