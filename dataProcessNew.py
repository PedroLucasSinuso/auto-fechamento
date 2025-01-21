from bs4 import BeautifulSoup
import re

#Carregar o conteúdo HTML
def getSoup(htmlFile):
    return BeautifulSoup(htmlFile,'html.parser')

#Identificar os terminais
def getTerminals(soup):
    terminalTags = soup.find_all(string = re.compile(r'Terminal:\s*\d+'))
    return list({
        re.search(r'\d+',tag).group() 
        for tag in terminalTags 
        if re.search(r'\d+',tag)
    })

#Venda bruta por terminal
def grossSale(soup):
    gsTags = soup.find_all(string=re.compile(r'Venda\s*bruta'))
    return {
        re.search(r'\d+', tag.find_previous(string=re.compile(r'Terminal:\s*\d+'))).group(): 
        float(tag.find_next('div', {'class': 'font1'}).text.replace(',', '.'))
        for tag in gsTags
    }

#Valor total das trocas
def exchangedItems(soup):
    exchangedTags = soup.find_all(string=re.compile(r'Troca$'))
    exchanged = [
        float(tag.find_next('div', {'class': 'font1'}).text.replace(',', '.'))
        for tag in exchangedTags
        if re.match(r'\d+,\d+', tag.find_next('div', {'class': 'font1'}).text)
    ]
    return sum(exchanged)

#Valor do frete
#Obs: Maldita string - (negativo). Linx, eu te odeio por não usar codificação de gente
def shipping(soup):
    shippingTags = soup.find_all(string=re.compile(r'FRETE'))
    shippingPrices = [
        float(tag.find_next("div",{'class':'font1'}).text.replace('‑', '-').replace(',', '.'))
        for tag in shippingTags
    ]
    return abs(sum(shippingPrices))

#Movimentações de caixa
def cashMovement(soup):
    movTags = soup.find_all(string=re.compile(r'Movimentação\s*de\s*caixa'))
    movPositive,movNegative = [],[]
    for tag in movTags:
        price = float(tag.find_next('div',{'class':'font1'}).text.replace('−', '-').replace(',', '.'))
        if price >= 0: movPositive.append(price)
        else: movNegative.append(price)
    movPositiveTotal = sum(movPositive)
    movNegativeTotal = sum(movNegative)
    return {"movPositive":movPositiveTotal,"movNegative":movNegativeTotal}

#Valor total do Omnichannel
def omnichannel(soup):
    omniTags = soup.find_all(string=re.compile(r'OMNICHANNEL'))
    omniValues = [
        float(tag.find_next('div',{'class':'font1'}).text.replace(',', '.'))
        for tag in omniTags
    ]
    return sum(omniValues)

#Valores por tipo de pagamento
def paymentMethods(soup):
    paymentTags = soup.find_all(string=re.compile(r'CARTAO|DINHEIRO|QR'))
    payments = {}
    for tag in paymentTags:
        payType = tag.strip()
        value = float(tag.find_next('div',{'class':'font1'}).text.replace(',', '.'))
        payments[payType] = payments.get(payType,0) + value
    return payments
    
# Gerar relatório
def genReport(htmlFile):

    soup = getSoup(htmlFile)
    terminals = getTerminals(soup)
    gross = grossSale(soup)
    termsales_Dict = dict(zip(terminals,gross))
    exchangedPrice = exchangedItems(soup)
    shippingPrice = shipping(soup)
    cashMove = cashMovement(soup)
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

    Movimentações de caixa:
    - Positivas: R$ {cashMove["movPositive"]:.2f}
    - Negativas: R$ {cashMove["movNegative"]:.2f}

    Valor total do Omnichannel: R$ {omniPrice:.2f}

    Totais por tipo de pagamento:
    """
    for tipo, valor in paymentTypes.items():
        relatorio += f"- {tipo}: R$ {valor:.2f}\n"

    print(relatorio)
