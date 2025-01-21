from bs4 import BeautifulSoup
import re

# Carregar o conteúdo do arquivo HTML
with open('OUTPUT_HTML_POR-TERMINAL.HTML', 'r', encoding='windows-1252') as file:
    html_content = file.read()

# Analisar o HTML com BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')

# Identificar os terminais
terminal_tags = soup.find_all(string=re.compile(r'Terminal:\s*\d+'))
terminais = list({re.search(r'\d+', tag).group() for tag in terminal_tags if re.search(r'\d+', tag)})

# Venda bruta por terminal
venda_bruta_tags = soup.find_all(string=re.compile(r'Venda\s*bruta'))
vendas_brutas = [
    float(tag.find_next('div', {'class': 'font1'}).text.replace(',', '.'))
    for tag in venda_bruta_tags
]

# Número de trocas
troca_tags = soup.find_all(string=re.compile(r'Troca$'))
trocas = [
    float(tag.find_next('div', {'class': 'font1'}).text.replace(',', '.'))
    for tag in troca_tags
    if re.match(r'\d+,\d+', tag.find_next('div', {'class': 'font1'}).text)
]
soma_trocas = sum(trocas)

# Valor do frete
frete_tags = soup.find_all(string=re.compile(r'FRETE'))
frete_valores = [
    float(tag.find_next('div', {'class': 'font1'}).text.replace('‑', '-').replace(',', '.'))
    for tag in frete_tags
]
frete_total = abs(sum(frete_valores))

# Movimentações de caixa
mov_tags = soup.find_all(string=re.compile(r'Movimentação\s*de\s*caixa'))
mov_positivas = []
mov_negativas = []
for tag in mov_tags:
    valor = float(tag.find_next('div', {'class': 'font1'}).text.replace('−', '-').replace(',', '.'))
    if valor >= 0:
        mov_positivas.append(valor)
    else:
        mov_negativas.append(abs(valor))

mov_positivas_total = sum(mov_positivas)
mov_negativas_total = sum(mov_negativas)

# Valor total do Omnichannel
omni_tags = soup.find_all(string=re.compile(r'OMNICHANNEL'))
omni_valores = [
    float(tag.find_next('div', {'class': 'font1'}).text.replace(',', '.'))
    for tag in omni_tags
]
omni_total = sum(omni_valores)

# Totais por tipo de pagamento
pagamento_tags = soup.find_all(string=re.compile(r'CARTAO|DINHEIRO|QR'))
pagamentos = {}
for tag in pagamento_tags:
    tipo = tag.strip()
    valor = float(tag.find_next('div', {'class': 'font1'}).text.replace(',', '.'))
    pagamentos[tipo] = pagamentos.get(tipo, 0) + valor

# Gerar o relatório
relatorio = f"""
Relatório de Terminais
=====================
Número de terminais: {len(terminais)}

Venda bruta por terminal:
"""
for terminal, venda in zip(terminais, vendas_brutas):
    relatorio += f"- Terminal {terminal}: R$ {venda:.2f}\n"

relatorio += f"""
Número total de trocas: R$ {soma_trocas:.2f}

Frete total: R$ {frete_total:.2f}

Movimentações de caixa:
- Positivas: R$ {mov_positivas_total:.2f}
- Negativas: R$ {mov_negativas_total:.2f}

Valor total do Omnichannel: R$ {omni_total:.2f}

Totais por tipo de pagamento:
"""
for tipo, valor in pagamentos.items():
    relatorio += f"- {tipo}: R$ {valor:.2f}\n"

print(relatorio)
