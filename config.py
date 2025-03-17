# Configurações da página
PAGE_CONFIG = {
    "page_title": "Auto-fechamento: Taco",
    "page_icon": ":bar_chart:",
    "layout": "wide",
    "initial_sidebar_state": "expanded",
}

# Estilo da página
STYLE_CONFIG = """
<div style='text-align: center;'><img src='https://taco.vtexassets.com/assets/vtex/assets-builder/taco.store-theme/3.0.4/img/logo-black___99045b7978d746cb5a089dde09b96c2e.png' width='200' />
</div>

<h1 style='text-align: center;'>
    Auto-fechamento
</h1>

<style>
.stApp {
    background-image: url("https://www.chromahouse.com/wp-content/uploads/2019/07/purpleandblue.png");
    background-size: cover;
    background-position: center;
    background-attachment: fixed;
}
</style>

<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
</style>
"""

# Formato de data
DATE_FORMAT = '%d/%m/%Y'

# Expressão regular para datas
DATE_PATTERN = r'\b\d{2}/\d{2}/\d{4}\b'

# Métodos de pagamento
PAYMENT_METHODS = {
    r'DINHEIRO': 'DINHEIRO',
    r'QR': 'QR',
    r'CARTAO\s+CREDITO\s+PDV': 'CARTAO CREDITO PDV',
    r'CARTAO\s+DEBITO\s+PDV': 'CARTAO DEBITO PDV',
    r'CARTAO\s+DE\s+CREDITO': 'CARTAO DE CREDITO',
    r'CARTAO\s+DE\s+DEBITO': 'CARTAO DE DEBITO',
    r'VALE\s+FUNCIONARIO': 'VALE FUNCIONARIO',
    r'VENDA\s+FATURADA': 'VENDA FATURADA',
    r'VOUCHER': 'VOUCHER',
    r'BOLETO': 'BOLETO',
}

# Mapeamento de métodos de pagamento para células
PAYMENT_MAPPINGS = {
    r'DINHEIRO': "G13",
    r'QR': "G18",
    r'CARTAO\s*CREDITO\s*PDV': "G14",
    r'CARTAO\s*DEBITO\s*PDV': "G15",
    r'CARTAO\s+DE\s+CREDITO': "G16",
    r'CARTAO\s+DE\s+DEBITO': "G17",
    r'VALE\s+FUNCIONARIO': "G21",
    r'VOUCHER': "G20",
    r'BOLETO': "G22"
}

# Padrões de entrada de caixa
CASH_INFLOW_PATTERNS = [
    r"ENTRADA DE TROCO",
    r"ENTRADA FUNDO DE CAIXA",
    r"ENTRADA NO CAIXA",
    r"SALDO INICIAL",
    r"VENDA FRANQUIA"
]

# Padrões de saída de caixa
CASH_OUTFLOW_PATTERNS = [
    r"BAINHAS E MATERIAL DE COSTURA",
    r"COLETA PROSEGUR",
    r"CRM BONUS",
    r"CÓPIAS/XEROX",
    r"DEPÓSITO DE TROCO",
    r"DESCONTO UNIFORME",
    r"DESP - BRINDES",
    r"DESP-OUTROS MAT. DE CONSU",
    r"DESPESA MATERIAL COPA",
    r"DESPESA NOTA FISCAL",
    r"DIFERENCA DE TROCA",
    r"ESTACIONAMENTO",
    r"HIGIENE E LIMPEZA",
    r"LANCHE",
    r"MANUTENÇÃO E REPAROS",
    r"MATERIAL ESCRITÓRIO E PAPELARIA",
    r"MEDICINA DO TRABALHO",
    r"PASSAGENS/CONDUÇÕES",
    r"PREMIAÇÃO CREDSYSTEM",
    r"PRÊMIOS E GRATIFICAÇÕES",
    r"RETIRADA DO CAIXA",
    r"SEGURANÇA E VIGILÂNCIA",
    r"SUPLEMENTO DE INFORMATICA",
    r"TREINAMENTO",
    r"UNIFORME",
    r"VIAGEM",
    r"VITRINE/MATERIAL DE VITRINE",
    r"ÁGUA"
]