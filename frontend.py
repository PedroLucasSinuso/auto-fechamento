import streamlit as st
from datetime import date, datetime
from dataProcess import genReport
from dataWrite import sheetEdit
from io import BytesIO
import pytz

# Configuração da página
st.set_page_config(
    page_title="Auto-fechamento: Taco",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilo da página
st.markdown(
    """
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
    """,
    unsafe_allow_html=True
)


# Data de hoje no formato dd/mm/yyyy
today = datetime.now(pytz.timezone('America/Sao_Paulo'))
st.markdown(f"### Dia {today.strftime('%d/%m/%Y')}")

# Divisão da página em duas colunas
col1, col2 = st.columns(2)

with col1:
    # Upload do arquivo Word
    wordFile = st.file_uploader("Faça o upload do arquivo Word", type=["doc"])
    
    # Coleta
    collect = st.checkbox("Coleta", value=False)
    
    # Troco  
    change = st.checkbox("Troco")
    if change: st.number_input("Valor do troco",label_visibility="collapsed", value=0.0,key='troco')
    else: st.session_state['troco'] = 0.0
    changeValue = st.session_state["troco"]

    # Fundo de caixa
    cashFund = st.number_input("Fundo de caixa", value=0.0)
    st.session_state['cashFund'] = cashFund
    
with col2:
    # Upload do arquivo XLSM
    xlsmFile = st.file_uploader("Faça o upload do arquivo XLSM", type=["xlsm"])

xlsmFileDownload = None

# Função para exibir o resumo do relatório
def display_report_summary(report):
    st.markdown("<h3 style='text-align: center;'>Resumo do Relatório</h3>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"### Terminais: {', '.join(str(t) for t in report['Terminals'])}")

        st.markdown("### Vendas Brutas por Terminal")
        for terminal, value in report['Gross_Sales'].items():
            st.write(f"Terminal {terminal}: R$ {value:.2f}")
        
        st.markdown("### Acréscimos e Descontos")
        st.write(f"Acréscimos: R$ {report['Gross_Add']:.2f}")
        st.write(f"Descontos: R$ {report['Discounts']:.2f}")
        
        st.markdown("### Trocas")
        st.write(f"Total de Trocas: R$ {report['Exchanged_Items']:.2f}")
        
        st.markdown("### Frete")
        st.write(f"Frete: R$ {report['Shipping']:.2f}")
    
    with col2:
        st.markdown("### Movimentações de Caixa")
        st.write(f"Entradas de Caixa: R$ {report['Total_Cash_Inflow']:.2f}")
        st.write(f"Saídas de Caixa: R$ {report['Total_Cash_Outflow']:.2f}")
        
        st.markdown("### Credsystem")
        st.write(f"Credsystem: R$ {report['Credsystem']:.2f}")
        
        st.markdown("### Omnichannel")
        st.write(f"Omnichannel: R$ {report['Omnichannel']:.2f}")
        
        st.markdown("### Métodos de Pagamento")
        for method, value in report['Payment_Methods'].items():
            st.write(f"{method}: R$ {value:.2f}")

# Divisão da página em duas colunas para os botões
col3, col4 = st.columns(2)
with col3:
    # Botão para gerar o relatório
    if st.button("Gerar Relatório"):
        if not wordFile or not xlsmFile or cashFund == 0.0:
            st.warning("Por favor, envie os arquivos Word, XLSM e o valor do fundo de caixa antes de gerar o relatório.")
        else:
            try:
                # Gerar relatório
                reportDict = genReport(wordFile)
                # Editar planilha
                xlsmFileDownload = sheetEdit(xlsmFile, reportDict, today, collect, changeValue, cashFund)
                # Exibir resumo do relatório
                display_report_summary(reportDict)
            except Exception as e:
                st.error(f"Erro ao gerar o relatório: {e}")

with col4:
    if xlsmFileDownload is not None:
        try:
            # Salvar o arquivo modificado em memória usando o BytesIO
            output = BytesIO()
            xlsmFileDownload.save(output)
            output.seek(0)
            file_content = output.read()

            # Renomeando o arquivo para download com a data atual do Brasil
            new_filename = f"FECHAMENTO {today.strftime('%d.%m.%Y')}.xlsm"
            
            # Disponibilizando para download
            st.download_button(
                label="Baixar planilha de fechamento",
                data=file_content,
                file_name=new_filename,
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                help="Clique para baixar a planilha de fechamento com os dados atualizados."
            )
        except Exception as e:
            st.error(f"Erro ao salvar o arquivo: {e}")
