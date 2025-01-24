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
    layout="centered",
    initial_sidebar_state="expanded",
)

# Título do aplicativo
st.title("Auto-fechamento :bar_chart:")

# Data de hoje no formato dd/mm/yyyy
today = datetime.now(pytz.timezone('America/Sao_Paulo'))
st.markdown(f"### Dia {today.strftime('%d/%m/%Y')}")

# Divisão da página em duas colunas
col1, col2 = st.columns(2)

with col1:
    # Upload do arquivo Word
    wordFile = st.file_uploader("Faça o upload do arquivo Word", type=["doc"])

with col2:
    # Upload do arquivo XLSM
    xlsmFile = st.file_uploader("Faça o upload do arquivo XLSM", type=["xlsm"])

xlsmFileDownload = None

# Coleta
collect = st.checkbox("Coleta", value=False)

def display_report_summary(report):
    st.subheader("Resumo do Relatório")
    st.markdown("### Terminais")
    st.write(report['Terminals'])
    
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
        if not wordFile or not xlsmFile:
            st.warning("Por favor, envie os arquivos Word e XLSM antes de gerar o relatório.")
        else:
            # Gerar relatório
            reportDict = genReport(wordFile)

            # Editar planilha
            xlsmFileDownload = sheetEdit(xlsmFile, reportDict, today, collect)
            
            # Exibir resumo do relatório
            display_report_summary(reportDict)

with col4:
    if xlsmFileDownload is not None:
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
