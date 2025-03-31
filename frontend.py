import streamlit as st
from datetime import datetime
from dataProcess import genReport
from dataWrite import sheetEdit
from io import BytesIO
import pytz
from config import PAGE_CONFIG, STYLE_CONFIG, DATE_FORMAT

# Configuração da página
st.set_page_config(**PAGE_CONFIG)

# Estilo da página
st.markdown(STYLE_CONFIG, unsafe_allow_html=True)

# Data de hoje no formato dd/mm/yyyy
today = datetime.now(pytz.timezone('America/Sao_Paulo'))
st.markdown(f"### Dia {today.strftime(DATE_FORMAT)}")

# Divisão da página em duas colunas
col1, col2 = st.columns(2)

with col1:
    # Upload do arquivo Word
    word_file = st.file_uploader("Faça o upload do arquivo Word", type=["doc"])
    
    # Coleta
    collect = st.checkbox("Coleta", value=False)
    
    # Troco  
    change = st.checkbox("Troco")
    change_value = st.number_input("Valor do troco", label_visibility="collapsed", value=0.0, key='troco') if change else 0.0
    
    # Fundo de caixa
    cash_fund = st.number_input("Fundo de caixa", value=0.0)

with col2:
    # Upload do arquivo XLSM
    xlsm_file = st.file_uploader("Faça o upload do arquivo XLSM", type=["xlsm"])
    
    # Link "Como usar"
    st.link_button("Como usar o App", "https://github.com/PedroLucasSinuso/auto-fechamento/tree/master#como-usar-o-aplicativo")

xlsm_file_download = None

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
        if not word_file or not xlsm_file or cash_fund == 0.0:
            st.warning("Por favor, envie os arquivos Word, XLSM e o valor do fundo de caixa antes de gerar o relatório.")
        else:
            try:
                # Gerar relatório
                report_dict = genReport(word_file)
                # Editar planilha
                xlsm_file_download = sheetEdit(xlsm_file, report_dict, today, collect, change_value, cash_fund)
                # Exibir resumo do relatório
                display_report_summary(report_dict)
            except Exception as e:
                st.error(f"Erro ao gerar o relatório: {e}")

with col4:
    if xlsm_file_download is not None:
        try:
            # Salvar o arquivo modificado em memória usando o BytesIO
            output = BytesIO()
            xlsm_file_download.save(output)
            output.seek(0)
            file_content = output.read()

            # Renomeando o arquivo para download com a data atual do Brasil
            new_filename = f"FECHAMENTO {today.strftime(DATE_FORMAT.replace('/', '.'))}.xlsm"
            
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