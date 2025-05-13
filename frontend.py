import streamlit as st
from datetime import datetime
import pytz
from io import BytesIO

from infrastructure.config.settings import PAGE_CONFIG, STYLE_CONFIG, DATE_FORMAT
from infrastructure.adapters.docx_reader import DocxReader
from infrastructure.adapters.excel_writer import ExcelWriter
from infrastructure.adapters.streamlit_presenter import StreamlitPresenter
from core.use_cases.generate_report import ReportGenerator
from core.use_cases.process_sheet import SheetProcessor

# Configuração inicial
st.set_page_config(**PAGE_CONFIG)
st.markdown(STYLE_CONFIG, unsafe_allow_html=True)

# Data atual
today = datetime.now(pytz.timezone('America/Sao_Paulo'))
st.markdown(f"### Dia {today.strftime(DATE_FORMAT)}")

# Inicialização de componentes
docx_reader = DocxReader()
excel_writer = ExcelWriter()
presenter = StreamlitPresenter()
report_generator = ReportGenerator(docx_reader)
sheet_processor = SheetProcessor(excel_writer)

# Frontend do aplicativo
col1, col2 = st.columns(2)

with col1:
    word_file = st.file_uploader("Faça o upload do arquivo Word", type=["doc"])
    collect = st.checkbox("Coleta", value=False)
    change = st.checkbox("Troco")
    change_value = st.number_input("Valor do troco", label_visibility="collapsed", value=0.0, key='troco') if change else 0.0
    cash_fund = st.number_input("Fundo de caixa", value=0.0)

with col2:
    xlsm_file = st.file_uploader("Faça o upload do arquivo XLSM", type=["xlsm"])
    st.link_button("Como usar o App", "https://github.com/PedroLucasSinuso/auto-fechamento/tree/master#como-usar-o-aplicativo")

# Botões de ação
col3, col4 = st.columns(2)
with col3:
    if st.button("Gerar Relatório"):
        if not word_file or not xlsm_file or cash_fund == 0.0:
            st.warning("Por favor, envie os arquivos Word, XLSM e o valor do fundo de caixa antes de gerar o relatório.")
        else:
            try:
                # Geração do relatório
                report = report_generator.generate(word_file)
                
                # Processamento da planilha
                wb, is_match = sheet_processor.process(
                    xlsm_file,
                    report,
                    today,
                    collect,
                    change_value,
                    cash_fund
                )
                
                # Exibição dos resultados
                presenter.show_comparison_result(is_match)
                presenter.show_report_summary(report)
                
                
                # Preparação para download
                file_content, filename = presenter.prepare_download(wb, today)
                
                # Botão de download
                with col4:
                    presenter.get_download_button(file_content, filename)
                    
            except Exception as e:
                st.error(f"Erro ao gerar o relatório: {e}")
