import streamlit as st
from io import BytesIO
from datetime import datetime
from core.interfaces.report_presenter import ReportPresenter
from core.entities.report import Report
from infrastructure.config.settings import DATE_FORMAT

class StreamlitPresenter(ReportPresenter):
    def show_report_summary(self, report: Report):
        st.markdown("<h3 style='text-align: center;'>Resumo do Relatório</h3>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"### Terminais: {', '.join(str(t) for t in report.terminals)}")
            st.markdown("### Vendas Brutas por Terminal")
            for terminal, value in report.gross_sales.items():
                st.write(f"Terminal {terminal}: R$ {value:.2f}")
            
            st.markdown("### Acréscimos e Descontos")
            st.write(f"Acréscimos: R$ {report.gross_add:.2f}")
            st.write(f"Descontos: R$ {report.discounts:.2f}")
            
            st.markdown("### Trocas")
            st.write(f"Total de Trocas: R$ {report.exchanged_items:.2f}")
            
            st.markdown("### Frete")
            st.write(f"Frete: R$ {report.shipping:.2f}")
        
        with col2:
            st.markdown("### Movimentações de Caixa")
            st.write(f"Entradas de Caixa: R$ {report.total_cash_inflow:.2f}")
            st.write(f"Saídas de Caixa: R$ {report.total_cash_outflow:.2f}")
            
            st.markdown("### Credsystem")
            st.write(f"Credsystem: R$ {report.credsystem:.2f}")
            
            st.markdown("### Omnichannel")
            st.write(f"Omnichannel: R$ {report.omnichannel:.2f}")
            
            st.markdown("### Métodos de Pagamento")
            for method, value in report.payment_methods.items():
                st.write(f"{method}: R$ {value:.2f}")
    
    def show_comparison_result(self, is_match: bool):
        if is_match:
            st.success("Bateu! :sunglasses:")
        else:
            st.error("Divergente. Confere a tabela por favor :skull:")
    
    def get_download_button(self, file_content: BytesIO, filename: str):
        return st.download_button(
            label="Baixar planilha de fechamento",
            data=file_content,
            file_name=filename,
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            help="Clique para baixar a planilha de fechamento com os dados atualizados."
        )
    
    def prepare_download(self, wb, day: datetime) -> tuple[BytesIO, str]:
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        file_content = output.read()
        filename = f"FECHAMENTO {day.strftime(DATE_FORMAT.replace('/', '.'))}.xlsm"
        return file_content, filename