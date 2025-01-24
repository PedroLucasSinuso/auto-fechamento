import streamlit as st
from datetime import date, datetime
from dataProcess import genReport
from dataWrite import sheetEdit
from io import BytesIO
import pytz
 

# Título do aplicativo
st.set_page_config(
    page_title="Auto-fechamento: Taco"
)
st.title("Auto-fechamento")
# Uploads de arquivos
wordFile = st.file_uploader("Faça o upload do arquivo Word", type=["doc"])
xlsmFile = st.file_uploader("Faça o upload do arquivo XLSM", type=["xlsm"])

xlsmFileDownload = None

# Coleta    
collect = st.checkbox("Coleta", value=False)

# Botão para gerar o relatório
if st.button("Ok"):

    # Gerar relatório
    reportDict = genReport(wordFile)

    # Editar planilha
    xlsmFileDownload = sheetEdit(xlsmFile, reportDict, collect=True) if collect else sheetEdit(xlsmFile, reportDict, collect=False)
    
    # Exibir resumo do relatório
    st.subheader("Resumo do Relatório")
    st.json(reportDict)

if xlsmFileDownload is not None:

    # Salvar o arquivo modificado em memória usando o BytesIO
    output = BytesIO()
    xlsmFileDownload.save(output)
    output.seek(0)
    file_content = output.read()

    # Renomeando o arquivo para download com a data atual do Brasil
    today = datetime.now(pytz.timezone('America/Sao_Paulo'))
    new_filename = f"FECHAMENTO {today.strftime('%d.%m.%Y')}.xlsm"
    
    # Disponibilizando para download
    st.download_button(
        label="Baixar planilha de fechamento",
        data=file_content,
        file_name=new_filename,
        mime="application/vnd.ms-excel.sheet.macroEnabled.12"
    )
