import streamlit as st
from datetime import date
from dataProcess import genReport
from dataWrite import sheetEdit
from io import BytesIO
 

# Título do aplicativo
st.set_page_config(
    page_title="Auto-fechamento: Taco",
    page_icon="🌟",
)
st.title("Auto-fechamento")
# Uploads de arquivos
htmlFile = st.file_uploader("Faça o upload do arquivo HTML", type=["html"])
xlsmFile = st.file_uploader("Faça o upload do arquivo XLSM", type=["xlsm"])

xlsmFileDownload = None

# Coleta    
collect = st.checkbox("Coleta", value=False)

# Botão para gerar o relatório
if st.button("Ok"):
    reportDict = genReport(htmlFile)
    print(reportDict)
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

    # Renomeando o arquivo para download com a data atual
    today = date.today()
    new_filename = f"FECHAMENTO {today.strftime('%d.%m.%Y')}.xlsm"
    
    # Disponibilizando para download
    st.download_button(
        label="Baixar arquivo com novo nome",
        data=file_content,
        file_name=new_filename,
        mime="application/vnd.ms-excel.sheet.macroEnabled.12"
    )
