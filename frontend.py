import streamlit as st
from datetime import date
from dataProcessNew import genReport

# Título do aplicativo
st.title("Auto-fechamento")

# Upload do arquivo
uploaded_file = st.file_uploader("Faça o upload do arquivo HTML", type=["html"])
if st.button("Ok"):
    genReport(uploaded_file)
    

if uploaded_file is not None:

    # Lendo o conteúdo do arquivo como bytes
    file_content = uploaded_file.read()

    # Renomeando o arquivo para download com a data atual
    today = date.today()
    new_filename = f"FECHAMENTO{today.strftime("%d.%m.%Y")}.html"
    
    # Disponibilizando para download
    st.download_button(
        label="Baixar arquivo com novo nome",
        data=file_content,
        file_name=new_filename,
        mime="text/html"
    )
