import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="Contagem de Arquivos Word (.docm)", layout="wide")
st.title("Dashboard de Contagem de Arquivos Word com Macro")

uploaded_files = st.file_uploader(
    "Arraste e solte arquivos Word (.docm) aqui",
    type=["docm"],
    accept_multiple_files=True
)

def extract_text_from_docm(file):
    try:
        with zipfile.ZipFile(file) as docm:
            xml_content = docm.read("word/document.xml")
        root = ET.fromstring(xml_content)
        namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        paragraphs = [node.text for node in root.findall('.//w:t', namespaces)]
        return "\n".join(filter(None, paragraphs))
    except Exception as e:
        st.error(f"Erro ao processar {file.name}: {e}")
        return ""

if uploaded_files:
    all_data = []

    for file in uploaded_files:
        text = extract_text_from_docm(file)
        if not text:
            continue

        # Regex simples para capturar Projeto e Quantidade
        matches = re.findall(r"Projeto:\s*(.+)\nQuantidade:\s*(\d+)", text)
        if not matches:
            st.warning(f"Nenhum dado encontrado no arquivo {file.name}.")
            continue

        df = pd.DataFrame(matches, columns=["Projeto", "Quantidade"])
        df["Quantidade"] = df["Quantidade"].astype(int)
        all_data.append(df)

    if all_data:
        data = pd.concat(all_data, ignore_index=True)
        st.subheader("Dados Consolidados")
        st.dataframe(data)

        project_summary = data.groupby("Projeto")["Quantidade"].sum().reset_index()
        st.subheader("Gráfico por Projeto")
        fig = px.bar(project_summary, x="Projeto", y="Quantidade", text="Quantidade", color="Projeto")
        fig.update_layout(showlegend=False, xaxis_title="Projeto", yaxis_title="Total de Itens")
        st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Aguardando arquivos .docm para upload.")
