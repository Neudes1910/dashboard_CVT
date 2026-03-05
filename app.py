import streamlit as st
from docx import Document
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

if uploaded_files:
    all_data = []

    for file in uploaded_files:
        try:
            doc = Document(file)
        except Exception as e:
            st.error(f"Erro ao abrir {file.name}: {e}")
            continue

        # Extrai texto de cada parágrafo
        text = "\n".join([p.text for p in doc.paragraphs])

        # Regex simples para extrair Projeto e Quantidade
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

        # Agrupando por projeto
        project_summary = data.groupby("Projeto")["Quantidade"].sum().reset_index()

        st.subheader("Gráfico por Projeto")
        fig = px.bar(project_summary, x="Projeto", y="Quantidade", text="Quantidade", color="Projeto")
        fig.update_layout(showlegend=False, xaxis_title="Projeto", yaxis_title="Total de Itens")
        st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Aguardando arquivos .docm para upload.")
