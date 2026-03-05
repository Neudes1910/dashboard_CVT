import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Dashboard de Relatórios DOCM", layout="wide")
st.title("Dashboard de Relatórios DOCM com Tabelas")

uploaded_files = st.file_uploader(
    "Arraste e solte arquivos Word (.docm) aqui",
    type=["docm"],
    accept_multiple_files=True
)

def extract_tables_from_docm(file):
    tables_data = []
    try:
        with zipfile.ZipFile(file) as docm:
            xml_content = docm.read("word/document.xml")
        root = ET.fromstring(xml_content)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        for tbl in root.findall('.//w:tbl', ns):
            table = []
            for row in tbl.findall('.//w:tr', ns):
                cells = []
                for cell in row.findall('.//w:tc', ns):
                    texts = [t.text for t in cell.findall('.//w:t', ns) if t.text]
                    cell_text = " ".join(texts).strip()
                    cells.append(cell_text)
                if cells:
                    table.append(cells)
            if table:
                tables_data.append(table)
    except Exception as e:
        st.error(f"Erro ao processar {file.name}: {e}")
    
    return tables_data

def tables_to_dataframes(tables):
    dfs = {"Retrabalho": [], "Horas Indisponíveis": []}
    for table in tables:
        df = pd.DataFrame(table[1:], columns=table[0])  # primeira linha como header
        cols = [c.lower() for c in df.columns]

        if any("retrabalho" in c for c in cols):
            dfs["Retrabalho"].append(df)
        elif any("horas indisponíveis" in c or "indisponível" in c for c in cols):
            dfs["Horas Indisponíveis"].append(df)
    return dfs

if uploaded_files:
    retrabalho_list = []
    indisponiveis_list = []

    for file in uploaded_files:
        tables = extract_tables_from_docm(file)
        dfs = tables_to_dataframes(tables)
        retrabalho_list.extend(dfs["Retrabalho"])
        indisponiveis_list.extend(dfs["Horas Indisponíveis"])

    if retrabalho_list:
        retrabalho_df = pd.concat(retrabalho_list, ignore_index=True)
        st.subheader("Dados de Retrabalho")
        st.dataframe(retrabalho_df)

        if "Projeto" in retrabalho_df.columns and "Quantidade" in retrabalho_df.columns:
            summary = retrabalho_df.groupby("Projeto")["Quantidade"].sum().reset_index()
            fig = px.bar(summary, x="Projeto", y="Quantidade", text="Quantidade", color="Projeto",
                         title="Retrabalho por Projeto")
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
    
    if indisponiveis_list:
        indisponiveis_df = pd.concat(indisponiveis_list, ignore_index=True)
        st.subheader("Dados de Horas Indisponíveis")
        st.dataframe(indisponiveis_df)

        if "Projeto" in indisponiveis_df.columns and "Horas" in indisponiveis_df.columns:
            summary = indisponiveis_df.groupby("Projeto")["Horas"].sum().reset_index()
            fig = px.bar(summary, x="Projeto", y="Horas", text="Horas", color="Projeto",
                         title="Horas Indisponíveis por Projeto")
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Aguardando arquivos .docm para upload.")
