import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="Retrabalho HidroMeter Connect", layout="wide")
st.title("Total de Retrabalhos por Natureza - HidroMeter Connect")

uploaded_files = st.file_uploader(
    "Arraste e solte arquivos Word (.docm) aqui",
    type=["docm"],
    accept_multiple_files=True
)

def extract_table_from_docm(file, index=0):
    """
    Extrai a tabela de índice 'index' do arquivo .docm.
    index=0 para primeira tabela, index=1 para segunda, etc.
    """
    try:
        with zipfile.ZipFile(file) as docm:
            xml_content = docm.read("word/document.xml")
        root = ET.fromstring(xml_content)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        tables = root.findall('.//w:tbl', ns)
        if len(tables) <= index:
            return None  # não existe tabela no índice solicitado

        tbl = tables[index]
        table_data = []
        for row in tbl.findall('.//w:tr', ns):
            cells = []
            for cell in row.findall('.//w:tc', ns):
                texts = [t.text for t in cell.findall('.//w:t', ns) if t.text]
                cell_text = " ".join(texts).strip()
                cells.append(cell_text)
            if cells:
                table_data.append(cells)
        return table_data
    except Exception as e:
        st.error(f"Erro ao processar {file.name}: {e}")
        return None

def is_hidrometer_connect(table_data):
    """
    Verifica se a tabela contém o produto HidroMeter Connect.
    Procura linha do tipo: Produto --- HidroMeter Connect
    """
    for row in table_data:
        for cell in row:
            if re.match(r"Produto\s*---\s*HidroMeter Connect", cell, re.IGNORECASE):
                return True
    return False

if uploaded_files:
    retrabalho_list = []

    for file in uploaded_files:
        # Altere index=1 se o produto estiver na segunda tabela
        table_data = extract_table_from_docm(file, index=1)  # segunda tabela
        if not table_data:
            st.warning(f"Nenhuma tabela encontrada no índice especificado em {file.name}.")
            continue

        if not is_hidrometer_connect(table_data):
            st.info(f"O arquivo {file.name} não contém o produto HidroMeter Connect na tabela selecionada.")
            continue

        # Considera primeira linha como header
        df = pd.DataFrame(table_data[1:], columns=table_data[0])

        # Verifica se existem colunas "Natureza" e "Quantidade"
        if "Natureza" in df.columns and "Quantidade" in df.columns:
            df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors='coerce').fillna(0)
            retrabalho_list.append(df)
        else:
            st.warning(f"As colunas 'Natureza' ou 'Quantidade' não foram encontradas em {file.name}.")

    if retrabalho_list:
        retrabalho_df = pd.concat(retrabalho_list, ignore_index=True)

        # Agrupa por natureza
        summary = retrabalho_df.groupby("Natureza")["Quantidade"].sum().reset_index()

        st.subheader("Total de Retrabalhos por Natureza - HidroMeter Connect")
        st.dataframe(summary)

        fig = px.bar(
            summary,
            x="Natureza",
            y="Quantidade",
            text="Quantidade",
            color="Natureza",
            title="Retrabalhos por Natureza - HidroMeter Connect"
        )
        fig.update_layout(showlegend=False, xaxis_title="Natureza", yaxis_title="Total de Retrabalhos")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Nenhum retrabalho encontrado para HidroMeter Connect.")
else:
    st.info("Aguardando arquivos .docm para upload.")

