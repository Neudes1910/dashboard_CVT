import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Retrabalho HidroMeter Connect", layout="wide")
st.title("Total de Retrabalhos por Natureza - HidroMeter Connect")

uploaded_files = st.file_uploader(
    "Arraste e solte arquivos Word (.docm) aqui",
    type=["docm"],
    accept_multiple_files=True
)

def extract_first_table_from_docm(file):
    """Extrai apenas a primeira tabela do arquivo .docm"""
    try:
        with zipfile.ZipFile(file) as docm:
            xml_content = docm.read("word/document.xml")
        root = ET.fromstring(xml_content)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        tbl = root.find('.//w:tbl', ns)  # primeira tabela
        if not tbl:
            return None

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

if uploaded_files:
    retrabalho_list = []

    for file in uploaded_files:
        table_data = extract_first_table_from_docm(file)
        if not table_data:
            st.warning(f"Nenhuma tabela encontrada em {file.name}.")
            continue

        df = pd.DataFrame(table_data[1:], columns=table_data[0])  # primeira linha como header

        # Filtra apenas o produto HidroMeter Connect
        if "Produto" in df.columns:
            df_produto = df[df["Produto"].str.strip() == "HidroMeter Connect"]
            if "Natureza" in df_produto.columns and "Quantidade" in df_produto.columns:
                df_produto["Quantidade"] = pd.to_numeric(df_produto["Quantidade"], errors='coerce').fillna(0)
                retrabalho_list.append(df_produto)
            else:
                st.warning(f"As colunas 'Natureza' ou 'Quantidade' não foram encontradas em {file.name}.")
        else:
            st.warning(f"A coluna 'Produto' não foi encontrada em {file.name}.")

    if retrabalho_list:
        retrabalho_df = pd.concat(retrabalho_list, ignore_index=True)

        # Agrupa por natureza e soma a quantidade
        summary = retrabalho_df.groupby("Natureza")["Quantidade"].sum().reset_index()

        st.subheader("Total de Retrabalhos por Natureza - HidroMeter Connect")
        st.dataframe(summary)

        fig = px.bar(summary, x="Natureza", y="Quantidade", text="Quantidade", color="Natureza",
                     title="Retrabalhos por Natureza - HidroMeter Connect")
        fig.update_layout(showlegend=False, xaxis_title="Natureza", yaxis_title="Total de Retrabalhos")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Nenhum retrabalho encontrado para HidroMeter Connect.")
else:
    st.info("Aguardando arquivos .docm para upload.")
