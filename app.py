import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd

st.set_page_config(layout="wide")
st.title("Diagnóstico de Estrutura DOCM")

uploaded_files = st.file_uploader(
    "Envie arquivos .docm",
    type=["docm"],
    accept_multiple_files=True
)

def extract_tables(file):

    with zipfile.ZipFile(file) as docm:
        xml_content = docm.read("word/document.xml")

    root = ET.fromstring(xml_content)

    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    tables = []

    for tbl in root.findall('.//w:tbl', ns):

        table_data = []

        for row in tbl.findall('.//w:tr', ns):

            row_data = []

            for cell in row.findall('.//w:tc', ns):

                texts = [
                    t.text for t in cell.findall('.//w:t', ns)
                    if t.text
                ]

                row_data.append(" ".join(texts))

            table_data.append(row_data)

        tables.append(table_data)

    return tables


if uploaded_files:

    for file in uploaded_files:

        st.header(file.name)

        tables = extract_tables(file)

        st.write(f"Tabelas encontradas: {len(tables)}")

        for i, table in enumerate(tables):

            st.subheader(f"Tabela {i}")

            df = pd.DataFrame(table)

            st.dataframe(df)
