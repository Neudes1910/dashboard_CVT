import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Análise HidroMeter Connect", layout="wide")
st.title("Análise de Ocorrências - HidroMeter Connect")

uploaded_files = st.file_uploader(
    "Arraste arquivos .docm",
    type=["docm"],
    accept_multiple_files=True
)

def extract_docm_content(file):

    try:
        with zipfile.ZipFile(file) as docm:
            xml_content = docm.read("word/document.xml")

        root = ET.fromstring(xml_content)

        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        # texto completo
        texts = [node.text for node in root.findall('.//w:t', ns) if node.text]
        full_text = " ".join(texts)

        # tabelas
        tables = []
        for tbl in root.findall('.//w:tbl', ns):

            table_data = []

            for row in tbl.findall('.//w:tr', ns):

                cells = []

                for cell in row.findall('.//w:tc', ns):

                    cell_text = " ".join(
                        t.text for t in cell.findall('.//w:t', ns) if t.text
                    ).strip()

                    cells.append(cell_text)

                if cells:
                    table_data.append(cells)

            if table_data:
                tables.append(table_data)

        return full_text, tables

    except Exception as e:
        st.error(f"Erro ao processar {file.name}: {e}")
        return "", []


def extract_records_from_tables(tables):

    registros = []

    for table in tables:

        if len(table) < 2:
            continue

        df = pd.DataFrame(table[1:], columns=table[0])

        cols = [c.lower() for c in df.columns]

        natureza_col = None
        equipamento_col = None
        horas_col = None

        for i, c in enumerate(cols):

            if "natureza" in c:
                natureza_col = df.columns[i]

            if "equipamento" in c:
                equipamento_col = df.columns[i]

            if "hora" in c:
                horas_col = df.columns[i]

        if natureza_col and equipamento_col and horas_col:

            df[horas_col] = pd.to_numeric(df[horas_col], errors='coerce')

            for _, row in df.iterrows():

                registros.append({
                    "Natureza": row[natureza_col],
                    "Equipamento": row[equipamento_col],
                    "Horas Indisponíveis": row[horas_col]
                })

    return registros


if uploaded_files:

    todos_registros = []

    for file in uploaded_files:

        full_text, tables = extract_docm_content(file)

        if "hidrometer connect" not in full_text.lower():
            st.info(f"{file.name} não contém HidroMeter Connect")
            continue

        registros = extract_records_from_tables(tables)

        todos_registros.extend(registros)

    if todos_registros:

        df = pd.DataFrame(todos_registros)

        st.subheader("Dados extraídos")
        st.dataframe(df)

        ocorrencias = df.groupby("Natureza").size().reset_index(name="Ocorrências")

        total_horas = df["Horas Indisponíveis"].sum()

        horas_equip = df.groupby("Equipamento")["Horas Indisponíveis"].sum().reset_index()

        st.subheader("Ocorrências por Natureza")
        st.dataframe(ocorrencias)

        st.metric("Total de Horas Indisponíveis", round(total_horas, 2))

        st.subheader("Horas por Equipamento")
        st.dataframe(horas_equip)

        fig1 = px.bar(
            ocorrencias,
            x="Natureza",
            y="Ocorrências",
            text="Ocorrências",
            title="Ocorrências por Natureza"
        )

        st.plotly_chart(fig1, use_container_width=True)

        fig2 = px.bar(
            horas_equip,
            x="Equipamento",
            y="Horas Indisponíveis",
            text="Horas Indisponíveis",
            title="Horas por Equipamento"
        )

        st.plotly_chart(fig2, use_container_width=True)

    else:
        st.warning("Nenhum registro encontrado nas tabelas.")

else:
    st.info("Aguardando upload de arquivos.")
        st.warning("Nenhum registro encontrado.")

else:
    st.info("Aguardando upload de arquivos.")

