import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="Retrabalhos HidroMeter Connect", layout="wide")
st.title("Análise de Ocorrências - HidroMeter Connect")

uploaded_files = st.file_uploader(
    "Arraste arquivos .docm",
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

            cells = []

            for cell in row.findall('.//w:tc', ns):

                texts = [
                    t.text for t in cell.findall('.//w:t', ns)
                    if t.text
                ]

                cells.append(" ".join(texts).strip())

            if cells:
                table_data.append(cells)

        if table_data:
            tables.append(table_data)

    return tables


if uploaded_files:

    registros = []

    for file in uploaded_files:

        tables = extract_tables(file)

        campos = {}

        # transforma tabela em dicionário campo -> valor
        for table in tables:

            for row in table:

                if len(row) >= 2:

                    chave = re.sub(r'\s+', ' ', row[0]).strip().lower()
                    valor = re.sub(r'\s+', ' ', row[1]).strip()

                    campos[chave] = valor

        # verifica produto
        produto = campos.get("produto", "").lower()

        if "hidrometer" not in produto:
            continue

        natureza = campos.get("natureza")
        equipamento = campos.get("equipamento")
        horas = campos.get("horas indisponíveis")

        if horas:
            try:
                horas = float(horas.replace(",", "."))
            except:
                horas = None

        registros.append({
            "Arquivo": file.name,
            "Natureza": natureza,
            "Equipamento": equipamento,
            "Horas Indisponíveis": horas
        })

    if registros:

        df = pd.DataFrame(registros)

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
            text="Ocorrências"
        )

        st.plotly_chart(fig1, use_container_width=True)

        fig2 = px.bar(
            horas_equip,
            x="Equipamento",
            y="Horas Indisponíveis",
            text="Horas Indisponíveis"
        )

        st.plotly_chart(fig2, use_container_width=True)

    else:

        st.warning("Nenhum registro encontrado nos arquivos.")

else:

    st.info("Aguardando upload de arquivos.")
