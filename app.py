import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="Análise HidroMeter Connect", layout="wide")
st.title("Análise de Ocorrências - HidroMeter Connect")

uploaded_files = st.file_uploader(
    "Arraste os arquivos .docm",
    type=["docm"],
    accept_multiple_files=True
)

def extract_full_text(file):

    try:
        with zipfile.ZipFile(file) as docm:
            xml_content = docm.read("word/document.xml")

        root = ET.fromstring(xml_content)

        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        texts = []

        for node in root.findall('.//w:t', ns):
            if node.text:
                texts.append(node.text)

        full_text = " ".join(texts)

        return full_text

    except Exception as e:
        st.error(f"Erro ao ler {file.name}: {e}")
        return ""


def parse_document(lines):

    registros = []

    natureza = None
    equipamento = None
    horas = None

    for line in lines:

        # produto
        if "HidroMeter Connect" in line:
            produto = "HidroMeter Connect"

        # natureza
        if "Natureza" in line:
            match = re.search(r'Natureza\s*[:\-]?\s*(.*)', line)
            if match:
                natureza = match.group(1).strip()

        # equipamento
        if "Equipamento" in line:
            match = re.search(r'Equipamento\s*[:\-]?\s*(.*)', line)
            if match:
                equipamento = match.group(1).strip()

        # horas indisponíveis
        if "Horas" in line and "Indispon" in line:
            match = re.search(r'([\d]+[\.,]?[\d]*)', line)
            if match:
                horas = float(match.group(1).replace(",", "."))

        # quando os três campos existem cria registro
        if natureza and equipamento and horas is not None:

            registros.append({
                "Natureza": natureza,
                "Equipamento": equipamento,
                "Horas Indisponíveis": horas
            })

            natureza = None
            equipamento = None
            horas = None

    return registros


if uploaded_files:

    registros_totais = []

    for file in uploaded_files:

        linhas = extract_text_from_docm(file)

        if not any("HidroMeter Connect" in l for l in linhas):
            st.info(f"{file.name} não contém HidroMeter Connect")
            continue

        registros = parse_document(linhas)

        registros_totais.extend(registros)

    if registros_totais:

        df = pd.DataFrame(registros_totais)

        st.subheader("Dados extraídos")
        st.dataframe(df)

        # ocorrências por natureza
        ocorrencias = df.groupby("Natureza").size().reset_index(name="Ocorrências")

        # total horas
        total_horas = df["Horas Indisponíveis"].sum()

        # horas por equipamento
        horas_equip = df.groupby("Equipamento")["Horas Indisponíveis"].sum().reset_index()

        st.subheader("Ocorrências por Natureza")
        st.dataframe(ocorrencias)

        st.metric("Total de Horas Indisponíveis", round(total_horas,2))

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
            title="Horas Indisponíveis por Equipamento"
        )

        st.plotly_chart(fig2, use_container_width=True)

    else:
        st.warning("Nenhum registro encontrado.")

else:
    st.info("Aguardando arquivos.")

