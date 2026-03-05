import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="Análise HidroMeter Connect", layout="wide")
st.title("Análise de Ocorrências - HidroMeter Connect")

uploaded_files = st.file_uploader(
    "Arraste arquivos .docm aqui",
    type=["docm"],
    accept_multiple_files=True
)

def extract_full_text(file):
    """
    Extrai TODO o texto do documento Word.
    Isso evita problemas com palavras quebradas no XML.
    """

    try:
        with zipfile.ZipFile(file) as docm:
            xml_content = docm.read("word/document.xml")

        root = ET.fromstring(xml_content)

        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        texts = []

        for node in root.findall('.//w:t', ns):
            if node.text:
                texts.append(node.text.strip())

        full_text = " ".join(texts)

        return full_text

    except Exception as e:
        st.error(f"Erro ao processar {file.name}: {e}")
        return ""


def extract_records(full_text):
    """
    Separa registros a partir de palavras chave
    """

    registros = []

    natureza = None
    equipamento = None
    horas = None

    linhas = full_text.split()

    buffer = " ".join(linhas)

    # procura natureza
    naturezas = re.findall(r'Natureza\s*[:\-]?\s*([A-Za-z ]+)', buffer)

    # procura equipamentos
    equipamentos = re.findall(r'Equipamento\s*[:\-]?\s*([A-Za-z0-9_\- ]+)', buffer)

    # procura horas
    horas_list = re.findall(r'Horas\s*Indispon[ií]veis\s*[:\-]?\s*([\d]+[\.,]?[\d]*)', buffer)

    for i in range(min(len(naturezas), len(equipamentos), len(horas_list))):

        registros.append({
            "Natureza": naturezas[i].strip(),
            "Equipamento": equipamentos[i].strip(),
            "Horas Indisponíveis": float(horas_list[i].replace(",", "."))
        })

    return registros


if uploaded_files:

    todos_registros = []

    for file in uploaded_files:

        texto = extract_full_text(file)

        # procura produto em QUALQUER parte do documento
        if "hidrometer connect" not in texto.lower():
            st.info(f"{file.name} não contém HidroMeter Connect")
            continue

        registros = extract_records(texto)

        if registros:
            todos_registros.extend(registros)

    if todos_registros:

        df = pd.DataFrame(todos_registros)

        st.subheader("Dados extraídos")
        st.dataframe(df)

        # ocorrências por natureza
        ocorrencias = df.groupby("Natureza").size().reset_index(name="Ocorrências")

        # total horas indisponíveis
        total_horas = df["Horas Indisponíveis"].sum()

        # horas por equipamento
        horas_equip = df.groupby("Equipamento")["Horas Indisponíveis"].sum().reset_index()

        st.subheader("Ocorrências por Natureza")
        st.dataframe(ocorrencias)

        st.metric("Total de Horas Indisponíveis", round(total_horas, 2))

        st.subheader("Horas Indisponíveis por Equipamento")
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
        st.warning("Nenhum registro encontrado.")

else:
    st.info("Aguardando upload de arquivos.")
