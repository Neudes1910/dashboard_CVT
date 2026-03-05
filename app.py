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

        # normaliza espaços
        full_text = re.sub(r'\s+', ' ', full_text)

        return full_text

    except Exception as e:
        st.error(f"Erro ao ler {file.name}: {e}")
        return ""


def extract_fields(text):

    natureza = None
    equipamento = None
    horas = None

    n = re.search(r'Natureza\s*:\s*([^\n\r]+)', text, re.IGNORECASE)
    if n:
        natureza = n.group(1).strip()

    e = re.search(r'Equipamento\s*:\s*([^\n\r]+)', text, re.IGNORECASE)
    if e:
        equipamento = e.group(1).strip()

    h = re.search(r'Horas\s*Indispon[ií]veis\s*:\s*([\d\.,]+)', text, re.IGNORECASE)
    if h:
        horas = float(h.group(1).replace(",", "."))

    return natureza, equipamento, horas


if uploaded_files:

    registros = []

    for file in uploaded_files:

        texto = extract_full_text(file)

        if not re.search(r'hidrometer\s+connect', texto, re.IGNORECASE):
            continue

        natureza, equipamento, horas = extract_fields(texto)

        if natureza or equipamento or horas:

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
        st.warning("Nenhum registro com Natureza/Equipamento/Horas foi encontrado.")

else:
    st.info("Aguardando upload de arquivos.")
