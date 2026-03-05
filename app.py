import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import re

st.set_page_config(page_title="Detector HidroMeter Connect", layout="wide")
st.title("Verificação de Produto nos Documentos")

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

        # normaliza múltiplos espaços
        full_text = re.sub(r'\s+', ' ', full_text)

        return full_text

    except Exception as e:
        st.error(f"Erro ao ler {file.name}: {e}")
        return ""


if uploaded_files:

    for file in uploaded_files:

        texto = extract_full_text(file)

        if re.search(r'hidrometer\s+connect', texto, re.IGNORECASE):

            st.success(f"{file.name} contém HidroMeter Connect")

        else:

            st.warning(f"{file.name} NÃO contém HidroMeter Connect")

else:

    st.info("Aguardando arquivos.")
