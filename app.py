import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import base64
import traceback

# ---------------------------------------------------------
# Configurações da página
# ---------------------------------------------------------
st.set_page_config(page_title="Analisador Automático de Relatórios - CVT", layout="wide")

# ---------------------------------------------------------
# FUNÇÃO PARA BACKGROUND
# ---------------------------------------------------------
def get_base64(file):
    with open(file, "rb") as f:
        data = f.read()
    return base64.b64encode(data).decode()

def set_background(png_file):
    try:
        bin_str = get_base64(png_file)
        page_bg = f"""
        <style>
        .stApp {{
            background: url("data:image/png;base64,{bin_str}");
            background-size: 250px;
            background-position: calc(100% - 40px) 60px;
            background-repeat: no-repeat;
        }}
        </style>
        """
        st.markdown(page_bg, unsafe_allow_html=True)
    except Exception:
        st.warning("Background não carregado.")

set_background("background.png")

st.title("Analisador Automático de Relatórios - CVT")

st.write("Checkpoint 1 - App iniciado")

# ---------------------------------------------------------
# Upload
# ---------------------------------------------------------
uploaded_files = st.file_uploader(
    "Envie os relatórios Word ou Excel",
    type=["docx", "docm", "dotm", "xlsx"],
    accept_multiple_files=True
)

NAMESPACE = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# ---------------------------------------------------------
# Funções Word
# ---------------------------------------------------------
def extract_text_and_tables(file):
    text_content = []
    tables = []

    with zipfile.ZipFile(file) as doc:
        xml_content = doc.read("word/document.xml")

    root = ET.fromstring(xml_content)
    body = root.find('w:body', NAMESPACE)

    for t in body.findall('.//w:t', NAMESPACE):
        if t.text:
            text_content.append(t.text)

    for tbl in body.findall('.//w:tbl', NAMESPACE):
        table_data = []
        for row in tbl.findall('.//w:tr', NAMESPACE):
            cells = []
            for cell in row.findall('.//w:tc', NAMESPACE):
                texts = [t.text for t in cell.findall('.//w:t', NAMESPACE) if t.text]
                cells.append(" ".join(texts).strip())
            if cells:
                table_data.append(cells)
        if table_data:
            tables.append(table_data)

    return " ".join(text_content), tables


def extract_product(tables):
    for table in tables:
        for row in table:
            if len(row) >= 2:
                if str(row[0]).lower().startswith("produto"):
                    return str(row[1]).strip()
    return "Produto não identificado"


def extrair_mes_do_arquivo(file):
    match = re.search(r'(\d{1,2})[._-](\d{1,2})(?:[._-](\d{2,4}))?', file.name)
    if match:
        _, mes, ano = match.groups()
        mes = mes.zfill(2)
        ano = "2026" if not ano else ("20" + ano if len(ano) == 2 else ano)
        return f"{mes}/{ano}"
    return "Não identificado"


def find_occurrence_table(tables):
    for table in tables:
        header = [str(x).upper() for x in table[0]]
        if "NATUREZA" in header and "OCORRÊNCIA" in header:
            return table
    return None


def find_downtime_table(tables):
    for table in tables:
        header = [str(x).upper() for x in table[0]]
        if "POR QUANTO TEMPO?" in header and "QUAL EQUIPAMENTO?" in header:
            return table
    return None


def converter_horas(valor):
    match = re.search(r'(\d+[.,]?\d*)', str(valor))
    return int(float(match.group(1).replace(",", "."))) if match else 0

# ---------------------------------------------------------
# Excel
# ---------------------------------------------------------
def process_excel(file):
    df = pd.read_excel(file, engine="openpyxl")
    df.columns = df.columns.str.strip()

    if "Data de ida (poderá ser uma data futura):" in df.columns:
        df["Data de ida"] = pd.to_datetime(df.iloc[:, 0], errors="coerce")
        df["MES"] = df["Data de ida"].dt.strftime("%m/%Y").fillna("Não identificado")
    else:
        df["MES"] = "Não identificado"

    for col in [
        "Quantos objetivos foram traçados antes da viagem? (apenas números)",
        "Dos objetivos traçados, quantos foram cumpridos? (apenas números)",
        "Houveram objetivos extras? (apenas números)",
        "Dos objetivos extras, quantos foram realizados? (apenas números)"
    ]:
        df[col] = pd.to_numeric(df.get(col, 0), errors="coerce").fillna(0)

    return df

# ---------------------------------------------------------
# Processamento
# ---------------------------------------------------------
if uploaded_files:

    st.write("Checkpoint 2 - Arquivos carregados")

    viagens_dados = []

    for file in uploaded_files:
        try:
            st.write(f"Processando: {file.name}")

            if file.name.endswith(("docx", "docm", "dotm")):
                text, tables = extract_text_and_tables(file)
                st.write(f"Tabelas encontradas: {len(tables)}")

            elif file.name.endswith("xlsx"):
                df_viagens = process_excel(file)
                df_viagens["MES_SORT"] = pd.to_datetime(
                    "01/" + df_viagens["MES"],
                    format="%d/%m/%Y",
                    errors="coerce"
                )
                viagens_dados.append(df_viagens)

        except Exception:
            st.error(f"Erro ao processar {file.name}")
            st.text(traceback.format_exc())

    # ---------------------------------------------------------
    # VIAGENS
    # ---------------------------------------------------------
    if viagens_dados:

        st.write("Checkpoint 3 - Iniciando concat")

        try:
            df_viagens_total = pd.concat(viagens_dados, ignore_index=True)
        except Exception:
            st.error("Erro no concat")
            st.text(traceback.format_exc())
            st.stop()

        try:
            meses_ordenados = (
                df_viagens_total[["MES", "MES_SORT"]]
                .drop_duplicates()
                .sort_values("MES_SORT", ascending=False)
            )
        except Exception:
            st.error("Erro ao ordenar meses")
            st.text(traceback.format_exc())
            st.stop()

        for _, row in meses_ordenados.iterrows():

            mes = row["MES"]
            df_mes = df_viagens_total[df_viagens_total["MES"] == mes]

            st.header(f"Viagens — {mes}")
            st.metric("Total de Viagens", len(df_mes))

            col_projeto = "Qual projeto foi visitado?"

            if col_projeto in df_mes.columns:

                df_filtrado = df_mes[
                    ~df_mes[col_projeto].astype(str).str.lower().isin(
                        ["nan", "não identificado", "escolha um item"]
                    )
                ]

                viagens_projeto = (
                    df_filtrado.groupby(col_projeto)
                    .size()
                    .reset_index(name="TOTAL VIAGENS")
                )

                st.dataframe(viagens_projeto, use_container_width=True)

            # MÉTRICAS SEGURAS
            def safe_sum(df, col):
                return int(pd.to_numeric(df.get(col, 0), errors="coerce").fillna(0).sum())

            st.metric("Objetivos Traçados", safe_sum(df_mes, "Quantos objetivos foram traçados antes da viagem? (apenas números)"))
            st.metric("Objetivos Cumpridos", safe_sum(df_mes, "Dos objetivos traçados, quantos foram cumpridos? (apenas números)"))
            st.metric("Objetivos Extras", safe_sum(df_mes, "Houveram objetivos extras? (apenas números)"))
            st.metric("Extras Cumpridos", safe_sum(df_mes, "Dos objetivos extras, quantos foram realizados? (apenas números)"))

else:
    st.info("Aguardando envio dos relatórios.")
