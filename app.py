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
# Background
# ---------------------------------------------------------
def get_base64(file):
    with open(file, "rb") as f:
        return base64.b64encode(f.read()).decode()

def set_background(png_file):
    try:
        bin_str = get_base64(png_file)
        st.markdown(f"""
        <style>
        .stApp {{
            background: url("data:image/png;base64,{bin_str}");
            background-size: 250px;
            background-position: calc(100% - 40px) 60px;
            background-repeat: no-repeat;
        }}
        </style>
        """, unsafe_allow_html=True)
    except:
        st.warning("Background não carregado")

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
# Word
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
            if len(row) >= 2 and str(row[0]).lower().startswith("produto"):
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
        df["Data de ida"] = pd.to_datetime(
            df["Data de ida (poderá ser uma data futura):"],
            errors="coerce"
        )
        df["MES"] = df["Data de ida"].dt.strftime("%m/%Y").fillna("Não identificado")
    else:
        df["MES"] = "Não identificado"

    objetivos = [
        "Quantos objetivos foram traçados antes da viagem? (apenas números)",
        "Dos objetivos traçados, quantos foram cumpridos? (apenas números)",
        "Houveram objetivos extras? (apenas números)",
        "Dos objetivos extras, quantos foram realizados? (apenas números)"
    ]

    for col in objetivos:
        df[col] = pd.to_numeric(df.get(col, 0), errors="coerce").fillna(0)

    return df

# ---------------------------------------------------------
# Processamento
# ---------------------------------------------------------
if uploaded_files:

    st.write("Checkpoint 2 - Arquivos carregados")

    ocorrencias = []
    horas_registros = []
    viagens_dados = []

    for file in uploaded_files:
        try:
            st.write(f"Processando: {file.name}")

            if file.name.endswith(("docx", "docm", "dotm")):

                text, tables = extract_text_and_tables(file)
                produto = extract_product(tables)
                mes_relatorio = extrair_mes_do_arquivo(file)

                occ_table = find_occurrence_table(tables)
                if occ_table:
                    df_occ = pd.DataFrame(occ_table[1:], columns=occ_table[0])
                    df_occ["PRODUTO"] = produto
                    df_occ["MES"] = mes_relatorio
                    ocorrencias.append(df_occ)

                downtime_table = find_downtime_table(tables)
                if downtime_table:
                    df_down = pd.DataFrame(downtime_table[1:], columns=downtime_table[0])
                    df_down["PRODUTO"] = produto
                    df_down["MES"] = mes_relatorio

                    col_tempo = next(c for c in df_down.columns if "TEMPO" in c.upper())
                    df_down["HORAS"] = df_down[col_tempo].apply(converter_horas)

                    horas_registros.append(df_down)

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
    # EXIBIÇÃO WORD
    # ---------------------------------------------------------
    if ocorrencias:
        st.header("Ocorrências")
        df_occ_total = pd.concat(ocorrencias, ignore_index=True)
        st.dataframe(df_occ_total, use_container_width=True)

    if horas_registros:
        st.header("Horas de Indisponibilidade")
        df_horas_total = pd.concat(horas_registros, ignore_index=True)
        st.dataframe(df_horas_total, use_container_width=True)

    # ---------------------------------------------------------
    # EXCEL
    # ---------------------------------------------------------
    if viagens_dados:

        df_viagens_total = pd.concat(viagens_dados, ignore_index=True)

        meses_ordenados = (
            df_viagens_total[["MES", "MES_SORT"]]
            .drop_duplicates()
            .sort_values("MES_SORT", ascending=False)
        )

        for _, row in meses_ordenados.iterrows():

            mes = row["MES"]
            df_mes = df_viagens_total[df_viagens_total["MES"] == mes]

            st.header(f"Viagens — {mes}")
            st.metric("Total de Viagens", len(df_mes))

            def safe_sum(df, col):
                return int(pd.to_numeric(df.get(col, 0), errors="coerce").fillna(0).sum())

            st.metric("Objetivos Traçados", safe_sum(df_mes, "Quantos objetivos foram traçados antes da viagem? (apenas números)"))
            st.metric("Objetivos Cumpridos", safe_sum(df_mes, "Dos objetivos traçados, quantos foram cumpridos? (apenas números)"))
            st.metric("Objetivos Extras", safe_sum(df_mes, "Houveram objetivos extras? (apenas números)"))
            st.metric("Extras Cumpridos", safe_sum(df_mes, "Dos objetivos extras, quantos foram realizados? (apenas números)"))

else:
    st.info("Aguardando envio dos relatórios.")
