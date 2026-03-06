import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re

st.set_page_config(page_title="Analisador Automático de Relatórios - CVT", layout="wide")
st.title("Analisador Automático de Relatórios - CVT")

uploaded_files = st.file_uploader(
    "Envie os relatórios Word",
    type=["docx", "docm", "dotm"],
    accept_multiple_files=True
)

NAMESPACE = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# ---------------------------------------------------------
# EXTRAÇÃO DE TEXTO E TABELAS
# ---------------------------------------------------------
def extract_text_and_tables(file):
    text_content = []
    tables = []

    with zipfile.ZipFile(file) as doc:
        xml_content = doc.read("word/document.xml")

    root = ET.fromstring(xml_content)

    # Extrai todo o texto do documento
    for t in root.findall('.//w:t', NAMESPACE):
        if t.text:
            text_content.append(t.text)

    # Extrai todas as tabelas
    for tbl in root.findall('.//w:tbl', NAMESPACE):
        table_data = []
        for row in tbl.findall('.//w:tr', NAMESPACE):
            cells = []
            for cell in row.findall('.//w:tc', NAMESPACE):
                texts = [t.text for t in cell.findall('.//w:t', NAMESPACE) if t.text]
                cell_text = " ".join(texts).strip()
                cells.append(cell_text)
            if cells:
                table_data.append(cells)
        if table_data:
            tables.append(table_data)

    return " ".join(text_content), tables

# ---------------------------------------------------------
# EXTRAIR PRODUTO
# ---------------------------------------------------------
def extract_product(tables):
    for table in tables:
        for row in table:
            if len(row) >= 2:
                chave = str(row[0]).strip().lower()
                if chave.startswith("produto"):
                    produto = str(row[1]).strip()
                    produto = re.sub(r"\s+", " ", produto)
                    return produto
    return "Produto não identificado"

# ---------------------------------------------------------
# EXTRAIR MÊS A PARTIR DE "Data início:"
# ---------------------------------------------------------
def extrair_mes_do_texto(texto):
    padrao = r"in[ií]cio\s*:\s*(\d{2})[./-](\d{2})[./-](\d{4})"
    match = re.search(padrao, texto, re.IGNORECASE)
    if match:
        dia, mes, ano = match.groups()
        return f"{mes}/{ano}"
    return "Não identificado"

# ---------------------------------------------------------
# IDENTIFICAR TABELAS
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# CONVERSÃO DE HORAS
# ---------------------------------------------------------
def converter_horas(valor):
    if valor is None:
        return 0
    texto = str(valor).lower()
    match = re.search(r'(\d+[.,]?\d*)', texto)
    if match:
        return float(match.group(1).replace(",", "."))
    return 0

# ---------------------------------------------------------
# PROCESSAMENTO
# ---------------------------------------------------------
if uploaded_files:
    ocorrencias = []
    horas_registros = []

    for file in uploaded_files:
        try:
            text, tables = extract_text_and_tables(file)
            produto = extract_product(tables)
            mes_relatorio = extrair_mes_do_texto(text)

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
                horas_registros.append(df_down)

        except Exception as e:
            st.warning(f"Erro ao processar {file.name}: {e}")

    # ---------------------------------------------------------
    # OCORRÊNCIAS POR MÊS
    # ---------------------------------------------------------
    if ocorrencias:
        df_total = pd.concat(ocorrencias, ignore_index=True)
        natureza_col = [c for c in df_total.columns if "NATUREZA" in c.upper()][0]
        df_total[natureza_col] = df_total[natureza_col].astype(str).str.strip()
        excluir = ["escolha um item", "escolher um item."]
        df_total = df_total[~df_total[natureza_col].str.lower().isin(excluir)]
        meses = sorted(df_total["MES"].unique())

        for mes in meses:
            st.header(f"Mês: {mes}")
            df_mes = df_total[df_total["MES"] == mes]
            resumo = (
                df_mes
                .groupby(["PRODUTO", natureza_col])
                .size()
                .reset_index(name="TOTAL OCORRÊNCIAS")
            )
            st.subheader("Ocorrências por Natureza")
            st.dataframe(
                resumo.style.set_properties(**{"font-size": "16px"}),
                use_container_width=True
            )

    # ---------------------------------------------------------
    # HORAS DE INDISPONIBILIDADE POR MÊS
    # ---------------------------------------------------------
    if horas_registros:
        df_horas = pd.concat(horas_registros, ignore_index=True)
        df_horas.columns = df_horas.columns.str.replace("\n", " ").str.strip()

        col_tempo = next(c for c in df_horas.columns if "TEMPO" in c.upper())
        col_equip = next(c for c in df_horas.columns if "QUAL EQUIPAMENTO" in c.upper())
        col_nat = next(c for c in df_horas.columns if "NATUREZA" in c.upper())

        df_horas["HORAS"] = df_horas[col_tempo].apply(converter_horas)
        df_horas[col_nat] = df_horas[col_nat].astype(str).str.strip()
        df_horas = df_horas[~df_horas[col_nat].str.lower().isin(["escolha um item", "escolher um item."])]
        meses = sorted(df_horas["MES"].unique())

        for mes in meses:
            st.header(f"Horas Indisponíveis — {mes}")
            df_mes = df_horas[df_horas["MES"] == mes]
            total_horas = df_mes["HORAS"].sum()
            st.metric("Total de Horas Indisponíveis", round(total_horas, 2))

            horas_nat = (
                df_mes
                .groupby(["PRODUTO", col_nat])["HORAS"]
                .sum()
                .reset_index()
            )
            st.subheader("Horas por Natureza")
            st.dataframe(horas_nat.style.set_properties(**{"font-size": "16px"}), use_container_width=True)

            horas_eq = (
                df_mes
                .groupby(["PRODUTO", col_equip])["HORAS"]
                .sum()
                .reset_index()
            )
            st.subheader("Horas por Equipamento")
            st.dataframe(horas_eq.style.set_properties(**{"font-size": "16px"}), use_container_width=True)

else:
    st.info("Aguardando envio dos relatórios.")

