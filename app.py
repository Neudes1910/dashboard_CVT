import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import base64

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
    bin_str = get_base64(png_file)
    page_bg = f"""
    <style>
    .stApp {{
        background: linear-gradient(
            rgba(0,0,0,0),
            rgba(0,0,0,0)
        ),
        url("data:image/png;base64,{bin_str}");
        background-size: 250px;
        background-position: calc(100% - 40px) 60px;
        background-attachment: fixed;
        background-repeat: no-repeat;
    }}
    </style>
    """
    st.markdown(page_bg, unsafe_allow_html=True)

set_background("background.png")

st.title("Analisador Automático de Relatórios - CVT")

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
                chave = str(row[0]).strip().lower()
                if chave.startswith("produto"):
                    return str(row[1]).strip()
    return "Produto não identificado"


def extrair_mes_do_arquivo(file):
    match = re.search(r'(\d{1,2})[._-](\d{1,2})(?:[._-](\d{2,4}))?', file.name)
    if match:
        _, mes, ano = match.groups()
        mes = mes.zfill(2)
        if not ano:
            ano = "2026"
        elif len(ano) == 2:
            ano = "20" + ano
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
        df["Data de ida"] = pd.to_datetime(df["Data de ida (poderá ser uma data futura):"], errors="coerce")
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
        df[col] = pd.to_numeric(df.get(col, 0), errors="coerce").fillna(0).astype(int)

    return df

# ---------------------------------------------------------
# Processamento
# ---------------------------------------------------------
if uploaded_files:

    ocorrencias = []
    horas_registros = []
    viagens_dados = []

    for file in uploaded_files:
        try:

            if file.name.endswith(("docx", "docm", "dotm")):

                text, tables = extract_text_and_tables(file)
                produto = extract_product(tables)
                mes_relatorio = extrair_mes_do_arquivo(file)

                occ_table = find_occurrence_table(tables)
                if occ_table:
                    df_occ = pd.DataFrame(occ_table[1:], columns=occ_table[0])
                    df_occ["PRODUTO"] = produto
                    df_occ["MES"] = mes_relatorio
                    df_occ["MES_SORT"] = pd.to_datetime("01/" + df_occ["MES"], format="%d/%m/%Y", errors="coerce")
                    ocorrencias.append(df_occ)

                downtime_table = find_downtime_table(tables)
                if downtime_table:
                    df_down = pd.DataFrame(downtime_table[1:], columns=downtime_table[0])
                    df_down["PRODUTO"] = produto
                    df_down["MES"] = mes_relatorio
                    df_down["MES_SORT"] = pd.to_datetime("01/" + df_down["MES"], format="%d/%m/%Y", errors="coerce")

                    col_tempo = next(c for c in df_down.columns if "TEMPO" in c.upper())
                    df_down["HORAS"] = df_down[col_tempo].apply(converter_horas)

                    horas_registros.append(df_down)

            elif file.name.endswith("xlsx"):

                df_viagens = process_excel(file)
                df_viagens["MES_SORT"] = pd.to_datetime("01/" + df_viagens["MES"], format="%d/%m/%Y", errors="coerce")
                viagens_dados.append(df_viagens)

        except Exception as e:
            st.warning(f"Erro ao processar {file.name}: {e}")

    # ---------------------------------------------------------
    # OCORRÊNCIAS
    # ---------------------------------------------------------
    if ocorrencias:

        df_total = pd.concat(ocorrencias, ignore_index=True)
        natureza_col = [c for c in df_total.columns if "NATUREZA" in c.upper()][0]

        df_total = df_total[
            ~df_total[natureza_col].astype(str).str.lower().isin(["escolha um item", "escolher um item."])
        ]

        meses_ordenados = df_total[["MES", "MES_SORT"]].drop_duplicates().sort_values("MES_SORT", ascending=False)

        for _, row in meses_ordenados.iterrows():

            mes = row["MES"]
            df_mes = df_total[df_total["MES"] == mes]

            st.header(f"Mês: {mes}")

            resumo = (
                df_mes.groupby(["PRODUTO", natureza_col])
                .size()
                .reset_index(name="TOTAL OCORRÊNCIAS")
            )

            st.dataframe(resumo, use_container_width=True)

    # ---------------------------------------------------------
# HORAS
# ---------------------------------------------------------
if horas_registros:

    df_horas = pd.concat(horas_registros, ignore_index=True)

    col_nat = next(c for c in df_horas.columns if "NATUREZA" in c.upper())
    col_equip = next(c for c in df_horas.columns if "EQUIPAMENTO" in c.upper())

    df_horas = df_horas[
        ~df_horas[col_nat].astype(str).str.lower().isin(["escolha um item", "escolher um item."])
    ]

    meses_ordenados = df_horas[["MES", "MES_SORT"]].drop_duplicates().sort_values("MES_SORT", ascending=False)

    for _, row in meses_ordenados.iterrows():

        mes = row["MES"]
        df_mes = df_horas[df_horas["MES"] == mes]

        st.header(f"Horas — {mes}")

        # TOTAL
        st.metric("Total de Horas", int(df_mes["HORAS"].sum()))

        # --------------------------------------------
        # LISTAGEM DETALHADA (o que você pediu)
        # --------------------------------------------
        df_detalhado = df_mes[[col_equip, "HORAS", "PRODUTO"]].copy()

        df_detalhado = df_detalhado[
            ~df_detalhado[col_equip].astype(str).str.lower().isin(
                ["nan", "não identificado", "escolha um item."]
            )
        ]

        df_detalhado["HORAS"] = df_detalhado["HORAS"].astype(int)

        st.subheader("Detalhamento de Indisponibilidade")

        st.dataframe(
            df_detalhado.rename(columns={
                col_equip: "EQUIPAMENTO",
                "HORAS": "TEMPO (h)"
            }),
            use_container_width=True
        )
    # ---------------------------------------------------------
    # VIAGENS
    # ---------------------------------------------------------
    if viagens_dados:

        df_viagens_total = pd.concat(viagens_dados, ignore_index=True)

        meses_ordenados = df_viagens_total[["MES", "MES_SORT"]].drop_duplicates().sort_values("MES_SORT", ascending=False)

        for _, row in meses_ordenados.iterrows():

            mes = row["MES"]
            df_mes = df_viagens_total[df_viagens_total["MES"] == mes]

            st.header(f"Viagens — {mes}")
            st.metric("Total de Viagens", len(df_mes))

else:
    st.info("Aguardando envio dos relatórios.")
