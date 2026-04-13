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
        return base64.b64encode(f.read()).decode()

#def set_background(png_file):
    try:
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
    except:
        pass

set_background("background.png")

st.title("Analisador Automático de Relatórios - CVT")

# ---------------------------------------------------------
# Upload de arquivos
# ---------------------------------------------------------
uploaded_files = st.file_uploader(
    "Envie os relatórios Word ou Excel",
    type=["docx", "docm", "dotm", "xlsx"],
    accept_multiple_files=True
)

NAMESPACE = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# ---------------------------------------------------------
# Funções para Word (XML robusto mantendo lógica original)
# ---------------------------------------------------------
def extract_text_and_tables(file):
    text_content = []
    tables = []

    try:
        with zipfile.ZipFile(file) as doc:

            if "word/document.xml" not in doc.namelist():
                return "", []

            xml_content = doc.read("word/document.xml")

        root = ET.fromstring(xml_content)
        body = root.find('w:body', NAMESPACE)

        if body is None:
            return "", []

        # TEXTO
        for t in body.findall('.//w:t', NAMESPACE):
            if t.text:
                text_content.append(t.text)

        # TABELAS (mesma lógica original)
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

    except Exception:
        return "", []

def extract_product(tables):
    for table in tables:
        for row in table:
            if len(row) >= 2:
                chave = str(row[0]).strip().lower()
                if chave.startswith("produto"):
                    produto = str(row[1]).strip()
                    return re.sub(r"\s+", " ", produto)
    return "Produto não identificado"

def extrair_mes_do_arquivo(file):
    filename = file.name
    padrao = r'(\d{1,2})[._-](\d{1,2})(?:[._-](\d{2,4}))?'
    match = re.search(padrao, filename)

    if match:
        dia, mes, ano = match.groups()
        mes = mes.zfill(2)

        if ano is None:
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
    if valor is None:
        return 0
    match = re.search(r'(\d+[.,]?\d*)', str(valor).lower())
    if match:
        return int(float(match.group(1).replace(",", ".")))
    return 0

# ---------------------------------------------------------
# Função para Excel
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
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
        else:
            df[col] = 0

    return df

# ---------------------------------------------------------
# Processamento principal
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
                    df_occ["MES_SORT"] = pd.to_datetime(
                        "01/" + df_occ["MES"],
                        format="%d/%m/%Y",
                        errors="coerce"
                    )
                    ocorrencias.append(df_occ)

                downtime_table = find_downtime_table(tables)

                if downtime_table:
                    df_down = pd.DataFrame(downtime_table[1:], columns=downtime_table[0])
                    df_down["PRODUTO"] = produto
                    df_down["MES"] = mes_relatorio
                    df_down["MES_SORT"] = pd.to_datetime(
                        "01/" + df_down["MES"],
                        format="%d/%m/%Y",
                        errors="coerce"
                    )

                    col_tempo = next(
                        (c for c in df_down.columns if "TEMPO" in c.upper()),
                        None
                    )

                    if col_tempo:
                        df_down["HORAS"] = df_down[col_tempo].apply(converter_horas)
                    else:
                        df_down["HORAS"] = 0

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
    # VIAGENS POR MÊS
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

            total_viagens = len(df_mes)
            st.metric("Total de Viagens", total_viagens)

            col_projeto = "Qual projeto foi visitado?"

            if col_projeto in df_mes.columns:

                df_mes[col_projeto] = df_mes[col_projeto].astype(str).str.strip()

                df_filtrado = df_mes[
                    ~df_mes[col_projeto].str.lower().isin(
                        ["nan", "não identificado", "escolha um item"]
                    )
                ]

                viagens_projeto = (
                    df_filtrado.groupby(col_projeto)
                    .size()
                    .reset_index(name="TOTAL VIAGENS")
                    .sort_values("TOTAL VIAGENS", ascending=False)
                )

                st.subheader("Viagens por Projeto")

                st.dataframe(
                    viagens_projeto.style.format({"TOTAL VIAGENS": "{:d}"}),
                    use_container_width=True
                )

            st.metric("Objetivos Traçados (total)", int(df_mes.iloc[:, -4].sum()))
            st.metric("Objetivos Cumpridos (total)", int(df_mes.iloc[:, -3].sum()))
            st.metric("Objetivos Extras (total)", int(df_mes.iloc[:, -2].sum()))
            st.metric("Objetivos Extras Cumpridos (total)", int(df_mes.iloc[:, -1].sum()))

else:
    st.info("Aguardando envio dos relatórios.")
