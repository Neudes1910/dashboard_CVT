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

# ---------------------------------------------------------
# DEFINA A IMAGEM DE FUNDO AQUI
# ---------------------------------------------------------
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
# Funções para Word
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
    # OCORRÊNCIAS POR MÊS
    # ---------------------------------------------------------
    if ocorrencias:

        df_total = pd.concat(ocorrencias, ignore_index=True)

        natureza_col = [c for c in df_total.columns if "NATUREZA" in c.upper()][0]

        df_total[natureza_col] = df_total[natureza_col].astype(str).str.strip()

        excluir = ["escolha um item", "escolher um item."]

        df_total = df_total[~df_total[natureza_col].str.lower().isin(excluir)]

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

            st.subheader("Ocorrências por Natureza")

            st.dataframe(resumo.style.format({"TOTAL OCORRÊNCIAS": "{:d}"}).set_properties(**{"font-size": "16px"}), use_container_width=True)
# ---------------------------------------------------------
    # HORAS DE INDISPONIBILIDADE POR MÊS
    # ---------------------------------------------------------
    if horas_registros:

        df_horas = pd.concat(horas_registros, ignore_index=True)

        col_nat = next(c for c in df_horas.columns if "NATUREZA" in c.upper())
        col_equip = next(c for c in df_horas.columns if "QUAL EQUIPAMENTO" in c.upper())

        df_horas[col_nat] = df_horas[col_nat].astype(str).str.strip()

        df_horas = df_horas[
            ~df_horas[col_nat].str.lower().isin(["escolha um item", "escolher um item."])
        ]

        meses_ordenados = df_horas[["MES", "MES_SORT"]].drop_duplicates().sort_values("MES_SORT", ascending=False)

        for _, row in meses_ordenados.iterrows():

            mes = row["MES"]

            df_mes = df_horas[df_horas["MES"] == mes]

            st.header(f"Horas Indisponíveis — {mes}")

            total_horas = int(df_mes["HORAS"].sum())

            st.metric("Total de Horas Indisponíveis", total_horas)

            horas_nat = df_mes.groupby(["PRODUTO", col_nat])["HORAS"].sum().reset_index()
            horas_nat["HORAS"] = horas_nat["HORAS"].astype(int)

            st.subheader("Horas por Natureza")

            st.dataframe(
                horas_nat.style.format({"HORAS": "{:d}"}).set_properties(**{"font-size": "16px"}),
                use_container_width=True
            )

            horas_eq = df_mes.groupby(["PRODUTO", col_equip])["HORAS"].sum().reset_index()
            horas_eq["HORAS"] = horas_eq["HORAS"].astype(int)

            st.subheader("Horas por Equipamento")

            st.dataframe(
                horas_eq.style.format({"HORAS": "{:d}"}).set_properties(**{"font-size": "16px"}),
                use_container_width=True
            )


    # ---------------------------------------------------------
    # VIAGENS POR MÊS
    # ---------------------------------------------------------
    if viagens_dados:

        df_viagens_total = pd.concat(viagens_dados, ignore_index=True)

        meses_ordenados = df_viagens_total[["MES", "MES_SORT"]].drop_duplicates().sort_values("MES_SORT", ascending=False)

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
        ~df_mes[col_projeto].str.lower().isin([
            "nan", "não identificado", "escolha um item"
        ])
    ]

    viagens_projeto = (
        df_filtrado.groupby(col_projeto)
        .size()
        .reset_index(name="TOTAL VIAGENS")
        .sort_values("TOTAL VIAGENS", ascending=False)
    )

    st.subheader("Viagens por Projeto")

    st.dataframe(
        viagens_projeto.style.format({"TOTAL VIAGENS": "{:d}"})
        .set_properties(**{"font-size": "16px"}),
        use_container_width=True
    )
            col_obj_trac = "Quantos objetivos foram traçados antes da viagem? (apenas números)"
            col_obj_cump = "Dos objetivos traçados, quantos foram cumpridos? (apenas números)"
            col_obj_extra = "Houveram objetivos extras? (apenas números)"
            col_obj_extra_cump = "Dos objetivos extras, quantos foram realizados? (apenas números)"

            st.metric("Objetivos Traçados (total)", int(df_mes[col_obj_trac].sum()))
            st.metric("Objetivos Cumpridos (total)", int(df_mes[col_obj_cump].sum()))
            st.metric("Objetivos Extras (total)", int(df_mes[col_obj_extra].sum()))
            st.metric("Objetivos Extras Cumpridos (total)", int(df_mes[col_obj_extra_cump].sum()))
else:

    st.info("Aguardando envio dos relatórios.")
