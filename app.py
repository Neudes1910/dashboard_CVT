import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import base64
import requests

# ---------------------------------------------------------
# Configurações da página
# ---------------------------------------------------------
st.set_page_config(page_title="Analisador Automático de Relatórios - CVT", layout="wide")
st.title("Analisador Automático de Relatórios - CVT")

# ---------------------------------------------------------
# Função de marca d'água via URL
# ---------------------------------------------------------
def set_watermark_url(image_url, opacity=0.05, size=200):
    """
    Adiciona marca d'água no fundo do app Streamlit usando uma imagem da internet.
    """
    response = requests.get(image_url)
    b64 = base64.b64encode(response.content).decode()

    st.markdown(
        f"""
        <style>
        .stApp::before {{
            content: "";
            position: fixed;
            top: 0; left: 0;
            width: 100%; height: 100%;
            background-image: url("data:image/png;base64,{b64}");
            background-repeat: repeat;
            background-size: {size}px {size}px;
            opacity: {opacity};
            pointer-events: none;
            z-index: -1;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

# ---------------------------------------------------------
# Aplicar marca d'água via URL
# ---------------------------------------------------------
watermark_url = "data:image/png;base64,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"
set_watermark_url(watermark_url, opacity=0.1, size=200)

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

    # Coluna data de ida
    if "Data de ida (poderá ser uma data futura):" in df.columns:
        df["Data de ida"] = pd.to_datetime(df["Data de ida (poderá ser uma data futura):"], errors="coerce")
        df["MES"] = df["Data de ida"].dt.strftime("%m/%Y").fillna("Não identificado")
    else:
        df["MES"] = "Não identificado"

    # Objetivos
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
                    ocorrencias.append(df_occ)

                downtime_table = find_downtime_table(tables)
                if downtime_table:
                    df_down = pd.DataFrame(downtime_table[1:], columns=downtime_table[0])
                    df_down["PRODUTO"] = produto
                    df_down["MES"] = mes_relatorio
                    df_horas = df_down.copy()
                    col_tempo = next(c for c in df_horas.columns if "TEMPO" in c.upper())
                    df_horas["HORAS"] = df_horas[col_tempo].apply(converter_horas)
                    horas_registros.append(df_horas)

            elif file.name.endswith("xlsx"):
                df_viagens = process_excel(file)
                viagens_dados.append(df_viagens)

        except Exception as e:
            st.warning(f"Erro ao processar {file.name}: {e}")

    # ---------------------- OCORRÊNCIAS ----------------------
    if ocorrencias:
        df_total = pd.concat(ocorrencias, ignore_index=True)
        natureza_col = [c for c in df_total.columns if "NATUREZA" in c.upper()][0]
        df_total[natureza_col] = df_total[natureza_col].astype(str).str.strip()
        excluir = ["escolha um item", "escolher um item."]
        df_total = df_total[~df_total[natureza_col].str.lower().isin(excluir)]

        df_total["MES_SORT"] = pd.to_datetime("01/" + df_total["MES"], format="%d/%m/%Y", errors="coerce")
        df_total = df_total.sort_values("MES_SORT", ascending=False)

        for mes in df_total["MES"].drop_duplicates():
            st.header(f"Mês: {mes}")
            df_mes = df_total[df_total["MES"] == mes]
            resumo = df_mes.groupby(["PRODUTO", natureza_col]).size().reset_index(name="TOTAL OCORRÊNCIAS")
            st.subheader("Ocorrências por Natureza")
            st.dataframe(resumo.style.set_properties(**{"font-size": "16px"}), use_container_width=True)

    # ---------------------- HORAS DE INDISPONIBILIDADE ----------------------
    if horas_registros:
        df_horas = pd.concat(horas_registros, ignore_index=True)
        col_nat = next(c for c in df_horas.columns if "NATUREZA" in c.upper())
        col_equip = next(c for c in df_horas.columns if "QUAL EQUIPAMENTO" in c.upper())
        df_horas[col_nat] = df_horas[col_nat].astype(str).str.strip()
        df_horas = df_horas[~df_horas[col_nat].str.lower().isin(["escolha um item", "escolher um item."])]
        df_horas = df_horas.sort_values("MES", ascending=False)

        for mes in df_horas["MES"].drop_duplicates():
            st.header(f"Horas Indisponíveis — {mes}")
            df_mes = df_horas[df_horas["MES"] == mes]
            total_horas = df_mes["HORAS"].sum()
            st.metric("Total de Horas Indisponíveis", total_horas)

            horas_nat = df_mes.groupby(["PRODUTO", col_nat])["HORAS"].sum().reset_index()
            st.subheader("Horas por Natureza")
            st.dataframe(horas_nat.style.set_properties(**{"font-size": "16px"}), use_container_width=True)

            horas_eq = df_mes.groupby(["PRODUTO", col_equip])["HORAS"].sum().reset_index()
            st.subheader("Horas por Equipamento")
            st.dataframe(horas_eq.style.set_properties(**{"font-size": "16px"}), use_container_width=True)

    # ---------------------- VIAGENS ----------------------
    if viagens_dados:
        df_viagens_total = pd.concat(viagens_dados, ignore_index=True)
        df_viagens_total = df_viagens_total.sort_values("MES", ascending=False)
        meses = df_viagens_total["MES"].drop_duplicates()

        for mes in meses:
            st.header(f"Viagens — {mes}")
            df_mes = df_viagens_total[df_viagens_total["MES"] == mes]

            total_viagens = len(df_mes)
            st.metric("Total de Viagens", total_viagens)

            col_obj_trac = "Quantos objetivos foram traçados antes da viagem? (apenas números)"
            col_obj_cump = "Dos objetivos traçados, quantos foram cumpridos? (apenas números)"
            col_obj_extra = "Houveram objetivos extras? (apenas números)"
            col_obj_extra_cump = "Dos objetivos extras, quantos foram realizados? (apenas números)"

            st.metric("Objetivos Traçados (total)", df_mes[col_obj_trac].sum())
            st.metric("Objetivos Cumpridos (total)", df_mes[col_obj_cump].sum())
            st.metric("Objetivos Extras (total)", df_mes[col_obj_extra].sum())
            st.metric("Objetivos Extras Cumpridos (total)", df_mes[col_obj_extra_cump].sum())

else:
    st.info("Aguardando envio dos relatórios.")
