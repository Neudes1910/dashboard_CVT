import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re

st.set_page_config(page_title="Analisador Automático de Relatórios - CVT", layout="wide")
st.title("Analisador Automático de Relatórios - CVT")

uploaded_files = st.file_uploader(
    "Envie os relatórios Word ou Excel",
    type=None,
    accept_multiple_files=True
)

NAMESPACE = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# ---------------------------------------------------------
# EXTRAÇÃO DE TEXTO E TABELAS WORD
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
                    return re.sub(r"\s+", " ", produto)
    return "Produto não identificado"

# ---------------------------------------------------------
# EXTRAIR MÊS A PARTIR DO NOME DO ARQUIVO
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# IDENTIFICAR TABELAS WORD
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
    match = re.search(r'(\d+[.,]?\d*)', str(valor).lower())
    if match:
        return int(float(match.group(1).replace(",", ".")))
    return 0

# ---------------------------------------------------------
# CONVERSÃO DE STRING MM/YYYY PARA DATETIME
# ---------------------------------------------------------
def mes_para_datetime(mes_str):
    try:
        return pd.to_datetime(mes_str, format="%m/%Y")
    except:
        return pd.NaT

# ---------------------------------------------------------
# PROCESSAMENTO
# ---------------------------------------------------------
if uploaded_files:
    ocorrencias = []
    horas_registros = []
    viagens = []

    for file in uploaded_files:
        try:
            if file.name.endswith(('.docx', '.docm', '.dotm')):
                text, tables = extract_text_and_tables(file)
                produto = extract_product(tables)
                mes_relatorio = extrair_mes_do_arquivo(file)
                mes_dt = mes_para_datetime(mes_relatorio)

                occ_table = find_occurrence_table(tables)
                if occ_table:
                    df_occ = pd.DataFrame(occ_table[1:], columns=occ_table[0])
                    df_occ["PRODUTO"] = produto
                    df_occ["MES"] = mes_relatorio
                    df_occ["MES_DT"] = mes_dt
                    ocorrencias.append(df_occ)

                downtime_table = find_downtime_table(tables)
                if downtime_table:
                    df_down = pd.DataFrame(downtime_table[1:], columns=downtime_table[0])
                    df_down["PRODUTO"] = produto
                    df_down["MES"] = mes_relatorio
                    df_down["MES_DT"] = mes_dt
                    df_down["HORAS"] = df_down.iloc[:,0].apply(converter_horas)
                    horas_registros.append(df_down)

            elif file.name.endswith('.xlsx'):
                df_excel = pd.read_excel(file)
                if "Data de ida (poderá ser uma data futura):" in df_excel.columns:
                    df_excel["Data de ida"] = pd.to_datetime(
                        df_excel["Data de ida (poderá ser uma data futura):"], errors='coerce'
                    )
                    df_excel = df_excel.dropna(subset=["Data de ida"])
                    df_excel["MES"] = df_excel["Data de ida"].dt.strftime("%m/%Y")
                    df_excel["MES_DT"] = pd.to_datetime(df_excel["MES"], format="%m/%Y")

                    # Objetivos traçados e cumpridos
                    objetivos_colunas = [c for c in df_excel.columns if "Quantos objetivos foram traçados antes da viagem" in c]
                    df_excel["OBJETIVOS"] = pd.to_numeric(df_excel[objetivos_colunas[0]], errors='coerce').fillna(0).astype(int) if objetivos_colunas else 0
                    cumpridos_colunas = [c for c in df_excel.columns if "Dos objetivos traçados, quantos foram cumpridos" in c]
                    df_excel["OBJETIVOS_CUMPRIDOS"] = pd.to_numeric(df_excel[cumpridos_colunas[0]], errors='coerce').fillna(0).astype(int) if cumpridos_colunas else 0

                    # Objetivos extras traçados e realizados
                    extras_col = [c for c in df_excel.columns if "Houveram objetivos extras" in c]
                    df_excel["OBJETIVOS_EXTRAS"] = pd.to_numeric(df_excel[extras_col[0]], errors='coerce').fillna(0).astype(int) if extras_col else 0
                    realizados_col = [c for c in df_excel.columns if "Dos objetivos extras, quantos foram realizados" in c]
                    df_excel["OBJETIVOS_EXTRAS_CUMPRIDOS"] = pd.to_numeric(df_excel[realizados_col[0]], errors='coerce').fillna(0).astype(int) if realizados_col else 0

                    viagens.append(df_excel)

            else:
                st.warning(f"Arquivo não suportado: {file.name}")

        except Exception as e:
            st.warning(f"Erro ao processar {file.name}: {e}")

    # ---------------------------------------------------------
    # ORDENAR MESES MAIS RECENTES PRIMEIRO
    # ---------------------------------------------------------
    def obter_meses_ordenados(df):
        df_meses = df[["MES", "MES_DT"]].drop_duplicates()
        df_meses = df_meses.sort_values("MES_DT", ascending=False)
        return df_meses["MES"].tolist()

    # ---------------------------------------------------------
    # OCORRÊNCIAS POR MÊS
    # ---------------------------------------------------------
    if ocorrencias:
        df_total = pd.concat(ocorrencias, ignore_index=True)
        natureza_col = [c for c in df_total.columns if "NATUREZA" in c.upper()][0]
        df_total[natureza_col] = df_total[natureza_col].astype(str).str.strip()
        excluir = ["escolha um item", "escolher um item."]
        df_total = df_total[~df_total[natureza_col].str.lower().isin(excluir)]
        df_total["MES_DT"] = df_total["MES"].apply(mes_para_datetime)

        for mes in obter_meses_ordenados(df_total):
            df_mes = df_total[df_total["MES"] == mes]
            st.header(f"Mês: {mes}")
            resumo = df_mes.groupby(["PRODUTO", natureza_col]).size().reset_index(name="TOTAL OCORRÊNCIAS")
            resumo["TOTAL OCORRÊNCIAS"] = resumo["TOTAL OCORRÊNCIAS"].astype(int)
            st.subheader("Ocorrências por Natureza")
            st.dataframe(resumo.style.set_properties(**{"font-size": "16px"}), use_container_width=True)

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
        df_horas["MES_DT"] = df_horas["MES"].apply(mes_para_datetime)

        for mes in obter_meses_ordenados(df_horas):
            df_mes = df_horas[df_horas["MES"] == mes]
            st.header(f"Horas Indisponíveis — {mes}")
            total_horas = int(df_mes["HORAS"].sum())
            st.metric("Total de Horas Indisponíveis", total_horas)
            horas_nat = df_mes.groupby(["PRODUTO", col_nat])["HORAS"].sum().reset_index()
            horas_nat["HORAS"] = horas_nat["HORAS"].astype(int)
            st.subheader("Horas por Natureza")
            st.dataframe(horas_nat.style.set_properties(**{"font-size": "16px"}), use_container_width=True)
            horas_eq = df_mes.groupby(["PRODUTO", col_equip])["HORAS"].sum().reset_index()
            horas_eq["HORAS"] = horas_eq["HORAS"].astype(int)
            st.subheader("Horas por Equipamento")
            st.dataframe(horas_eq.style.set_properties(**{"font-size": "16px"}), use_container_width=True)

    # ---------------------------------------------------------
    # VIAGENS E OBJETIVOS POR MÊS (Excel)
    # ---------------------------------------------------------
    if viagens:
        df_viagens = pd.concat(viagens, ignore_index=True)
        df_viagens["MES_DT"] = df_viagens["MES"].apply(mes_para_datetime)

        for mes in obter_meses_ordenados(df_viagens):
            df_mes = df_viagens[df_viagens["MES"] == mes]
            st.header(f"Viagens Realizadas — {mes}")
            st.metric("Total de Viagens", int(df_mes.shape[0]))
            st.metric("Total de Objetivos Traçados", int(df_mes["OBJETIVOS"].sum()))
            st.metric("Total de Objetivos Cumpridos", int(df_mes["OBJETIVOS_CUMPRIDOS"].sum()))
            st.metric("Total de Objetivos Extras", int(df_mes["OBJETIVOS_EXTRAS"].sum()))
            st.metric("Total de Objetivos Extras Cumpridos", int(df_mes["OBJETIVOS_EXTRAS_CUMPRIDOS"].sum()))

else:
    st.info("Aguardando envio dos relatórios.")
