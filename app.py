import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re

st.set_page_config(page_title="Analisador Automático de Relatórios - CVT", layout="wide")
st.title("Analisador Automático de Relatórios - CVT")
st.image("data:image/png;base64,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", use_column_width=False)

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







