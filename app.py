import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="Analisador Automático de Relatórios - CVT", layout="wide")
st.title("Analisador Automático de Relatórios - CVT")

uploaded_files = st.file_uploader(
    "Envie os relatórios Word",
    type=["docx", "docm", "dotm"],
    accept_multiple_files=True
)

NAMESPACE = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def extract_text_and_tables(file):

    text_content = []
    tables = []

    with zipfile.ZipFile(file) as doc:
        xml_content = doc.read("word/document.xml")

    root = ET.fromstring(xml_content)

    for t in root.findall('.//w:t', NAMESPACE):
        if t.text:
            text_content.append(t.text)

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


def extrair_produto(texto):

    linhas = texto.split()

    for i, palavra in enumerate(linhas):

        if palavra.lower() == "produto:" or palavra.lower() == "produto":

            try:
                produto = linhas[i+1] + " " + linhas[i+2]
                return produto.strip()
            except:
                return linhas[i+1]

    return "Produto não identificado"


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

    texto = str(valor).lower()

    match = re.search(r'(\d+[.,]?\d*)', texto)

    if match:
        return float(match.group(1).replace(",", "."))

    return 0


if uploaded_files:

    ocorrencias = []
    horas_registros = []

    for file in uploaded_files:

        try:

            text, tables = extract_text_and_tables(file)

            produto = extrair_produto(text)

            occ_table = find_occurrence_table(tables)

            if occ_table:
                df_occ = pd.DataFrame(occ_table[1:], columns=occ_table[0])
                df_occ["PRODUTO"] = produto
                ocorrencias.append(df_occ)

            downtime_table = find_downtime_table(tables)

            if downtime_table:
                df_down = pd.DataFrame(downtime_table[1:], columns=downtime_table[0])
                df_down["PRODUTO"] = produto
                horas_registros.append(df_down)

        except Exception as e:
            st.warning(f"Erro ao processar {file.name}: {e}")

    # -------------------------
    # OCORRÊNCIAS
    # -------------------------

    if ocorrencias:

        df_total = pd.concat(ocorrencias, ignore_index=True)

        natureza_col = [c for c in df_total.columns if "NATUREZA" in c.upper()][0]

        df_total[natureza_col] = df_total[natureza_col].astype(str).str.strip()

        excluir = ["escolha um item", "escolher um item."]

        df_total = df_total[
            ~df_total[natureza_col].str.lower().isin(excluir)
        ]

        resumo = (
            df_total
            .groupby(["PRODUTO", natureza_col])
            .size()
            .reset_index(name="Total de Ocorrências")
        )

        st.subheader("Total de Ocorrências por Natureza")

        fig = px.bar(
            resumo,
            x=natureza_col,
            y="Total de Ocorrências",
            color="PRODUTO",
            barmode="group",
            text="Total de Ocorrências"
        )

        st.plotly_chart(fig, use_container_width=True)

    # -------------------------
    # HORAS DE INDISPONIBILIDADE
    # -------------------------

    if horas_registros:

        df_horas = pd.concat(horas_registros, ignore_index=True)

        df_horas.columns = (
            df_horas.columns
            .str.replace("\n", " ")
            .str.strip()
        )

        col_tempo = next(c for c in df_horas.columns if "TEMPO" in c.upper())
        col_equip = next(c for c in df_horas.columns if "QUAL EQUIPAMENTO" in c.upper())
        col_nat = next(c for c in df_horas.columns if "NATUREZA" in c.upper())

        df_horas["HORAS"] = df_horas[col_tempo].apply(converter_horas)

        df_horas[col_nat] = df_horas[col_nat].astype(str).str.strip()

        df_horas = df_horas[
            ~df_horas[col_nat].str.lower().isin(["escolha um item", "escolher um item."])
        ]

        df_horas[col_equip] = df_horas[col_equip].astype(str).str.strip()

        total_horas = df_horas["HORAS"].sum()

        st.subheader("Horas Indisponíveis Totais")
        st.metric("Total de Horas", round(total_horas, 2))

        # horas por natureza

        horas_nat = (
            df_horas
            .groupby(["PRODUTO", col_nat])["HORAS"]
            .sum()
            .reset_index()
        )

        fig2 = px.bar(
            horas_nat,
            x=col_nat,
            y="HORAS",
            color="PRODUTO",
            barmode="group",
            text="HORAS",
            title="Horas Indisponíveis por Natureza"
        )

        st.plotly_chart(fig2, use_container_width=True)

        # horas por equipamento

        horas_eq = (
            df_horas
            .groupby(["PRODUTO", col_equip])["HORAS"]
            .sum()
            .reset_index()
        )

        fig3 = px.bar(
            horas_eq,
            x=col_equip,
            y="HORAS",
            color="PRODUTO",
            barmode="group",
            text="HORAS",
            title="Horas Indisponíveis por Equipamento"
        )

        st.plotly_chart(fig3, use_container_width=True)

else:
    st.info("Aguardando envio dos relatórios.")

