import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="Retrabalhos HidroMeter", layout="wide")
st.title("Análise de Ocorrências - HidroMeter Connect")

uploaded_files = st.file_uploader(
    "Envie arquivos Word",
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


def find_occurrence_table(tables):

    for table in tables:

        header = [str(x).upper() for x in table[0]]

        if "NATUREZA" in header and "OCORRÊNCIA" in header:
            return table

    return None


if uploaded_files:

    registros = []

    for file in uploaded_files:

        try:

            text, tables = extract_text_and_tables(file)

            if not re.search(r"Hidro\s*Meter\s*Connect", text, re.IGNORECASE):
                continue

            occ_table = find_occurrence_table(tables)

            if occ_table is None:
                continue

            header = occ_table[0]
            df = pd.DataFrame(occ_table[1:], columns=header)

            registros.append(df)

        except Exception as e:
            st.warning(f"Erro ao processar {file.name}: {e}")

    if len(registros) > 0:

        df_total = pd.concat(registros, ignore_index=True)

        natureza_col = [c for c in df_total.columns if "NATUREZA" in c.upper()][0]

        # encontrar coluna de horas automaticamente
        horas_col = None
        for c in df_total.columns:
            if "hora" in c.lower() or "indispon" in c.lower():
                horas_col = c
                break

        df_total[natureza_col] = df_total[natureza_col].astype(str).str.strip()

        excluir = ["escolha um item", "escolher um item."]

        df_total = df_total[
            ~df_total[natureza_col].str.lower().isin(excluir)
        ]

        resumo = (
            df_total
            .groupby(natureza_col)
            .size()
            .reset_index(name="Total de Ocorrências")
        )

        st.subheader("Total de Ocorrências por Natureza")
        st.dataframe(resumo)

        fig = px.bar(
            resumo,
            x=natureza_col,
            y="Total de Ocorrências",
            text="Total de Ocorrências",
            color=natureza_col
        )

        fig.update_layout(
            showlegend=False,
            xaxis_title="Natureza",
            yaxis_title="Total de Ocorrências"
        )

        st.plotly_chart(fig, use_container_width=True)

        # --- HORAS INDISPONÍVEIS ---

        if horas_col:

            df_total[horas_col] = pd.to_numeric(df_total[horas_col], errors="coerce").fillna(0)

            total_horas = df_total[horas_col].sum()

            st.subheader("Total de Horas Indisponíveis")
            st.metric("Horas Totais", round(total_horas, 2))

            horas_natureza = (
                df_total
                .groupby(natureza_col)[horas_col]
                .sum()
                .reset_index()
            )

            fig2 = px.bar(
                horas_natureza,
                x=natureza_col,
                y=horas_col,
                text=horas_col,
                color=natureza_col,
                title="Horas Indisponíveis por Natureza"
            )

            fig2.update_layout(
                showlegend=False,
                xaxis_title="Natureza",
                yaxis_title="Horas Indisponíveis"
            )

            st.plotly_chart(fig2, use_container_width=True)

        else:

            st.warning("Nenhuma coluna de horas indisponíveis foi encontrada nos relatórios.")

    else:
        st.warning("Nenhuma ocorrência encontrada para HidroMeter Connect.")

else:
    st.info("Aguardando envio de arquivos Word.")
