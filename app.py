import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import plotly.express as px
import re

st.set_page_config(page_title="Retrabalho HidroMeter Connect", layout="wide")
st.title("Retrabalhos e Horas Indisponíveis - HidroMeter Connect")

uploaded_files = st.file_uploader(
    "Arraste e solte arquivos Word (.docm) aqui",
    type=["docm"],
    accept_multiple_files=True
)

def extract_table_from_docm(file, index=1):
    """
    Extrai a tabela de índice 'index' do arquivo .docm.
    index=0 para primeira tabela, index=1 para segunda, etc.
    """
    try:
        with zipfile.ZipFile(file) as docm:
            xml_content = docm.read("word/document.xml")
        root = ET.fromstring(xml_content)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        tables = root.findall('.//w:tbl', ns)
        if len(tables) <= index:
            return None

        tbl = tables[index]
        table_data = []
        for row in tbl.findall('.//w:tr', ns):
            cells = []
            for cell in row.findall('.//w:tc', ns):
                texts = [t.text for t in cell.findall('.//w:t', ns) if t.text]
                cell_text = " ".join(texts).strip()
                cells.append(cell_text)
            if cells:
                table_data.append(cells)
        return table_data
    except Exception as e:
        st.error(f"Erro ao processar {file.name}: {e}")
        return None

def is_hidrometer_connect(table_data):
    """
    Verifica se a tabela contém o produto HidroMeter Connect.
    """
    for row in table_data:
        for cell in row:
            if re.match(r"Produto\s*---\s*HidroMeter Connect", cell, re.IGNORECASE):
                return True
    return False

if uploaded_files:
    retrabalho_list = []

    for file in uploaded_files:
        table_data = extract_table_from_docm(file, index=1)  # segunda tabela
        if not table_data:
            st.warning(f"Nenhuma tabela encontrada no índice especificado em {file.name}.")
            continue

        if not is_hidrometer_connect(table_data):
            st.info(f"O arquivo {file.name} não contém o produto HidroMeter Connect na tabela selecionada.")
            continue

        # Considera primeira linha como header
        df = pd.DataFrame(table_data[1:], columns=table_data[0])

        # Verifica se existem colunas necessárias
        required_cols = ["Natureza", "Horas Indisponíveis", "Equipamento"]
        if all(col in df.columns for col in required_cols):
            # Converte Horas Indisponíveis para número
            df["Horas Indisponíveis"] = pd.to_numeric(df["Horas Indisponíveis"], errors='coerce').fillna(0)
            retrabalho_list.append(df)
        else:
            st.warning(f"As colunas {required_cols} não foram encontradas em {file.name}.")

    if retrabalho_list:
        retrabalho_df = pd.concat(retrabalho_list, ignore_index=True)

        # Total de ocorrências por natureza
        ocorrencias_por_natureza = retrabalho_df.groupby("Natureza").size().reset_index(name="Total Ocorrências")

        # Total de horas indisponíveis (todas as naturezas)
        total_horas = retrabalho_df["Horas Indisponíveis"].sum()

        # Horas indisponíveis por equipamento
        horas_por_equipamento = retrabalho_df.groupby("Equipamento")["Horas Indisponíveis"].sum().reset_index()

        # Exibição
        st.subheader("Total de Ocorrências por Natureza")
        st.dataframe(ocorrencias_por_natureza)

        st.subheader("Total de Horas Indisponíveis")
        st.metric(label="Horas Indisponíveis", value=total_horas)

        st.subheader("Horas Indisponíveis por Equipamento")
        st.dataframe(horas_por_equipamento)

        # Gráfico de ocorrências por natureza
        fig1 = px.bar(
            ocorrencias_por_natureza,
            x="Natureza",
            y="Total Ocorrências",
            text="Total Ocorrências",
            color="Natureza",
            title="Ocorrências por Natureza - HidroMeter Connect"
        )
        fig1.update_layout(showlegend=False, xaxis_title="Natureza", yaxis_title="Total Ocorrências")
        st.plotly_chart(fig1, use_container_width=True)

        # Gráfico de horas por equipamento
        fig2 = px.bar(
            horas_por_equipamento,
            x="Equipamento",
            y="Horas Indisponíveis",
            text="Horas Indisponíveis",
            color="Equipamento",
            title="Horas Indisponíveis por Equipamento - HidroMeter Connect"
        )
        fig2.update_layout(showlegend=False, xaxis_title="Equipamento", yaxis_title="Horas Indisponíveis")
        st.plotly_chart(fig2, use_container_width=True)

    else:
        st.info("Nenhum dado encontrado para HidroMeter Connect.")
else:
    st.info("Aguardando arquivos .docm para upload.")
