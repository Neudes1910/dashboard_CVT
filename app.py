import streamlit as st
import pandas as pd
from docx import Document
from collections import defaultdict

st.title("Análise de Ocorrências - HidroMeter Connect")

arquivos = st.file_uploader(
    "Envie os arquivos Word",
    type=["docx","docm"],
    accept_multiple_files=True
)

def texto_documento(doc):
    textos = []

    for p in doc.paragraphs:
        textos.append(p.text)

    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                textos.append(celula.text)

    return " ".join(textos).lower()


def encontrar_tabela_ocorrencias(doc):

    for tabela in doc.tables:

        cabecalho = [c.text.strip().lower() for c in tabela.rows[0].cells]

        if "natureza" in cabecalho and "ocorrência" in cabecalho:
            return tabela

    return None


def encontrar_tabela_horas(doc):

    for tabela in doc.tables:

        cabecalho = [c.text.strip().lower() for c in tabela.rows[0].cells]

        if "equipamento" in cabecalho and "hora" in " ".join(cabecalho):
            return tabela

    return None


def processar_doc(doc):

    natureza_contagem = defaultdict(int)
    horas_equipamento = defaultdict(float)
    total_horas = 0

    tabela_oc = encontrar_tabela_ocorrencias(doc)

    if tabela_oc:

        cab = [c.text.strip().lower() for c in tabela_oc.rows[0].cells]

        idx_nat = cab.index("natureza")

        for linha in tabela_oc.rows[1:]:

            natureza = linha.cells[idx_nat].text.strip()

            if natureza:
                natureza_contagem[natureza] += 1


    tabela_hr = encontrar_tabela_horas(doc)

    if tabela_hr:

        cab = [c.text.strip().lower() for c in tabela_hr.rows[0].cells]

        idx_eq = cab.index("equipamento")

        idx_hr = None

        for i, c in enumerate(cab):
            if "hora" in c:
                idx_hr = i

        if idx_hr is not None:

            for linha in tabela_hr.rows[1:]:

                equipamento = linha.cells[idx_eq].text.strip()

                horas = linha.cells[idx_hr].text.replace(",", ".").strip()

                try:
                    horas = float(horas)
                except:
                    horas = 0

                horas_equipamento[equipamento] += horas
                total_horas += horas

    return natureza_contagem, horas_equipamento, total_horas


if arquivos:

    total_natureza = defaultdict(int)
    total_horas_equip = defaultdict(float)
    horas_total = 0

    arquivos_processados = 0

    for arquivo in arquivos:

        doc = Document(arquivo)

        texto = texto_documento(doc)

        if "hidrometer connect" in texto:

            nat, eq, hrs = processar_doc(doc)

            for k, v in nat.items():
                total_natureza[k] += v

            for k, v in eq.items():
                total_horas_equip[k] += v

            horas_total += hrs

            arquivos_processados += 1


    st.subheader("Arquivos analisados")
    st.write(arquivos_processados)

    st.subheader("Ocorrências por Natureza")

    if total_natureza:

        df_nat = pd.DataFrame(
            total_natureza.items(),
            columns=["Natureza", "Total de Ocorrências"]
        )

        st.dataframe(df_nat)

    else:
        st.write("Nenhuma ocorrência encontrada.")


    st.subheader("Horas indisponíveis por Equipamento")

    if total_horas_equip:

        df_eq = pd.DataFrame(
            total_horas_equip.items(),
            columns=["Equipamento", "Horas Indisponíveis"]
        )

        st.dataframe(df_eq)

    else:
        st.write("Nenhuma hora registrada.")


    st.subheader("Total de Horas Indisponíveis")
    st.write(horas_total)

