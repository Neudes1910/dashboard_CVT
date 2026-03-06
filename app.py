import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="Analisador Automático de Relatórios - CVT", layout="wide")
st.title("Analisador Automático de Relatórios - CVT")

uploaded_files = st.file_uploader(
    "Envie os relatórios (.docx ou .docm)",
    type=["docx","docm"],
    accept_multiple_files=True
)


def limpar(txt):
    return txt.lower().replace("\n"," ").strip()


def tem_hidrometer(doc):

    for p in doc.paragraphs:
        if "hidrometer connect" in limpar(p.text):
            return True

    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                if "hidrometer connect" in limpar(c.text):
                    return True

    return False


def extrair_ocorrencias(doc):

    lista = []

    for tabela in doc.tables:

        header = [limpar(c.text) for c in tabela.rows[0].cells]

        if any("natureza" in h for h in header) and any("ocorr" in h for h in header):

            idx_nat = next(i for i,h in enumerate(header) if "natureza" in h)

            for row in tabela.rows[1:]:

                natureza = limpar(row.cells[idx_nat].text)

                if natureza in ["", "escolha um item", "escolher um item."]:
                    continue

                lista.append(natureza)

    return lista


def extrair_indisponibilidade(doc):

    registros = []

    for tabela in doc.tables:

        header = [limpar(c.text) for c in tabela.rows[0].cells]

        if (
            any("tempo" in h for h in header) and
            any("equipamento" in h for h in header)
        ):

            idx_horas = next(i for i,h in enumerate(header) if "tempo" in h)
            idx_equip = next(i for i,h in enumerate(header) if "equipamento" in h)
            idx_nat = next(i for i,h in enumerate(header) if "natureza" in h)

            for row in tabela.rows[1:]:

                horas_txt = limpar(row.cells[idx_horas].text)
                equipamento = limpar(row.cells[idx_equip].text)
                natureza = limpar(row.cells[idx_nat].text)

                if natureza in ["", "escolha um item", "escolher um item."]:
                    continue

                try:
                    horas = float(horas_txt.replace(",","."))
                except:
                    horas = 0

                registros.append({
                    "equipamento": equipamento,
                    "natureza": natureza,
                    "horas": horas
                })

    return registros


if uploaded_files:

    ocorrencias = []
    indisponibilidades = []

    for arquivo in uploaded_files:

        try:
            doc = Document(io.BytesIO(arquivo.read()))
        except:
            continue

        if not tem_hidrometer(doc):
            continue

        ocorrencias += extrair_ocorrencias(doc)
        indisponibilidades += extrair_indisponibilidade(doc)

    if not ocorrencias and not indisponibilidades:
        st.warning("Nenhum registro encontrado nos arquivos.")
        st.stop()

    col1, col2 = st.columns(2)

    if ocorrencias:

        df = pd.DataFrame({"Natureza": ocorrencias})

        contagem = df.value_counts().reset_index()
        contagem.columns = ["Natureza","Total"]

        with col1:
            st.subheader("Total de Ocorrências por Natureza")
            st.bar_chart(contagem.set_index("Natureza"))

    if indisponibilidades:

        df = pd.DataFrame(indisponibilidades)

        total_horas = df["horas"].sum()

        st.subheader("Total de Horas Indisponíveis")
        st.metric("Horas totais", round(total_horas,2))

        col3, col4 = st.columns(2)

        horas_nat = df.groupby("natureza")["horas"].sum()

        with col3:
            st.subheader("Horas por Natureza")
            st.bar_chart(horas_nat)

        horas_eq = df.groupby("equipamento")["horas"].sum()

        with col4:
            st.subheader("Horas por Equipamento")
            st.bar_chart(horas_eq)
