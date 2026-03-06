import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd

st.set_page_config(page_title="Analisador Automático de Relatórios - CVT", layout="wide")
st.title("Analisador Automático de Relatórios - CVT")

uploaded_files = st.file_uploader(
    "Envie os relatórios (.docm ou .docx)",
    type=["docx","docm"],
    accept_multiple_files=True
)


def limpar(txt):
    return txt.lower().replace("\n"," ").strip()


def extrair_tabelas(arquivo):

    with zipfile.ZipFile(arquivo) as doc:
        xml = doc.read("word/document.xml")

    root = ET.fromstring(xml)

    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    tabelas = []

    for tbl in root.findall('.//w:tbl', ns):

        dados = []

        for tr in tbl.findall('.//w:tr', ns):

            linha = []

            for tc in tr.findall('.//w:tc', ns):

                textos = [t.text for t in tc.findall('.//w:t', ns) if t.text]
                linha.append(" ".join(textos))

            dados.append(linha)

        tabelas.append(dados)

    return tabelas


def documento_tem_hidrometer(tabelas):

    for tabela in tabelas:
        for linha in tabela:
            for cel in linha:
                if "hidrometer connect" in limpar(cel):
                    return True

    return False


if uploaded_files:

    naturezas = []
    indisponibilidade = []

    for arquivo in uploaded_files:

        try:
            tabelas = extrair_tabelas(arquivo)
        except:
            continue

        if not documento_tem_hidrometer(tabelas):
            continue

        for tabela in tabelas:

            if len(tabela) < 2:
                continue

            header = [limpar(c) for c in tabela[0]]

            # -----------------------------
            # TABELA DE OCORRÊNCIAS
            # -----------------------------

            if any("natureza" in h for h in header) and any("ocorr" in h for h in header):

                idx_nat = next(i for i,h in enumerate(header) if "natureza" in h)

                for row in tabela[1:]:

                    if idx_nat >= len(row):
                        continue

                    nat = limpar(row[idx_nat])

                    if nat in ["", "escolha um item", "escolher um item."]:
                        continue

                    naturezas.append(nat)

            # -----------------------------
            # TABELA DE HORAS
            # -----------------------------

            if (
                any("tempo" in h for h in header)
                and any("equipamento" in h for h in header)
            ):

                idx_tempo = next(i for i,h in enumerate(header) if "tempo" in h)
                idx_eq = next(i for i,h in enumerate(header) if "equipamento" in h)
                idx_nat = next(i for i,h in enumerate(header) if "natureza" in h)

                for row in tabela[1:]:

                    if max(idx_tempo, idx_eq, idx_nat) >= len(row):
                        continue

                    horas_txt = limpar(row[idx_tempo])
                    equipamento = limpar(row[idx_eq])
                    natureza = limpar(row[idx_nat])

                    if natureza in ["", "escolha um item", "escolher um item."]:
                        continue

                    try:
                        horas = float(horas_txt.replace(",","."))
                    except:
                        horas = 0

                    indisponibilidade.append({
                        "equipamento": equipamento,
                        "natureza": natureza,
                        "horas": horas
                    })


    if not naturezas and not indisponibilidade:

        st.warning("Nenhum registro encontrado nos arquivos.")
        st.stop()


    # -----------------------------
    # OCORRÊNCIAS POR NATUREZA
    # -----------------------------

    if naturezas:

        df = pd.DataFrame({"Natureza": naturezas})

        contagem = df.value_counts().reset_index()
        contagem.columns = ["Natureza","Total"]

        st.subheader("Total de Ocorrências por Natureza")

        st.bar_chart(contagem.set_index("Natureza"))


    # -----------------------------
    # HORAS INDISPONÍVEIS
    # -----------------------------

    if indisponibilidade:

        df = pd.DataFrame(indisponibilidade)

        total = df["horas"].sum()

        st.subheader("Total de Horas Indisponíveis")

        st.metric("Horas totais", round(total,2))

        col1, col2 = st.columns(2)

        horas_nat = df.groupby("natureza")["horas"].sum()

        with col1:
            st.subheader("Horas por Natureza")
            st.bar_chart(horas_nat)

        horas_eq = df.groupby("equipamento")["horas"].sum()

        with col2:
            st.subheader("Horas por Equipamento")
            st.bar_chart(horas_eq)

else:

    st.info("Envie os relatórios para análise.")
