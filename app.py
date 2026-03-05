import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Contagem de Arquivos por Projeto", layout="wide")

st.title("Dashboard de Contagem de Arquivos")

# Upload de múltiplos arquivos
uploaded_files = st.file_uploader(
    "Arraste e solte os arquivos aqui",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []

    for file in uploaded_files:
        try:
            if file.name.endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)
        except Exception as e:
            st.error(f"Erro ao ler o arquivo {file.name}: {e}")
            continue

        # Aqui assumimos que há uma coluna 'Projeto' e 'Quantidade'
        if "Projeto" not in df.columns or "Quantidade" not in df.columns:
            st.warning(f"O arquivo {file.name} não possui as colunas necessárias: 'Projeto' e 'Quantidade'.")
            continue

        all_data.append(df)

    if all_data:
        data = pd.concat(all_data, ignore_index=True)

        st.subheader("Dados Consolidados")
        st.dataframe(data)

        # Agrupando por projeto
        project_summary = data.groupby("Projeto")["Quantidade"].sum().reset_index()

        st.subheader("Gráfico por Projeto")
        fig = px.bar(project_summary, x="Projeto", y="Quantidade", text="Quantidade", color="Projeto")
        fig.update_layout(showlegend=False, xaxis_title="Projeto", yaxis_title="Total de Itens")
        st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Aguardando arquivos para upload.")
