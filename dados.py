import io
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Gerar planilha por cliente (corrigida)", layout="wide")
st.title("📊 Automatizador — Gerar planilha final por cliente (versão revisada)")

uploaded_raw = st.file_uploader("1️⃣ Envie a planilha **BRUTA** (dados originais)", type=["xlsx", "xls"])
uploaded_model = st.file_uploader("2️⃣ Envie o **MODELO** (planilha base)", type=["xlsx", "xls"])

if uploaded_raw and uploaded_model:
    # Lê os dados brutos
    df = pd.read_excel(uploaded_raw)
    st.write("🧩 Colunas detectadas na planilha bruta:")
    st.dataframe(pd.DataFrame({"Colunas": df.columns}))

    # Normaliza nomes de colunas (remove espaços e padroniza maiúsculas)
    df.columns = df.columns.str.strip()

    # Detectar coluna do cliente
    colunas_cliente = [c for c in df.columns if "cliente" in c.lower() or "processo" in c.lower() or "contrato" in c.lower()]
    if not colunas_cliente:
        colunas_cliente = list(df.columns)

    coluna_cliente = st.selectbox("📌 Escolha a coluna que identifica o CLIENTE:", options=colunas_cliente)

    # Carregar o modelo
    wb = load_workbook(uploaded_model)
    modelo_aba = st.selectbox("📄 Escolha a aba modelo para copiar:", options=wb.sheetnames)

    if st.button("🚀 Gerar planilha final"):
        modelo = wb[modelo_aba]

        if coluna_cliente not in df.columns:
            st.error(f"❌ A coluna '{coluna_cliente}' não foi encontrada nas colunas da planilha.")
            st.stop()

        if df[coluna_cliente].isnull().all():
            st.error("❌ Todos os valores da coluna de cliente estão vazios — não há como agrupar.")
            st.stop()

        # Remove espaços e normaliza valores de cliente
        df[coluna_cliente] = df[coluna_cliente].astype(str).str.strip()

        # Loop pelos clientes únicos
        for cliente, dados_cliente in df.groupby(coluna_cliente):
            if not cliente or cliente.lower() in ["nan", "none"]:
                cliente = "SemNome"

            nome_aba = re.sub(r'[\\/*?:\[\]]', "-", str(cliente))[:31]
            nova_aba = wb.copy_worksheet(modelo)
            nova_aba.title = nome_aba

            # Escrever cabeçalhos e dados
            for c_idx, coluna in enumerate(dados_cliente.columns, start=1):
                nova_aba.cell(row=1, column=c_idx, value=coluna)

            for r_idx, (_, linha) in enumerate(dados_cliente.iterrows(), start=2):
                for c_idx, valor in enumerate(linha, start=1):
                    nova_aba.cell(row=r_idx, column=c_idx, value=valor)

        wb.remove(modelo)

        # Salvar resultado
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ Planilha criada com sucesso! Cada aba contém os dados do cliente correspondente.")
        st.download_button(
            label="⬇️ Baixar planilha final",
            data=output,
            file_name="Planilha_Final_Clientes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
