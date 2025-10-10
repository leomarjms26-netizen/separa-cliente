import io
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Gerar planilha por cliente (com dados preenchidos)", layout="wide")
st.title("📊 Automatizador — Gerar planilha final por cliente")

uploaded_raw = st.file_uploader("1️⃣ Envie a planilha **BRUTA** (dados originais)", type=["xlsx", "xls"])
uploaded_model = st.file_uploader("2️⃣ Envie o **MODELO** (planilha de exemplo)", type=["xlsx", "xls"])

if uploaded_raw and uploaded_model:
    df = pd.read_excel(uploaded_raw)
    st.success(f"✅ Planilha bruta carregada com {len(df)} linhas e {len(df.columns)} colunas.")
    st.dataframe(df.head())

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

        # Garantir que a coluna do cliente existe
        if coluna_cliente not in df.columns:
            st.error("❌ A coluna selecionada não existe na planilha.")
            st.stop()

        # Loop pelos clientes únicos
        for cliente, dados_cliente in df.groupby(coluna_cliente):
            nome_aba = str(cliente).strip()
            if nome_aba == "" or nome_aba.lower() == "nan":
                nome_aba = "SemNome"

            # Corrige nome da aba
            nome_aba = re.sub(r'[\\/*?:\[\]]', "-", nome_aba)[:31]

            # Copia a aba modelo
            nova_aba = wb.copy_worksheet(modelo)
            nova_aba.title = nome_aba

            # Escrever cabeçalho
            for c_idx, coluna in enumerate(dados_cliente.columns, start=1):
                nova_aba.cell(row=1, column=c_idx, value=coluna)

            # Escrever os dados (logo abaixo do cabeçalho)
            for r_idx, (_, linha) in enumerate(dados_cliente.iterrows(), start=2):
                for c_idx, valor in enumerate(linha, start=1):
                    nova_aba.cell(row=r_idx, column=c_idx, value=valor)

        # Remove aba modelo original
        wb.remove(modelo)

        # Salvar em memória
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ Planilha gerada com sucesso! Cada aba contém os dados do cliente correspondente.")
        st.download_button(
            label="⬇️ Baixar planilha final",
            data=output,
            file_name="Planilha_Final_Clientes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
