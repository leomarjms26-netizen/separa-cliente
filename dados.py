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
    # Lê os dados brutos
    df = pd.read_excel(uploaded_raw)
    st.success(f"✅ Planilha bruta carregada com {len(df)} linhas e {len(df.columns)} colunas.")
    st.dataframe(df.head())

    # Detectar coluna do cliente
    colunas_cliente = [c for c in df.columns if "cliente" in c.lower() or "processo" in c.lower() or "contrato" in c.lower()]
    coluna_cliente = st.selectbox("📌 Escolha a coluna que identifica o CLIENTE:", options=colunas_cliente)

    # Carregar o modelo
    wb = load_workbook(uploaded_model)
    modelo_aba = st.selectbox("📄 Escolha a aba modelo para copiar:", options=wb.sheetnames)

    if st.button("🚀 Gerar planilha final"):
        modelo = wb[modelo_aba]

        # Loop pelos clientes únicos
        for cliente, dados_cliente in df.groupby(coluna_cliente):
            nome_aba = str(cliente).strip()
            if nome_aba == "" or nome_aba.lower() == "nan":
                nome_aba = "SemNome"

            # Corrige caracteres inválidos para nome de aba
            nome_aba = re.sub(r'[\\/*?:\[\]]', "-", nome_aba)
            nome_aba = nome_aba[:31]  # limite Excel

            # Copia a aba modelo
            nova_aba = wb.copy_worksheet(modelo)
            nova_aba.title = nome_aba

            # Acha a próxima linha vazia (para não sobrescrever o cabeçalho)
            linha_inicio = nova_aba.max_row + 1

            # Escreve o DataFrame do cliente na aba
            for r_idx, row in enumerate(dados_cliente.itertuples(index=False), start=linha_inicio):
                for c_idx, valor in enumerate(row, start=1):
                    nova_aba.cell(row=r_idx, column=c_idx, value=valor)

        # Remove a aba modelo original (opcional)
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
