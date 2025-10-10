import io
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ---------------- CONFIGURAÇÕES ----------------
MODELO_PATH = "PLANILHA GERAL CLIENTES MENSAIS - ATUALIZADA SET25 (1).xlsx"  # nome do arquivo do modelo (fixo)
MODELO_ABA = "GERAL CLIENTES"        # nome da aba dentro do modelo que servirá de base
# ------------------------------------------------

st.set_page_config(page_title="Gerar planilha por cliente", layout="wide")
st.title("📊 Automatizador — Gerar planilha final por cliente")

# Upload apenas da planilha bruta
uploaded_raw = st.file_uploader("📤 Envie a planilha **BRUTA** (dados originais)", type=["xlsx", "xls"])

if uploaded_raw:
    # Lê os dados brutos
    df = pd.read_excel(uploaded_raw)
    df.columns = df.columns.str.strip()

    st.write("🧩 Colunas detectadas:")
    st.dataframe(pd.DataFrame({"Colunas": df.columns}))

    # Detectar coluna do cliente
    colunas_cliente = [c for c in df.columns if "cliente" in c.lower() or "processo" in c.lower() or "contrato" in c.lower()]
    if not colunas_cliente:
        colunas_cliente = list(df.columns)

    coluna_cliente = st.selectbox("📌 Escolha a coluna que identifica o CLIENTE:", options=colunas_cliente)

    if st.button("🚀 Gerar planilha final"):
        try:
            wb = load_workbook(MODELO_PATH)
        except Exception as e:
            st.error(f"❌ Não foi possível carregar o modelo fixo: {e}")
            st.stop()

        if MODELO_ABA not in wb.sheetnames:
            st.error(f"❌ A aba '{MODELO_ABA}' não foi encontrada no modelo.")
            st.stop()

        modelo = wb[MODELO_ABA]

        df[coluna_cliente] = df[coluna_cliente].astype(str).str.strip()

        for cliente, dados_cliente in df.groupby(coluna_cliente):
            if not cliente or cliente.lower() in ["nan", "none"]:
                cliente = "SemNome"

            nome_aba = re.sub(r'[\\/*?:\[\]]', "-", str(cliente))[:31]
            nova_aba = wb.copy_worksheet(modelo)
            nova_aba.title = nome_aba

            # Escrever dados
            for c_idx, coluna in enumerate(dados_cliente.columns, start=1):
                nova_aba.cell(row=1, column=c_idx, value=coluna)

            for r_idx, (_, linha) in enumerate(dados_cliente.iterrows(), start=2):
                for c_idx, valor in enumerate(linha, start=1):
                    nova_aba.cell(row=r_idx, column=c_idx, value=valor)

        # Remove aba modelo original
        wb.remove(modelo)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("✅ Planilha gerada com sucesso! Cada aba contém os dados do cliente.")
        st.download_button(
            label="⬇️ Baixar planilha final",
            data=output,
            file_name="Planilha_Final_Clientes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


