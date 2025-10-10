# app.py — Streamlit: processador que usa modelo fixo e detecta cabeçalho automaticamente
import io, re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# ---------- CONFIG ----------
MODELO_PATH = "TS Setembro 2025).xlsx"
MODELO_ABA = "Geral"
# ----------------------------

st.set_page_config(page_title="Gerar planilha por cliente", layout="wide")
st.title("📊 Automatizador — Gerar planilha final por cliente")

uploaded_raw = st.file_uploader("📤 Envie a planilha BRUTA (xlsx/xls)", type=["xlsx", "xls"])

if uploaded_raw:
    st.info("Detectando cabeçalho e lendo arquivo...")
    # Detect header row (procura uma linha com palavras-chave como 'cliente'/'processo'/'contrato')
    temp = pd.read_excel(uploaded_raw, header=None)
    header_row = None
    for i, row in temp.iterrows():
        if any(isinstance(v, str) and re.search(r"(cliente|processo|contrato)", v, re.IGNORECASE) for v in row.values):
            header_row = i
            break
    if header_row is None:
        header_row = 0
        st.warning("Não foi possível detectar automaticamente o cabeçalho — usando a primeira linha.")

    df = pd.read_excel(uploaded_raw, header=header_row)
    df.columns = df.columns.astype(str).str.strip()
    st.write("Colunas detectadas:")
    st.dataframe(pd.DataFrame({"Colunas": df.columns}))

    # detect candidate client columns
    candidates = [c for c in df.columns if re.search(r"(cliente|processo|contrato)", c, re.IGNORECASE)]
    if not candidates:
        candidates = list(df.columns)
    client_col = st.selectbox("Coluna que identifica o CLIENTE:", options=candidates, index=0)

    keep_model_tab = st.checkbox("Manter a aba modelo (não remover 'Geral')", value=True)
    start_row_override = st.number_input("Linha onde os dados devem começar na aba modelo (opcional)", min_value=1, value=None if header_row is None else None)

    if st.button("🚀 Gerar planilha final"):
        # carregar modelo
        try:
            wb = load_workbook(MODELO_PATH)
        except Exception as e:
            st.error(f"Erro ao carregar modelo: {e}")
            st.stop()

        if MODELO_ABA not in wb.sheetnames:
            st.error(f"A aba '{MODELO_ABA}' não existe no modelo. Abas disponíveis: {wb.sheetnames}")
            st.stop()

        template_ws = wb[MODELO_ABA]

        # normaliza coluna cliente
        df[client_col] = df[client_col].fillna('').astype(str).str.strip()
        df[client_col] = df[client_col].replace({'': 'SemNome', 'nan': 'SemNome', 'None': 'SemNome', 'none': 'SemNome'})

        # encontra header row no template (heurística)
        def find_model_header(ws, df_cols, max_scan=25):
            dfset = set([c.strip().lower() for c in df_cols])
            best_row = 1
            best_count = 0
            max_row = min(ws.max_row, max_scan)
            for r in range(1, max_row+1):
                row_vals = [str(ws.cell(row=r, column=c).value).strip().lower() if ws.cell(row=r, column=c).value is not None else "" for c in range(1, ws.max_column+1)]
                count = sum(1 for v in row_vals if v in dfset and v!="")
                if count > best_count:
                    best_count = count
                    best_row = r
            return best_row, best_count

        model_header_row, matches = find_model_header(template_ws, df.columns)
        st.write(f"Cabeçalho do modelo detectado na linha {model_header_row} (matches: {matches})")

        # lista de colunas do modelo (na linha detectada)
        model_headers = [str(template_ws.cell(row=model_header_row, column=c).value).strip() if template_ws.cell(row=model_header_row, column=c).value is not None else "" for c in range(1, template_ws.max_column+1)]

        # mapear colunas do modelo para as colunas do DataFrame (ignorar espaços/maiúsculas)
        model_idx_to_dfcol = {}
        for idx, mh in enumerate(model_headers, start=1):
            if not mh:
                continue
            mh_norm = re.sub(r'\s+', '', str(mh)).lower()
            for dfcol in df.columns:
                if re.sub(r'\s+', '', str(dfcol)).lower() == mh_norm:
                    model_idx_to_dfcol[idx] = dfcol
                    break

        mapped_dfcols = set(model_idx_to_dfcol.values())
        extra_dfcols = [c for c in df.columns if c not in mapped_dfcols]

        clients = df[client_col].dropna().unique().tolist()
        st.write(f"Clientes detectados: {len(clients)} (ex.: {clients[:8]})")

        created = []
        for cliente in clients:
            dados_cliente = df[df[client_col] == cliente]
            if dados_cliente.empty:
                continue

            # copia a aba modelo
            new_ws = wb.copy_worksheet(template_ws)
            base_name = re.sub(r'[\\/*?:\[\]]', '-', str(cliente))[:31]
            final_name = base_name
            i = 1
            while final_name in wb.sheetnames:
                suffix = f"_{i}"
                final_name = base_name[:31-len(suffix)] + suffix
                i += 1
            new_ws.title = final_name

            # escreve cabeçalhos do modelo (preservar) e adiciona colunas extras do df à direita
            for col_idx, h in enumerate(model_headers, start=1):
                if h:
                    new_ws.cell(row=model_header_row, column=col_idx, value=h)
            next_free_col = new_ws.max_column + 1
            for extra in extra_dfcols:
                new_ws.cell(row=model_header_row, column=next_free_col, value=extra)
                model_idx_to_dfcol[next_free_col] = extra
                next_free_col += 1

            # define linha inicial para preencher dados
            start_row = model_header_row + 1
            for r_offset, (_, row_series) in enumerate(dados_cliente.iterrows(), start=start_row):
                for col_idx, dfcol in model_idx_to_dfcol.items():
                    val = row_series.get(dfcol, None)
                    new_ws.cell(row=r_offset, column=col_idx, value=val)

            created.append(final_name)

        # opcional: remover a aba modelo
        if not keep_model_tab:
            wb.remove(template_ws)

        # salvar em memória e disponibilizar download
        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)

        st.success(f"Planilha gerada — {len(created)} abas criadas.")
        st.download_button("⤓ Baixar planilha final", data=bio,
                           file_name="Planilha_Final_Clientes.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

