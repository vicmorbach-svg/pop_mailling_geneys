import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Configuração da página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Populador de Campanha",
    page_icon="📋",
    layout="wide",
)

# ── Colunas do template padrão ──────────────────────────────────────────────
TEMPLATE_COLUMNS = [
    "matricula", "cliente", "valor", "cidade", "inin-outbound-id",
    "Tel1", "Tel2", "Tel3", "Tel4", "Tel5", "Tel6",
    "extra 3", "Contrato", "codigo",
    "timezone", "priority", "concessionaria", "subconcessionaria", "dados",
]

# Colunas com valor fixo (não mapeáveis pelo usuário)
FIXED_COLUMNS = {
    "timezone": "America/Belem",
    "priority": 5,
    "concessionaria": "CORSAN",
}

# Colunas disponíveis para mapeamento (excluindo as fixas e a coluna 'dados' para tratamento especial)
# A coluna 'dados' será tratada separadamente para permitir múltiplos mapeamentos.
MAPPABLE_SINGLE_COLUMNS = [col for col in TEMPLATE_COLUMNS if col not in FIXED_COLUMNS and col != "dados"]


# ── Helpers ─────────────────────────────────────────────────────────────────
def load_file(uploaded_file) -> pd.DataFrame | None:
    """Carrega XLSX ou CSV em um DataFrame."""
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, dtype=str)
        elif name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded_file, dtype=str)
        else:
            st.error("Formato não suportado. Envie um arquivo .xlsx, .xls ou .csv.")
            return None
        return df
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        return None


def build_output(
    source_df: pd.DataFrame,
    single_mapping: dict[str, str],
    multi_mapping_dados: list[str]
) -> pd.DataFrame:
    """
    Constrói o DataFrame de saída com todas as colunas do template.
    - single_mapping: {coluna_template -> coluna_origem} para mapeamentos 1:1
    - multi_mapping_dados: [coluna_origem_1, coluna_origem_2, ...] para a coluna 'dados'
    """
    out = pd.DataFrame(index=range(len(source_df)), columns=TEMPLATE_COLUMNS)

    # Aplica valores fixos
    for col, val in FIXED_COLUMNS.items():
        out[col] = val

    # Aplica mapeamento do usuário para colunas 1:1
    for template_col, source_col in single_mapping.items():
        if source_col and source_col != "— não mapear —":
            out[template_col] = source_df[source_col].values

    # Aplica mapeamento para a coluna 'dados' (múltiplas colunas de origem)
    if "dados" in TEMPLATE_COLUMNS and multi_mapping_dados:
        # Filtra colunas que foram realmente selecionadas e existem no source_df
        selected_source_cols = [col for col in multi_mapping_dados if col in source_df.columns]

        if selected_source_cols:
            # Converte todas as colunas selecionadas para string e concatena com ';'
            # Preenche NaN com string vazia antes de concatenar para evitar 'nan;valor'
            out["dados"] = source_df[selected_source_cols].fillna('').astype(str).agg(';'.join, axis=1)
        else:
            out["dados"] = None # Ou uma string vazia, dependendo da sua preferência

    return out


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Converte o DataFrame para bytes de um XLSX estilizado."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Campanha"

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    fixed_fill = PatternFill("solid", fgColor="E2EFDA")   # verde claro p/ fixas
    mapped_fill = PatternFill("solid", fgColor="DDEEFF")  # azul claro p/ mapeadas
    empty_fill  = PatternFill("solid", fgColor="F2F2F2")  # cinza p/ vazias

    # Cabeçalho
    for c_idx, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = border
        ws.column_dimensions[get_column_letter(c_idx)].width = max(14, len(col) + 4)

    ws.row_dimensions[1].height = 30

    # Dados
    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, (col, val) in enumerate(zip(df.columns, row), start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val if pd.notna(val) else None)
            cell.border = border
            cell.alignment = Alignment(vertical="center")

            if col in FIXED_COLUMNS:
                cell.fill = fixed_fill
            elif val and str(val).strip():
                cell.fill = mapped_fill
            else:
                cell.fill = empty_fill

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── Interface ────────────────────────────────────────────────────────────────
st.title("📋 Populador de Campanha CORSAN")
st.caption("Faça upload de qualquer planilha e mapeie as colunas para o template padrão.")

st.divider()

# Legenda das colunas fixas
with st.expander("ℹ️ Colunas preenchidas automaticamente", expanded=False):
    cols_info = st.columns(3)
    for i, (col, val) in enumerate(FIXED_COLUMNS.items()):
        cols_info[i].metric(label=col, value=str(val))

st.divider()

# Upload
uploaded = st.file_uploader(
    "📂 Selecione o arquivo de origem (.xlsx, .xls ou .csv)",
    type=["xlsx", "xls", "csv"],
)

if uploaded:
    source_df = load_file(uploaded)

    if source_df is not None:
        st.success(f"Arquivo carregado: **{uploaded.name}** — {len(source_df)} linhas × {len(source_df.columns)} colunas")

        with st.expander("👁️ Prévia do arquivo de origem (primeiras 5 linhas)", expanded=False):
            st.dataframe(source_df.head(5), use_container_width=True)

        st.divider()
        st.subheader("🔗 Mapeamento de colunas")
        st.caption(
            "Para cada coluna do **template padrão** (esquerda), escolha a coluna "
            "correspondente no **arquivo enviado** (direita). Selecione '— não mapear —' "
            "para deixar a coluna em branco. Para a coluna 'dados', você pode selecionar múltiplas colunas."
        )

        source_options = ["— não mapear —"] + list(source_df.columns)
        source_options_for_multi = list(source_df.columns) # Para multiselect, não precisa de "não mapear"

        # Tenta sugerir mapeamento automático por nome exato (case-insensitive)
        source_lower = {c.lower(): c for c in source_df.columns}

        single_mapping: dict[str, str] = {}
        multi_mapping_dados: list[str] = []

        # Divide em duas colunas de selects para os mapeamentos 1:1
        left_cols = MAPPABLE_SINGLE_COLUMNS[: len(MAPPABLE_SINGLE_COLUMNS) // 2 + len(MAPPABLE_SINGLE_COLUMNS) % 2]
        right_cols = MAPPABLE_SINGLE_COLUMNS[len(MAPPABLE_SINGLE_COLUMNS) // 2 + len(MAPPABLE_SINGLE_COLUMNS) % 2 :]

        col_a, col_b = st.columns(2)

        for col in left_cols:
            suggestion = source_lower.get(col.lower(), "— não mapear —")
            single_mapping[col] = col_a.selectbox(
                f"template: **{col}**",
                options=source_options,
                index=source_options.index(suggestion) if suggestion in source_options else 0,
                key=f"map_{col}",
            )

        for col in right_cols:
            suggestion = source_lower.get(col.lower(), "— não mapear —")
            single_mapping[col] = col_b.selectbox(
                f"template: **{col}**",
                options=source_options,
                index=source_options.index(suggestion) if suggestion in source_options else 0,
                key=f"map_{col}",
            )

        st.markdown("---") # Separador visual para a coluna 'dados'
        st.subheader("Configuração da coluna 'dados'")
        st.caption("Selecione as colunas do arquivo de origem que deseja concatenar na coluna 'dados' do template, separadas por ';'.")
        multi_mapping_dados = st.multiselect(
            "template: **dados**",
            options=source_options_for_multi,
            key="map_dados_multi",
            help="Selecione uma ou mais colunas do arquivo de origem. Os valores serão unidos por ';'."
        )


        st.divider()

        # Botão para pré-visualizar e aplicar a lógica de duplicidade
        if st.button("🔍 Pré-visualizar resultado e remover duplicidades", use_container_width=True):
            output_df = build_output(source_df, single_mapping, multi_mapping_dados)

            # --- Lógica para remover duplicidades em 'Tel1' ---
            if "Tel1" in output_df.columns:
                initial_rows = len(output_df)
                # Garante que 'Tel1' é string e remove espaços em branco antes de verificar duplicidade
                output_df['Tel1'] = output_df['Tel1'].astype(str).str.strip()
                output_df_cleaned = output_df.drop_duplicates(subset=['Tel1'], keep='first')
                removed_duplicates = initial_rows - len(output_df_cleaned)

                if removed_duplicates > 0:
                    st.warning(f"⚠️ Foram removidas **{removed_duplicates}** linhas duplicadas com base na coluna 'Tel1'.")
                    output_df = output_df_cleaned
                else:
                    st.info("✅ Nenhuma duplicidade encontrada na coluna 'Tel1'.")
            else:
                st.warning("Coluna 'Tel1' não encontrada no template de saída para verificação de duplicidades.")
            # --- Fim da lógica de duplicidade ---

            st.subheader("📊 Prévia do arquivo de saída (primeiras 10 linhas)")
            st.dataframe(output_df.head(10), use_container_width=True)
            st.session_state["output_df"] = output_df # Salva o DF já limpo

        # Gerar e baixar
        st.divider()
        if st.button("⚙️ Gerar arquivo XLSX", type="primary", use_container_width=True):
            # Se o usuário já pré-visualizou, usamos o DF da session_state
            # Caso contrário, construímos e limpamos novamente
            if "output_df" in st.session_state:
                output_df = st.session_state["output_df"]
            else:
                output_df = build_output(source_df, single_mapping, multi_mapping_dados)
                # --- Lógica para remover duplicidades em 'Tel1' (repetida para garantir se não houve pré-visualização) ---
                if "Tel1" in output_df.columns:
                    initial_rows = len(output_df)
                    output_df['Tel1'] = output_df['Tel1'].astype(str).str.strip()
                    output_df_cleaned = output_df.drop_duplicates(subset=['Tel1'], keep='first')
                    removed_duplicates = initial_rows - len(output_df_cleaned)
                    if removed_duplicates > 0:
                        st.warning(f"⚠️ Foram removidas **{removed_duplicates}** linhas duplicadas com base na coluna 'Tel1' durante a geração do arquivo.")
                        output_df = output_df_cleaned
                    else:
                        st.info("✅ Nenhuma duplicidade encontrada na coluna 'Tel1' durante a geração do arquivo.")
                else:
                    st.warning("Coluna 'Tel1' não encontrada no template de saída para verificação de duplicidades durante a geração do arquivo.")
                # --- Fim da lógica de duplicidade ---
                st.session_state["output_df"] = output_df # Salva o DF limpo na session_state

            xlsx_bytes = to_excel_bytes(output_df)
            st.session_state["xlsx_bytes"] = xlsx_bytes
            st.success(f"Arquivo gerado com sucesso! {len(output_df)} linhas prontas para download.")

        if "xlsx_bytes" in st.session_state:
            st.download_button(
                label="⬇️ Baixar TEMPLATE_CAMPANHA.xlsx",
                data=st.session_state["xlsx_bytes"],
                file_name="TEMPLATE_CAMPANHA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
