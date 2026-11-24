import io
import os
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Variáveis de posição das colunas na planilha de saída
COL_DOC1_IDX = 1
COL_VAL1_IDX = 2
COL_DOC2_IDX = 4
COL_VAL2_IDX = 5
COL_DIFF_IDX = 7

COL_DOC1_LET = get_column_letter(COL_DOC1_IDX)
COL_VAL1_LET = get_column_letter(COL_VAL1_IDX)
COL_DOC2_LET = get_column_letter(COL_DOC2_IDX)
COL_VAL2_LET = get_column_letter(COL_VAL2_IDX)
COL_DIFF_LET = get_column_letter(COL_DIFF_IDX)


def to_dec(v) -> Decimal:
    """Converte qualquer valor (numérico ou texto formatado) para Decimal(0.01) sem usar regex."""
    if v is None:
        return Decimal("0.00")
    if isinstance(v, float) and pd.isna(v):
        return Decimal("0.00")

    s_raw = str(v).strip()
    if not s_raw:
        return Decimal("0.00")

    if isinstance(v, (int, float, Decimal)):
        return Decimal(str(v)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    s = s_raw.replace("R$", "").replace("\xa0", "").replace(" ", "")
    negativo_parenteses = False
    if s.startswith("(") and s.endswith(")"):
        negativo_parenteses = True
        s = s[1:-1].strip()

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")

    # Mantém apenas caracteres numéricos válidos para Decimal
    permitidos = set("0123456789.-+")
    s_limpo = "".join(ch for ch in s if ch in permitidos)
    if not s_limpo or s_limpo in {"-", "+", "."}:
        return Decimal("0.00")

    try:
        val = Decimal(s_limpo).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return -val if negativo_parenteses else val
    except Exception:
        return Decimal("0.00")


def norm_doc_ate_7(v) -> str:
    """Extrai e normaliza o número do documento (até 7 dígitos, sem zeros à esquerda) sem regex."""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""

    s = str(v).strip()
    if not s or s.lower() in {"nan", "none", "nat"}:
        return ""

    if s.endswith(".0"):
        s = s[:-2]

    marcadores = ("NR:", "BORD", "MUTUO", "IMPL", "BXA", "ESTT")
    s_upper = s.upper()
    if "/" in s and not any(m in s_upper for m in marcadores):
        s = s.split("/")[0].strip()

    if len(s) <= 7 and s.isdigit():
        limpo = s.lstrip("0")
        return limpo if limpo else "0"

    pos_bord = s_upper.find("BORD")
    if pos_bord != -1:
        s = s[:pos_bord]

    # Separa por delimitadores comuns e busca blocos numéricos de 1 a 7 dígitos
    texto_sep = s
    for sep in ("-", ":", " ", ";", "|", ","):
        texto_sep = texto_sep.replace(sep, "/")

    candidatos = []
    for parte in texto_sep.split("/"):
        bloco_num = "".join(ch for ch in parte.strip() if ch.isdigit())
        if 1 <= len(bloco_num) <= 7:
            candidatos.append(bloco_num)

    # Prioriza o primeiro bloco com dígitos não-zero (evita pegar apenas "000")
    for cand in candidatos:
        limpo = cand.lstrip("0")
        if limpo:
            return limpo

    if candidatos:
        return "0"

    return ""


def detectar_linha_cabecalho(df_raw: pd.DataFrame) -> int:
    """Detecta automaticamente a linha com mais células textuais de cabeçalho."""
    melhor_idx = 0
    melhor_pontuacao = -1
    palavras_chave = {
        "histórico",
        "historico",
        "título",
        "titulo",
        "valor",
        "saldo",
        "débito",
        "debito",
        "crédito",
        "credito",
        "documento",
    }

    limite = min(15, len(df_raw))
    for idx in range(limite):
        linha = df_raw.iloc[idx].tolist()
        nao_vazios = [str(x).strip() for x in linha if pd.notna(x) and str(x).strip() != ""]
        if not nao_vazios:
            continue
        score = len(nao_vazios)
        for item in nao_vazios:
            if item.lower() in palavras_chave:
                score += 5
        if score > melhor_pontuacao:
            melhor_pontuacao = score
            melhor_idx = idx
    return melhor_idx


def ler_planilha_simples(file_bytes: bytes) -> pd.DataFrame:
    """Lê a primeira aba e encontra o cabeçalho automaticamente."""
    excel_io = io.BytesIO(file_bytes)
    primeira_aba = pd.ExcelFile(excel_io, engine="openpyxl").sheet_names[0]
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=primeira_aba, header=None, engine="openpyxl")
    if df_raw.empty:
        return pd.DataFrame()

    cab_idx = detectar_linha_cabecalho(df_raw)
    raw_headers = df_raw.iloc[cab_idx].tolist()

    seen = {}
    cols = []
    for i, c in enumerate(raw_headers):
        nome = str(c).strip() if pd.notna(c) and str(c).strip() != "" else f"Coluna_{i + 1}"
        if nome in seen:
            seen[nome] += 1
            nome = f"{nome}_{seen[nome]}"
        else:
            seen[nome] = 0
        cols.append(nome)

    df = df_raw.iloc[cab_idx + 1 :].copy()
    df.columns = cols
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def consolidar_coluna_selecionada(df: pd.DataFrame, col_doc: str, col_val: str) -> dict[str, Decimal]:
    """Agrupa e soma os valores tratados da coluna selecionada."""
    totais: dict[str, Decimal] = {}
    if df.empty or col_doc not in df.columns or col_val not in df.columns:
        return totais

    for _, row in df.iterrows():
        raw_doc = row[col_doc]
        if raw_doc is None or (isinstance(raw_doc, float) and pd.isna(raw_doc)):
            continue
        texto_doc = str(raw_doc).strip()
        if not texto_doc or texto_doc.upper().startswith("SUBTOTAL"):
            continue

        doc = norm_doc_ate_7(raw_doc)
        if not doc:
            continue

        val_tratado = to_dec(row[col_val])
        totais[doc] = totais.get(doc, Decimal("0.00")) + val_tratado

    return totais


def gerar_excel_bytes(soma_arq1: dict[str, Decimal], soma_arq2: dict[str, Decimal]):
    ambos_docs = sorted(set(soma_arq1.keys()).intersection(set(soma_arq2.keys())))
    nao_encontrados_docs = sorted(set(soma_arq1.keys()).symmetric_difference(set(soma_arq2.keys())))

    rows_data = []
    for doc in ambos_docs:
        v1 = soma_arq1[doc]
        v2 = soma_arq2[doc]
        a1, a2 = abs(v1), abs(v2)
        diff = Decimal("0.00") if a1 == a2 else (a1 - a2)
        rows_data.append({
            "doc1": doc,
            "val1": float(v1),
            "doc2": doc,
            "val2": float(v2),
            "diff": float(diff),
            "equal_mag": (a1 == a2),
        })

    for doc in nao_encontrados_docs:
        in_1 = doc in soma_arq1
        in_2 = doc in soma_arq2
        v1 = soma_arq1.get(doc, Decimal("0.00"))
        v2 = soma_arq2.get(doc, Decimal("0.00"))
        a1 = abs(v1) if in_1 else Decimal("0.00")
        a2 = abs(v2) if in_2 else Decimal("0.00")
        diff = Decimal("0.00") if (in_1 and in_2 and a1 == a2) else (a1 - a2)
        rows_data.append({
            "doc1": doc if in_1 else "—",
            "val1": float(v1) if in_1 else 0.0,
            "doc2": doc if in_2 else "—",
            "val2": float(v2) if in_2 else 0.0,
            "diff": float(diff),
            "equal_mag": (in_1 and in_2 and a1 == a2),
        })

    wb = Workbook()
    ws = wb.active
    ws.title = "Comparativo_Ordenado"

    arq1_banner_fill = PatternFill(start_color="1F4E5F", end_color="1F4E5F", fill_type="solid")
    arq2_banner_fill = PatternFill(start_color="3B4252", end_color="3B4252", fill_type="solid")
    diff_banner_fill = PatternFill(start_color="4C566A", end_color="4C566A", fill_type="solid")

    subhdr_arq1_fill = PatternFill(start_color="2E6B82", end_color="2E6B82", fill_type="solid")
    subhdr_arq2_fill = PatternFill(start_color="4C566A", end_color="4C566A", fill_type="solid")
    subhdr_diff_fill = PatternFill(start_color="616E88", end_color="616E88", fill_type="solid")

    section_unmatched_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    section_unmatched_font = Font(name="Calibri", size=10, bold=True, color="92400E")

    banner_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    subhdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11, bold=False, color="111827")
    bold_font = Font(name="Calibri", size=11, bold=True, color="111827")
    muted_font = Font(name="Calibri", size=11, bold=False, color="6B7280")
    zero_diff_font = Font(name="Calibri", size=11, bold=True, color="166534")
    nonzero_diff_font = Font(name="Calibri", size=11, bold=True, color="991B1B")

    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    match_fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    total_fill = PatternFill(start_color="E5E7EB", end_color="E5E7EB", fill_type="solid")

    thin = Side(style="thin", color="CBD5E1")
    thin_top = Side(style="thin", color="111827")
    double_bottom = Side(style="double", color="111827")

    grid_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    total_border = Border(left=thin, right=thin, top=thin_top, bottom=double_bottom)

    ws.merge_cells(f"{COL_DOC1_LET}1:{COL_VAL1_LET}1")
    ws[f"{COL_DOC1_LET}1"] = "arquivo 1"
    ws[f"{COL_DOC1_LET}1"].fill = arq1_banner_fill
    ws[f"{COL_DOC1_LET}1"].font = banner_font
    ws[f"{COL_DOC1_LET}1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"{COL_DOC2_LET}1:{COL_VAL2_LET}1")
    ws[f"{COL_DOC2_LET}1"] = "arquivo 2"
    ws[f"{COL_DOC2_LET}1"].fill = arq2_banner_fill
    ws[f"{COL_DOC2_LET}1"].font = banner_font
    ws[f"{COL_DOC2_LET}1"].alignment = Alignment(horizontal="center", vertical="center")

    ws[f"{COL_DIFF_LET}1"] = "diferença"
    ws[f"{COL_DIFF_LET}1"].fill = diff_banner_fill
    ws[f"{COL_DIFF_LET}1"].font = banner_font
    ws[f"{COL_DIFF_LET}1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    sub_headers = {
        COL_DOC1_IDX: ("documento", subhdr_arq1_fill, "center"),
        COL_VAL1_IDX: ("valor", subhdr_arq1_fill, "right"),
        COL_DOC2_IDX: ("documento", subhdr_arq2_fill, "center"),
        COL_VAL2_IDX: ("valor", subhdr_arq2_fill, "right"),
        COL_DIFF_IDX: ("diferença", subhdr_diff_fill, "right"),
    }
    for col_i, (text, fill, align) in sub_headers.items():
        c = ws.cell(2, col_i, text)
        c.fill = fill
        c.font = subhdr_font
        c.alignment = Alignment(horizontal=align, vertical="center")
        c.border = grid_border
    ws.row_dimensions[2].height = 22

    current_row = 3
    data_row_indices = []

    for idx, item in enumerate(rows_data):
        if idx == len(ambos_docs) and len(nao_encontrados_docs) > 0:
            ws.merge_cells(
                start_row=current_row,
                start_column=COL_DOC1_IDX,
                end_row=current_row,
                end_column=COL_DIFF_IDX,
            )
            sep_cell = ws.cell(current_row, COL_DOC1_IDX, "Documentos sem correspondência no outro arquivo")
            sep_cell.fill = section_unmatched_fill
            sep_cell.font = section_unmatched_font
            sep_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[current_row].height = 20
            current_row += 1

        r = current_row
        data_row_indices.append(r)
        ws.row_dimensions[r].height = 20

        ref_v1 = f"{COL_VAL1_LET}{r}"
        ref_v2 = f"{COL_VAL2_LET}{r}"

        c_d1 = ws.cell(r, COL_DOC1_IDX, item["doc1"])
        c_v1 = ws.cell(r, COL_VAL1_IDX, item["val1"])
        c_d2 = ws.cell(r, COL_DOC2_IDX, item["doc2"])
        c_v2 = ws.cell(r, COL_VAL2_IDX, item["val2"])
        c_df = ws.cell(r, COL_DIFF_IDX, f"=IF(ABS({ref_v1})=ABS({ref_v2}),0,ABS({ref_v1})-{ref_v2})")

        row_bg = match_fill if item["equal_mag"] else (zebra_fill if idx % 2 == 1 else white_fill)
        for cell in (c_d1, c_v1, c_d2, c_v2, c_df):
            cell.fill = row_bg
            cell.border = grid_border

        c_d1.font = bold_font if item["doc1"] != "—" else muted_font
        c_d1.alignment = Alignment(horizontal="center", vertical="center")

        c_v1.font = data_font
        c_v1.number_format = "#,##0.00;-#,##0.00;0.00"
        c_v1.alignment = Alignment(horizontal="right", vertical="center")

        c_d2.font = bold_font if item["doc2"] != "—" else muted_font
        c_d2.alignment = Alignment(horizontal="center", vertical="center")

        c_v2.font = data_font
        c_v2.number_format = "#,##0.00;-#,##0.00;0.00"
        c_v2.alignment = Alignment(horizontal="right", vertical="center")

        is_zero = abs(item["diff"]) < 1e-9
        c_df.font = zero_diff_font if is_zero else nonzero_diff_font
        c_df.number_format = "#,##0.00;-#,##0.00;0.00"
        c_df.alignment = Alignment(horizontal="right", vertical="center")

        current_row += 1

    if data_row_indices:
        last_data_r = data_row_indices[-1]
        tot_row = current_row
        ws.row_dimensions[tot_row].height = 24

        ws.cell(tot_row, COL_DOC1_IDX, "TOTAL").font = bold_font
        ws.cell(tot_row, COL_DOC1_IDX).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(tot_row, COL_VAL1_IDX, "—").font = muted_font
        ws.cell(tot_row, COL_VAL1_IDX).alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(tot_row, COL_DOC2_IDX, "TOTAL").font = bold_font
        ws.cell(tot_row, COL_DOC2_IDX).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(tot_row, COL_VAL2_IDX, "—").font = muted_font
        ws.cell(tot_row, COL_VAL2_IDX).alignment = Alignment(horizontal="center", vertical="center")

        sum_diff_expr = "+".join([f"{COL_DIFF_LET}{rr}" for rr in data_row_indices])
        c_tot_diff = ws.cell(tot_row, COL_DIFF_IDX, f"={sum_diff_expr}")
        c_tot_diff.font = bold_font
        c_tot_diff.number_format = "#,##0.00;-#,##0.00;0.00"
        c_tot_diff.alignment = Alignment(horizontal="right", vertical="center")

        for col_i in (COL_DOC1_IDX, COL_VAL1_IDX, COL_DOC2_IDX, COL_VAL2_IDX, COL_DIFF_IDX):
            cell = ws.cell(tot_row, col_i)
            cell.fill = total_fill
            cell.border = total_border

        ws.freeze_panes = f"{COL_DOC1_LET}3"
        ws.auto_filter.ref = f"{COL_DOC1_LET}2:{COL_DIFF_LET}{last_data_r}"

        ws_res = wb.create_sheet(title="Resumo_e_Grafico")
        ws_res["A1"] = "Indicador da Conciliação"
        ws_res["B1"] = "Valor"
        ws_res["A1"].font = subhdr_font
        ws_res["B1"].font = subhdr_font
        ws_res["A1"].fill = arq1_banner_fill
        ws_res["B1"].fill = arq1_banner_fill

        zero_count = sum(1 for x in rows_data if x["equal_mag"])
        nonzero_count = sum(1 for x in rows_data if not x["equal_mag"])

        resumo_rows = [
            ("Documentos com Valores Iguais (Diferença = 0)", zero_count),
            ("Documentos com Valores Diferentes / Pendentes", nonzero_count),
            ("Soma das Diferenças", f"={ws.title}!{COL_DIFF_LET}{tot_row}"),
        ]

        for idx, (label, val) in enumerate(resumo_rows, start=2):
            ca = ws_res.cell(idx, 1, label)
            cb = ws_res.cell(idx, 2, val)
            ca.font = bold_font
            cb.font = bold_font
            ca.border = grid_border
            cb.border = grid_border
            cb.number_format = "#,##0.00;-#,##0.00;0.00" if idx == 4 else "#,##0"
            cb.alignment = Alignment(horizontal="right", vertical="center")

        ws_res.column_dimensions["A"].width = 44
        ws_res.column_dimensions["B"].width = 22

        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Quantidade de Documentos: Valores Iguais vs Diferentes"
        chart.y_axis.title = "Quantidade"
        chart.x_axis.title = "Classificação"
        data_ref = Reference(ws_res, min_col=2, min_row=1, max_row=3)
        cats_ref = Reference(ws_res, min_col=1, min_row=2, max_row=3)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.legend = None
        chart.width = 16
        chart.height = 10
        ws_res.add_chart(chart, "D2")

    col_widths = {
        COL_DOC1_IDX: 18,
        COL_VAL1_IDX: 20,
        3: 4,
        COL_DOC2_IDX: 18,
        COL_VAL2_IDX: 20,
        6: 4,
        COL_DIFF_IDX: 20,
    }
    for col_i, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_i)].width = width

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue(), rows_data


def listar_sem_correspondencia_com_valor(
    soma_arq1: dict[str, Decimal], soma_arq2: dict[str, Decimal]
) -> list[tuple[str, Decimal]]:
    """Retorna apenas (documento, valor) de documentos sem correspondência no outro arquivo e com valor != 0."""
    itens: list[tuple[str, Decimal]] = []
    for doc in sorted(set(soma_arq1.keys()) - set(soma_arq2.keys())):
        v1 = soma_arq1[doc]
        if v1 != Decimal("0.00"):
            itens.append((doc, v1))
    for doc in sorted(set(soma_arq2.keys()) - set(soma_arq1.keys())):
        v2 = soma_arq2[doc]
        if v2 != Decimal("0.00"):
            itens.append((doc, v2))
    return itens


porta_env = os.environ.get("PORT", "8501")
st.set_page_config(page_title=f"Conciliação por Documento e Valor ({porta_env})", layout="wide")
st.title("Conciliação de Planilhas")

col_left, col_right = st.columns(2)

with col_left:
    st.subheader("Arquivo 1")
    arq1_file = st.file_uploader("Selecione o Arquivo 1 (.xlsx)", type=["xlsx"], key="arq1")
    doc1_col = None
    val1_col = None
    df1 = pd.DataFrame()

    if arq1_file is not None:
        df1 = ler_planilha_simples(arq1_file.getvalue())
        if not df1.empty:
            cols1 = list(df1.columns)
            doc1_col = st.selectbox("Documento", cols1, key="doc1_col")
            val1_col = st.selectbox("Valor", cols1, key="val1_col")

with col_right:
    st.subheader("Arquivo 2")
    arq2_file = st.file_uploader("Selecione o Arquivo 2 (.xlsx)", type=["xlsx"], key="arq2")
    doc2_col = None
    val2_col = None
    df2 = pd.DataFrame()

    if arq2_file is not None:
        df2 = ler_planilha_simples(arq2_file.getvalue())
        if not df2.empty:
            cols2 = list(df2.columns)
            doc2_col = st.selectbox("Documento", cols2, key="doc2_col")
            val2_col = st.selectbox("Valor", cols2, key="val2_col")

if arq1_file is not None and arq2_file is not None and doc1_col and val1_col and doc2_col and val2_col:
    soma1 = consolidar_coluna_selecionada(df1, doc1_col, val1_col)
    soma2 = consolidar_coluna_selecionada(df2, doc2_col, val2_col)
    excel_bytes, rows_preview = gerar_excel_bytes(soma1, soma2)

    df_view = pd.DataFrame([
        {
            f"arquivo 1 ({doc1_col})": r["doc1"],
            f"valor 1 ({val1_col})": f"{r['val1']:,.2f}",
            f"arquivo 2 ({doc2_col})": r["doc2"],
            f"valor 2 ({val2_col})": f"{r['val2']:,.2f}",
            "diferença": f"{r['diff']:,.2f}",
        }
        for r in rows_preview
    ])

    sem_correspondencia = listar_sem_correspondencia_com_valor(soma1, soma2)

    st.divider()
    st.dataframe(df_view, use_container_width=True, hide_index=True)

    if sem_correspondencia:
        linhas_msg = [f"**Sem correspondência no outro arquivo ({len(sem_correspondencia)}):**"]
        for doc, valor in sem_correspondencia:
            linhas_msg.append(f"- `{doc}` | `{valor:,.2f}`")
        st.warning("\n\n".join(linhas_msg))

    st.download_button(
        label="Baixar Excel Comparativo",
        data=excel_bytes,
        file_name="comparativo_conciliacao.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )