from __future__ import annotations
from dataclasses import dataclass
from typing import List, Optional
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# MODELO DE LINHA
@dataclass
class SheetRow:
    sheet_name: str
    row_index: int
    scenario: str
    tcode: str
    explanation: str
    parameter: str
    mode: str

# UTILITÁRIOS
def _norm(value: str) -> str:
    return (value or "").strip().lower()


def _find_col(header_row, *names: str) -> Optional[int]:
    wanted = {_norm(n) for n in names if n}
    for idx, cell in enumerate(header_row, start=1):
        if _norm(cell.value) in wanted:
            return idx
    return None

# LEITURA
def read_rows(xlsx_path: str, sheet_name: str) -> List[SheetRow]:
    wb = load_workbook(xlsx_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Aba '{sheet_name}' não existe.")

    ws = wb[sheet_name]
    header = list(ws.iter_rows(min_row=1, max_row=1, values_only=False))[0]

    col_scenario = _find_col(header, "Scenario", "Cenario", "Cenário")
    col_scen_expl = _find_col(header, "Scenario Explanation")
    col_tcode = _find_col(header, "Transação", "Transacao", "TCODE", "Transaction")
    col_test_expl = _find_col(header, "Test Explanation", "Explanation", "Test", "Descricao")
    col_param = _find_col(header, "Parâmetro", "Parametro", "Parameters", "Parameter")
    col_mode = _find_col(header, "Modo", "Mode")

    rows: List[SheetRow] = []
    for r in range(2, ws.max_row + 1):
        tcode = (ws.cell(r, col_tcode).value or "").strip() if col_tcode else ""
        if not tcode:
            continue

        scenario = (ws.cell(r, col_scenario).value or "").strip() if col_scenario else ""
        scen_expl = (ws.cell(r, col_scen_expl).value or "").strip() if col_scen_expl else ""
        test_expl = (ws.cell(r, col_test_expl).value or "").strip() if col_test_expl else ""

        textos = [txt for txt in [scen_expl, test_expl] if txt]
        explanation = " - ".join(textos)

        parameter = (ws.cell(r, col_param).value or "").strip() if col_param else ""
        mode = (ws.cell(r, col_mode).value or "real").strip() if col_mode else "real"

        rows.append(SheetRow(sheet_name, r, scenario, tcode, explanation, parameter, mode))

    return rows


def list_sheet_names(xlsx_path: str) -> List[str]:
    return list(load_workbook(xlsx_path).sheetnames)

# COLUNAS DE STATUS
def ensure_status_columns(xlsx_path: str, sheet_name: str) -> dict:
    wb = load_workbook(xlsx_path)
    ws = wb[sheet_name]
    header = list(ws.iter_rows(min_row=1, max_row=1, values_only=False))[0]

    col_status = _find_col(header, "Status")
    if not col_status:
        raise ValueError("Coluna 'Status' não encontrada.")

    col_source = col_status + 1
    col_message = col_status + 2
    col_suggestion = col_status + 3
    col_fix_conf = col_status + 4
    col_fix_just = col_status + 5
    col_evidence = col_status + 6

    headers = {
        col_source: "Status Source",
        col_message: "Status Message",
        col_suggestion: "Suggested Fix",
        col_fix_conf: "Fix Confidence",
        col_fix_just: "Fix Justification",
        col_evidence: "Evidence Path"
    }

    for col, title in headers.items():
        if not (ws.cell(1, col).value or "").strip():
            ws.cell(1, col).value = title

    wb.save(xlsx_path)

    return {
        "status": col_status,
        "source": col_source,
        "message": col_message,
        "evidence": col_evidence,
        "suggestion": col_suggestion,
        "fix_confidence": col_fix_conf,
        "fix_justification": col_fix_just,
    }

# ESCRITA (LEGADO)
def write_status_triplet(
    xlsx_path: str,
    sheet_name: str,
    row_index: int,
    status: str,
    source: str,
    message: str,
    evidence_path: str | None = None,
    suggestion: str = "",
) -> None:
    cols = ensure_status_columns(xlsx_path, sheet_name)
    wb = load_workbook(xlsx_path)
    ws = wb[sheet_name]

    ws.cell(row_index, cols["status"]).value = status
    ws.cell(row_index, cols["source"]).value = source
    ws.cell(row_index, cols["message"]).value = message

    ws.cell(row_index, cols["message"]).alignment = Alignment(wrap_text=True)
    ws.cell(row_index, cols["source"]).alignment = Alignment(wrap_text=True)

    ws.cell(row_index, cols["suggestion"]).value = suggestion
    ws.cell(row_index, cols["suggestion"]).alignment = Alignment(wrap_text=True)

    if evidence_path:
        ev_cell = ws.cell(row_index, cols["evidence"])
        ev_cell.value = evidence_path
        ev_cell.hyperlink = evidence_path
        ev_cell.style = "Hyperlink"

    wb.save(xlsx_path)

# ESCRITA (NOVA — CORRETA)
def write_status_with_fix_details(
    xlsx_path: str,
    sheet_name: str,
    row_index: int,
    status: str,
    source: str,
    message: str,
    suggested_fix: str,
    fix_confidence: int,
    fix_justification: str,
    evidence_path: str | None = None,
) -> None:
    cols = ensure_status_columns(xlsx_path, sheet_name)
    wb = load_workbook(xlsx_path)
    ws = wb[sheet_name]

    ws.cell(row_index, cols["status"]).value = status
    ws.cell(row_index, cols["source"]).value = source
    ws.cell(row_index, cols["message"]).value = message

    ws.cell(row_index, cols["suggestion"]).value = suggested_fix
    ws.cell(row_index, cols["fix_confidence"]).value = fix_confidence
    ws.cell(row_index, cols["fix_justification"]).value = fix_justification

    for key in ("message", "suggestion", "fix_justification"):
        ws.cell(row_index, cols[key]).alignment = Alignment(wrap_text=True, vertical="top")

    if evidence_path:
        ev_cell = ws.cell(row_index, cols["evidence"])
        ev_cell.value = evidence_path
        ev_cell.hyperlink = evidence_path
        ev_cell.style = "Hyperlink"

    wb.save(xlsx_path)

# FORMATAÇÃO
def format_output_sheet(xlsx_path: str, sheet_name: str) -> None:
    cols = ensure_status_columns(xlsx_path, sheet_name)
    wb = load_workbook(xlsx_path)
    ws = wb[sheet_name]

    max_col = ws.max_column
    max_row = ws.max_row

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for c in range(1, max_col + 1):
        cell = ws.cell(1, c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

    pass_fill = PatternFill(fill_type="solid", fgColor="E2F0D9")
    fail_fill = PatternFill(fill_type="solid", fgColor="FCE4D6")

    for r in range(2, max_row + 1):
        st = (ws.cell(r, cols["status"]).value or "").strip().upper()
        for c in range(1, max_col + 1):
            ws.cell(r, c).border = thin_border

        if st == "PASS":
            ws.cell(r, cols["status"]).fill = pass_fill
        elif st == "FAIL":
            for key in ("status", "message", "suggestion"):
                ws.cell(r, cols[key]).fill = fail_fill

    preferred = {
        cols["status"]: 10,
        cols["source"]: 18,
        cols["message"]: 60,
        cols["suggestion"]: 45,
        cols["fix_confidence"]: 15,
        cols["fix_justification"]: 60,
        cols["evidence"]: 45,
    }

    for c in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = preferred.get(c, 15)

    wb.save(xlsx_path)