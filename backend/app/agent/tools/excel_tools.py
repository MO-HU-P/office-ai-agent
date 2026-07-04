"""Excel (.xlsx) 操作ツール群"""
from typing import Any, Optional

from langchain_core.tools import tool
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, range_boundaries

from ...atomic import atomic_save
from ...config import resolve_workspace_path


def _open(filename: str):
    path = resolve_workspace_path(filename, must_exist=True)
    if path.suffix != ".xlsx":
        raise ValueError("Excelブックは .xlsx ファイルを指定してください")
    return load_workbook(str(path)), str(path)


def _sheet(wb, sheet: Optional[str]):
    if sheet is None:
        return wb.active
    if sheet in wb.sheetnames:
        return wb[sheet]
    return wb.create_sheet(sheet)


def _set_value(ws, ref: str, value: Any):
    cell = ws[ref]
    if isinstance(value, str) and value.startswith("="):
        cell.value = value  # 数式
    else:
        cell.value = value


@tool
def excel_create(filename: str, sheet_name: str = "Sheet1") -> str:
    """新しいExcelブック(.xlsx)を作成する。filenameは必ず.xlsxで終わること。"""
    path = resolve_workspace_path(filename)
    if path.suffix != ".xlsx":
        return "エラー: filenameは .xlsx で終わる必要があります"
    wb = Workbook()
    wb.active.title = sheet_name
    atomic_save(wb.save, path)
    return f"{filename} を作成しました (シート: {sheet_name})"


@tool
def excel_read(filename: str, sheet: str = "", cell_range: str = "", mode: str = "full") -> str:
    """Excelブックの内容を読む。sheet省略時は全シート名一覧+アクティブシートの内容を返す。
    cell_rangeで範囲指定可能(例: "A1:D10")。数式セルは数式のまま表示される。
    mode="summary"にすると各シートの名前と使用範囲だけを返す。大きなブックはまずsummaryで
    全体を把握し、必要なシート・範囲だけを読むこと。"""
    wb, _ = _open(filename)
    if mode == "summary":
        lines = [f"{ws.title}: {ws.max_row}行 x {ws.max_column}列" for ws in wb.worksheets]
        return "シート一覧:\n" + "\n".join(lines)
    lines = [f"シート一覧: {', '.join(wb.sheetnames)}"]
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.active
    lines.append(f"--- シート「{ws.title}」({ws.max_row}行 x {ws.max_column}列) ---")
    if cell_range:
        min_c, min_r, max_c, max_r = range_boundaries(cell_range)
    else:
        min_r, min_c = 1, 1
        max_r, max_c = min(ws.max_row, 100), min(ws.max_column, 30)
    for r in range(min_r, max_r + 1):
        cells = []
        for c in range(min_c, max_c + 1):
            v = ws.cell(row=r, column=c).value
            cells.append("" if v is None else str(v))
        if any(cells):
            lines.append(f"{r}: " + " | ".join(cells))
    return "\n".join(lines)


@tool
def excel_write_cells(filename: str, cells: dict[str, Any], sheet: str = "") -> str:
    """Excelのセルに値を書き込む。cellsはセル番地→値の辞書。例: {"A1": "商品名", "B1": 100, "C1": "=A1*B1"}
    "=" で始まる文字列は数式として書き込まれる。sheetが存在しない場合は新規作成される。"""
    wb, path = _open(filename)
    ws = _sheet(wb, sheet or None)
    for ref, value in cells.items():
        _set_value(ws, ref, value)
    atomic_save(wb.save, path)
    return f"{filename} のシート「{ws.title}」に {len(cells)} セルを書き込みました"


@tool
def excel_write_rows(filename: str, start_cell: str, rows: list[list[Any]], sheet: str = "") -> str:
    """Excelに複数行のデータを一括で書き込む。start_cell(例: "A2")を左上として、rows(2次元配列)を展開する。
    表データを書くときはexcel_write_cellsよりこちらを使うこと。"=" で始まる値は数式。"""
    wb, path = _open(filename)
    ws = _sheet(wb, sheet or None)
    min_c, min_r, _, _ = range_boundaries(start_cell)
    n = 0
    for ri, row in enumerate(rows):
        for ci, value in enumerate(row):
            ref = f"{get_column_letter(min_c + ci)}{min_r + ri}"
            _set_value(ws, ref, value)
            n += 1
    atomic_save(wb.save, path)
    return f"{filename} のシート「{ws.title}」に {len(rows)}行 ({n}セル) を書き込みました"


@tool
def excel_format(
    filename: str,
    cell_range: str,
    sheet: str = "",
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_size: Optional[float] = None,
    font_color: str = "",
    bg_color: str = "",
    align: str = "",
    number_format: str = "",
    col_width: Optional[float] = None,
) -> str:
    """Excelのセル範囲に書式を設定する。cell_rangeは "A1:D1" のような範囲か単一セル。
    "A1:D1,A5:D5" のようにカンマ区切りで複数範囲へ同じ書式を一括設定できる(1範囲ずつ繰り返さないこと)。
    font_color/bg_colorは "#RRGGBB" 形式。alignは left/center/right。
    number_formatは "#,##0" "0.0%" "yyyy/mm/dd" など。col_widthで列幅も設定できる。"""
    wb, path = _open(filename)
    ws = _sheet(wb, sheet or None)
    for part in cell_range.split(","):
        min_c, min_r, max_c, max_r = range_boundaries(part.strip())
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                cell = ws.cell(row=r, column=c)
                font_kw = {}
                f = cell.font
                if bold is not None:
                    font_kw["bold"] = bold
                if italic is not None:
                    font_kw["italic"] = italic
                if font_size is not None:
                    font_kw["size"] = font_size
                if font_color:
                    font_kw["color"] = font_color.lstrip("#").upper()
                if font_kw:
                    cell.font = Font(
                        bold=font_kw.get("bold", f.bold),
                        italic=font_kw.get("italic", f.italic),
                        size=font_kw.get("size", f.size),
                        color=font_kw.get("color", f.color),
                        name=f.name,
                    )
                if bg_color:
                    hexv = bg_color.lstrip("#").upper()
                    cell.fill = PatternFill(start_color=hexv, end_color=hexv, fill_type="solid")
                if align:
                    cell.alignment = Alignment(horizontal=align, vertical=cell.alignment.vertical)
                if number_format:
                    cell.number_format = number_format
        if col_width is not None:
            for c in range(min_c, max_c + 1):
                ws.column_dimensions[get_column_letter(c)].width = col_width
    atomic_save(wb.save, path)
    return f"{cell_range} に書式を設定しました"


@tool
def excel_add_sheet(filename: str, sheet_name: str) -> str:
    """Excelブックに新しいシートを追加する。"""
    wb, path = _open(filename)
    if sheet_name in wb.sheetnames:
        return f"シート「{sheet_name}」は既に存在します"
    wb.create_sheet(sheet_name)
    atomic_save(wb.save, path)
    return f"シート「{sheet_name}」を追加しました"


EXCEL_TOOLS = [excel_create, excel_read, excel_write_cells, excel_write_rows, excel_format, excel_add_sheet]
