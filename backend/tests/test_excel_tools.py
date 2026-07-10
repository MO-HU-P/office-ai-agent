"""Excel書き込みツールの上書き警告のテスト。

既に値があるセルに書き込んだとき、戻り値に「⚠️ 上書きしました」の注意書きが付き、
モデルが誤った位置への書き込み(既存の表を消す等)に気づけることを確認する。
"""
from openpyxl import Workbook

from app.agent.tools.excel_tools import excel_write_cells, excel_write_rows
from app.atomic import atomic_save


def _make_book(ws_dir):
    """A1:B3にヘッダー+2行のデータが入ったブックを作る。"""
    path = ws_dir / "book.xlsx"
    wb = Workbook()
    sh = wb.active
    sh.append(["Group", "Value"])
    sh.append(["A", 12])
    sh.append(["A", 14])
    atomic_save(wb.save, path)


def test_write_rows_warns_on_overwrite(ws):
    _make_book(ws)
    result = excel_write_rows.invoke(
        {"filename": "book.xlsx", "start_cell": "A2", "rows": [["X", 1], ["Y", 2]]}
    )
    assert "書き込みました" in result
    assert "⚠️" in result and "4個" in result and "A2" in result


def test_write_rows_no_warning_on_empty_area(ws):
    _make_book(ws)
    result = excel_write_rows.invoke(
        {"filename": "book.xlsx", "start_cell": "A5", "rows": [["X", 1]]}
    )
    assert "⚠️" not in result


def test_write_cells_warns_only_for_nonempty_cells(ws):
    _make_book(ws)
    result = excel_write_cells.invoke(
        {"filename": "book.xlsx", "cells": {"B2": 99, "D1": "新規"}}
    )
    assert "⚠️" in result and "1個" in result and "B2" in result
    assert "D1" not in result.split("⚠️")[1]  # 空セルへの書き込みは警告に含めない
