"""右ペイン用のプレビュー生成。

- Excel: openpyxlでセル値+基本スタイルをJSON化(フロントでSheets風グリッド描画)
- PowerPoint: LibreOffice headlessでPDF化 → pdftoppmでスライドPNG化
- Word: フロント側で docx-preview がrawファイルを直接描画するため変換不要
"""
import asyncio
import datetime as dt
import shutil
import subprocess
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from ..config import PREVIEW_CACHE_DIR

MAX_ROWS = 300
MAX_COLS = 60

_soffice_lock = asyncio.Lock()  # LibreOfficeは同時起動に弱いため直列化


def _color_hex(color) -> str | None:
    try:
        rgb = color.rgb
    except AttributeError:
        return None
    if not isinstance(rgb, str) or len(rgb) != 8:
        return None
    if rgb == "00000000":
        return None
    return "#" + rgb[2:]


def _display_value(value, number_format: str) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, dt.datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return value.strftime("%Y/%m/%d")
        return value.strftime("%Y/%m/%d %H:%M")
    if isinstance(value, dt.date):
        return value.strftime("%Y/%m/%d")
    if isinstance(value, dt.time):
        return value.strftime("%H:%M")
    if isinstance(value, float):
        nf = number_format or ""
        if "%" in nf:
            digits = 1 if "0.0" in nf else 0
            return f"{value * 100:.{digits}f}%"
        if value == int(value) and abs(value) < 1e15:
            value = int(value)
        else:
            value = round(value, 6)
    if isinstance(value, (int, float)) and "#,##" in (number_format or ""):
        return f"{value:,}"
    return str(value)


def excel_preview(path: Path) -> dict[str, Any]:
    wb_v = load_workbook(str(path), data_only=True, read_only=False)
    wb_f = load_workbook(str(path), data_only=False, read_only=False)
    sheets = []
    for name in wb_v.sheetnames:
        ws_v, ws_f = wb_v[name], wb_f[name]
        n_rows = min(max(ws_v.max_row, 1), MAX_ROWS)
        n_cols = min(max(ws_v.max_column, 1), MAX_COLS)
        rows = []
        for r in range(1, n_rows + 1):
            row_cells = []
            for c in range(1, n_cols + 1):
                cv = ws_v.cell(row=r, column=c)
                cf = ws_f.cell(row=r, column=c)
                formula = None
                if isinstance(cf.value, str) and cf.value.startswith("="):
                    formula = cf.value
                raw = cv.value if cv.value is not None else (None if formula else cf.value)
                display = _display_value(raw, cv.number_format)
                if display == "" and formula:
                    display = formula  # 未計算の数式は数式文字列を表示
                cell: dict[str, Any] = {}
                if display != "":
                    cell["v"] = display
                if formula:
                    cell["f"] = formula
                style: dict[str, Any] = {}
                font = cv.font
                if font.bold:
                    style["b"] = 1
                if font.italic:
                    style["i"] = 1
                if font.size and float(font.size) != 11.0:
                    style["fs"] = float(font.size)
                fc = _color_hex(font.color) if font.color else None
                if fc and fc != "#000000":
                    style["fc"] = fc
                if cv.fill and cv.fill.patternType == "solid":
                    bg = _color_hex(cv.fill.start_color)
                    if bg:
                        style["bg"] = bg
                if cv.alignment and cv.alignment.horizontal in ("center", "right"):
                    style["ha"] = cv.alignment.horizontal
                elif isinstance(raw, (int, float)) and not isinstance(raw, bool):
                    style["ha"] = "right"
                if style:
                    cell["s"] = style
                row_cells.append(cell)
            rows.append(row_cells)
        merges = []
        for m in ws_v.merged_cells.ranges:
            if m.min_row <= n_rows and m.min_col <= n_cols:
                merges.append({
                    "r": m.min_row - 1,
                    "c": m.min_col - 1,
                    "rs": min(m.max_row, n_rows) - m.min_row + 1,
                    "cs": min(m.max_col, n_cols) - m.min_col + 1,
                })
        col_widths = []
        for c in range(1, n_cols + 1):
            dim = ws_v.column_dimensions.get(get_column_letter(c))
            width = dim.width if dim and dim.width else 8.43
            col_widths.append(round(width * 7.5 + 5))
        sheets.append({
            "name": name,
            "rows": rows,
            "merges": merges,
            "colWidths": col_widths,
            "truncated": ws_v.max_row > MAX_ROWS or ws_v.max_column > MAX_COLS,
        })
    wb_v.close()
    wb_f.close()
    return {"type": "excel", "sheets": sheets}


def _cache_dir_for(path: Path) -> Path:
    mtime = int(path.stat().st_mtime * 1000)
    return PREVIEW_CACHE_DIR / f"{path.stem}__{mtime}"


async def pptx_preview(path: Path) -> dict[str, Any]:
    """PPTXをスライドPNG群に変換し、画像URLのリストを返す(mtimeでキャッシュ)。"""
    cache_dir = _cache_dir_for(path)
    if not cache_dir.exists():
        async with _soffice_lock:
            if not cache_dir.exists():
                await asyncio.to_thread(_convert_pptx, path, cache_dir)
    def _slide_no(p: Path) -> int:
        try:
            return int(p.stem.split("-")[-1])
        except ValueError:
            return 0

    images = sorted(cache_dir.glob("slide-*.png"), key=_slide_no)
    if not images:
        raise RuntimeError("スライド画像の生成に失敗しました")
    return {
        "type": "pptx",
        "slides": [f"/api/preview_cache/{cache_dir.name}/{img.name}" for img in images],
    }


def _convert_pptx(path: Path, cache_dir: Path):
    tmp_dir = cache_dir.with_suffix(".tmp")
    shutil.rmtree(tmp_dir, ignore_errors=True)
    tmp_dir.mkdir(parents=True)
    try:
        subprocess.run(
            [
                "soffice", "--headless", "--norestore",
                f"-env:UserInstallation=file:///tmp/lo_profile",
                "--convert-to", "pdf", "--outdir", str(tmp_dir), str(path),
            ],
            check=True, capture_output=True, timeout=120,
        )
        pdfs = list(tmp_dir.glob("*.pdf"))
        if not pdfs:
            raise RuntimeError("PDF変換に失敗しました")
        subprocess.run(
            ["pdftoppm", "-png", "-r", "110", str(pdfs[0]), str(tmp_dir / "slide")],
            check=True, capture_output=True, timeout=120,
        )
        pdfs[0].unlink()
        # 古いキャッシュを掃除してから確定(作業中のtmpディレクトリ自身は除外)
        for old in PREVIEW_CACHE_DIR.glob(f"{path.stem}__*"):
            if old != tmp_dir:
                shutil.rmtree(old, ignore_errors=True)
        tmp_dir.rename(cache_dir)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
