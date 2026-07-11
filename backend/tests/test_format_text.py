"""キーワードへの文字書式ツール(word/ppt/excel_format_text)のテスト。

「重要な用語を赤字にして」のような依頼で、キーワードの部分だけに色が付き、
前後の文字や既存の書式(太字等)が壊れないことを確認する。
特にWord/PPTでは、キーワードが複数のrunにまたがっている場合の分割処理が肝。
"""
from docx import Document
from docx.shared import RGBColor as DocxRGB
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from pptx import Presentation
from pptx.util import Cm

from app.agent.tools.excel_tools import excel_format_text
from app.agent.tools.ppt_tools import ppt_format_text
from app.agent.tools.word_tools import word_format_text
from app.atomic import atomic_save

RED = DocxRGB(0xFF, 0x00, 0x00)


# ---------- Word ----------

def _make_docx(ws_dir, build):
    path = ws_dir / "doc.docx"
    doc = Document()
    build(doc)
    atomic_save(doc.save, path)
    return path


def test_word_colors_only_keyword(ws):
    _make_docx(ws, lambda d: d.add_paragraph("機械学習ではファインチューニングが重要です。"))
    result = word_format_text.invoke(
        {"filename": "doc.docx", "keywords": ["ファインチューニング"], "color": "#FF0000"}
    )
    assert "1段落・1箇所" in result
    p = Document(str(ws / "doc.docx")).paragraphs[0]
    assert p.text == "機械学習ではファインチューニングが重要です。"  # 本文は変わらない
    red = [r.text for r in p.runs if r.font.color and r.font.color.rgb == RED]
    assert red == ["ファインチューニング"]


def test_word_keyword_split_across_runs(ws):
    def build(d):
        p = d.add_paragraph()
        p.add_run("手法として")
        p.add_run("ファインチ")  # Wordが不規則にrunを分けた状態を再現
        p.add_run("ューニングを使う")

    _make_docx(ws, build)
    result = word_format_text.invoke(
        {"filename": "doc.docx", "keywords": ["ファインチューニング"], "color": "#FF0000"}
    )
    assert "1箇所" in result
    p = Document(str(ws / "doc.docx")).paragraphs[0]
    assert p.text == "手法としてファインチューニングを使う"
    red = "".join(r.text for r in p.runs if r.font.color and r.font.color.rgb == RED)
    assert red == "ファインチューニング"
    not_red = "".join(r.text for r in p.runs if not (r.font.color and r.font.color.rgb == RED))
    assert not_red == "手法としてを使う"


def test_word_split_preserves_existing_format(ws):
    def build(d):
        p = d.add_paragraph()
        run = p.add_run("前段RLHF後段")
        run.bold = True

    _make_docx(ws, build)
    word_format_text.invoke({"filename": "doc.docx", "keywords": ["RLHF"], "color": "#FF0000"})
    p = Document(str(ws / "doc.docx")).paragraphs[0]
    assert p.text == "前段RLHF後段"
    assert all(r.bold for r in p.runs)  # 分割で増えたrunも元の太字を引き継ぐ
    red = [r.text for r in p.runs if r.font.color and r.font.color.rgb == RED]
    assert red == ["RLHF"]


def test_word_formats_table_and_multiple_hits(ws):
    def build(d):
        d.add_paragraph("報酬設計と報酬設計")
        table = d.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "報酬設計の一覧"

    _make_docx(ws, build)
    result = word_format_text.invoke(
        {"filename": "doc.docx", "keywords": ["報酬設計"], "bold": True}
    )
    assert "2段落" in result and "3箇所" in result
    doc = Document(str(ws / "doc.docx"))
    bold = [r.text for p in doc.paragraphs for r in p.runs if r.bold]
    assert bold == ["報酬設計", "報酬設計"]
    cell_p = doc.tables[0].cell(0, 0).paragraphs[0]
    assert [r.text for r in cell_p.runs if r.bold] == ["報酬設計"]


def test_word_not_found_and_bad_args(ws):
    _make_docx(ws, lambda d: d.add_paragraph("本文"))
    assert "見つかりませんでした" in word_format_text.invoke(
        {"filename": "doc.docx", "keywords": ["存在しない語"], "color": "#FF0000"}
    )
    assert "エラー" in word_format_text.invoke({"filename": "doc.docx", "keywords": [], "color": "#FF0000"})
    assert "エラー" in word_format_text.invoke({"filename": "doc.docx", "keywords": ["本文"]})
    assert "エラー" in word_format_text.invoke({"filename": "doc.docx", "keywords": ["本文"], "color": "赤"})


# ---------- PowerPoint ----------

def _make_pptx(ws_dir):
    path = ws_dir / "deck.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 白紙
    box = slide.shapes.add_textbox(Cm(1), Cm(1), Cm(10), Cm(3))
    box.text_frame.text = "強化学習とRLHFの概要"
    table = slide.shapes.add_table(1, 1, Cm(1), Cm(5), Cm(8), Cm(2)).table
    table.cell(0, 0).text = "RLHFの手順"
    atomic_save(prs.save, path)
    return path


def test_ppt_colors_keyword_in_textbox_and_table(ws):
    _make_pptx(ws)
    result = ppt_format_text.invoke(
        {"filename": "deck.pptx", "keywords": ["RLHF"], "color": "#FF0000", "bold": True}
    )
    assert "2箇所" in result
    prs = Presentation(str(ws / "deck.pptx"))
    slide = prs.slides[0]
    box_p = slide.shapes[0].text_frame.paragraphs[0]
    assert "".join(r.text for r in box_p.runs) == "強化学習とRLHFの概要"
    red = [r for r in box_p.runs if r.font.color.type is not None and str(r.font.color.rgb) == "FF0000"]
    assert [r.text for r in red] == ["RLHF"] and red[0].font.bold is True
    cell_p = slide.shapes[1].table.cell(0, 0).text_frame.paragraphs[0]
    assert "".join(r.text for r in cell_p.runs) == "RLHFの手順"
    assert [r.text for r in cell_p.runs if r.font.color.type is not None] == ["RLHF"]


def test_ppt_slide_number_out_of_range(ws):
    _make_pptx(ws)
    result = ppt_format_text.invoke(
        {"filename": "deck.pptx", "keywords": ["RLHF"], "color": "#FF0000", "slide_number": 5}
    )
    assert "範囲外" in result


# ---------- Excel ----------

def test_excel_formats_matching_cells(ws):
    path = ws / "book.xlsx"
    wb = Workbook()
    sh = wb.active
    sh["A1"] = "項目"
    sh["A2"] = "強化学習の説明"
    sh["A3"] = "その他"
    sh["B2"] = '="強化学習"'  # 数式は対象外
    sh["A4"] = "強化学習"
    sh["A4"].font = Font(bold=True)
    atomic_save(wb.save, path)

    result = excel_format_text.invoke(
        {"filename": "book.xlsx", "keywords": ["強化学習"], "color": "#FF0000"}
    )
    assert "2セル" in result and "A2" in result and "A4" in result
    wb2 = load_workbook(str(path))
    sh2 = wb2.active
    assert sh2["A2"].font.color.rgb == "00FF0000"
    assert sh2["A4"].font.color.rgb == "00FF0000"
    assert sh2["A4"].font.bold is True  # 既存の太字は保たれる
    assert sh2["A3"].font.color is None or sh2["A3"].font.color.rgb != "00FF0000"
    assert sh2["B2"].font.color is None or sh2["B2"].font.color.rgb != "00FF0000"


def test_excel_missing_sheet_and_not_found(ws):
    path = ws / "book.xlsx"
    wb = Workbook()
    wb.active["A1"] = "データ"
    atomic_save(wb.save, path)
    assert "ありません" in excel_format_text.invoke(
        {"filename": "book.xlsx", "keywords": ["データ"], "color": "#FF0000", "sheet": "無いシート"}
    )
    assert "ありませんでした" in excel_format_text.invoke(
        {"filename": "book.xlsx", "keywords": ["無い語"], "color": "#FF0000"}
    )
