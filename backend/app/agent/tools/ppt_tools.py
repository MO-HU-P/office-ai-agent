"""PowerPoint (.pptx) 操作ツール群"""
from typing import Optional

from langchain_core.tools import tool
from pptx import Presentation
from pptx.util import Pt

from ...atomic import atomic_save
from ...config import resolve_workspace_path

LAYOUT_TITLE = 0
LAYOUT_TITLE_CONTENT = 1
LAYOUT_SECTION = 2
LAYOUT_TITLE_ONLY = 5
LAYOUT_BLANK = 6


def _open(filename: str):
    path = resolve_workspace_path(filename, must_exist=True)
    if path.suffix != ".pptx":
        raise ValueError("PowerPointは .pptx ファイルを指定してください")
    return Presentation(str(path)), str(path)


def _fill_bullets(body_shape, bullets: list[str]):
    tf = body_shape.text_frame
    tf.clear()
    first = True
    for item in bullets:
        stripped = item.lstrip(" ")
        level = min((len(item) - len(stripped)) // 2, 4)  # 先頭スペース2つでインデント1段
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        p.text = stripped
        p.level = level
        for run in p.runs:
            run.font.size = Pt(20 if level == 0 else 18)
        first = False


@tool
def ppt_create(filename: str, title: str = "", subtitle: str = "") -> str:
    """新しいPowerPoint(.pptx)を作成する。filenameは必ず.pptxで終わること。
    titleを指定するとタイトルスライドが追加される。"""
    path = resolve_workspace_path(filename)
    if path.suffix != ".pptx":
        return "エラー: filenameは .pptx で終わる必要があります"
    prs = Presentation()
    if title:
        slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_TITLE])
        slide.shapes.title.text = title
        if subtitle and len(slide.placeholders) > 1:
            slide.placeholders[1].text = subtitle
    atomic_save(prs.save, path)
    return f"{filename} を作成しました"


@tool
def ppt_add_slide(filename: str, title: str, bullets: Optional[list[str]] = None, section_header: bool = False) -> str:
    """PowerPointにスライドを1枚追加する。bulletsは箇条書きのリスト。
    項目の先頭にスペース2つを付けると1段インデントされる(例: ["親項目", "  子項目"])。
    section_header=Trueにすると章の区切り用スライドになる。"""
    prs, path = _open(filename)
    if section_header:
        slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_SECTION])
        slide.shapes.title.text = title
    elif bullets:
        slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_TITLE_CONTENT])
        slide.shapes.title.text = title
        _fill_bullets(slide.placeholders[1], bullets)
    else:
        slide = prs.slides.add_slide(prs.slide_layouts[LAYOUT_TITLE_ONLY])
        slide.shapes.title.text = title
    atomic_save(prs.save, path)
    return f"{filename} にスライド{len(prs.slides)}「{title}」を追加しました"


@tool
def ppt_read(filename: str, mode: str = "full") -> str:
    """PowerPointの内容を読む。スライド番号ごとにタイトルと本文テキストを返す。編集前に必ず呼ぶこと。
    mode="outline"にすると各スライドのタイトルだけを返す。枚数が多いときはまずoutlineで全体を把握すること。"""
    prs, _ = _open(filename)
    lines = [f"全{len(prs.slides)}枚"]
    if mode == "outline":
        for i, slide in enumerate(prs.slides, start=1):
            title = slide.shapes.title.text.strip() if slide.shapes.title is not None else ""
            lines.append(f"スライド{i}: {title or '(タイトルなし)'}")
        return "\n".join(lines)
    for i, slide in enumerate(prs.slides, start=1):
        lines.append(f"--- スライド{i} ---")
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                text = "".join(run.text for run in p.runs)
                if text.strip():
                    prefix = "  " * p.level
                    lines.append(f"{prefix}{text}")
    return "\n".join(lines)


def _apply_slide_edit(prs, slide_number: int, title: str, bullets: Optional[list[str]]) -> Optional[str]:
    """1枚分の編集を適用する。失敗時はエラーメッセージを返し、成功時はNoneを返す。"""
    if not isinstance(slide_number, int) or slide_number < 1 or slide_number > len(prs.slides):
        return f"スライド番号 {slide_number} は範囲外です (1〜{len(prs.slides)})"
    slide = prs.slides[slide_number - 1]
    if title and slide.shapes.title is not None:
        slide.shapes.title.text = title
    if bullets is not None:
        body = None
        for shape in slide.placeholders:
            if shape.placeholder_format.idx != 0 and shape.has_text_frame:
                body = shape
                break
        if body is None:
            return f"スライド{slide_number}には本文プレースホルダーがありません"
        _fill_bullets(body, bullets)
    return None


@tool
def ppt_edit_slide(filename: str, slide_number: int, title: str = "", bullets: Optional[list[str]] = None) -> str:
    """既存スライドのタイトル・箇条書きを書き換える。slide_numberは1始まり(ppt_readの番号)。
    titleまたはbulletsのうち指定したものだけが置き換えられる。"""
    prs, path = _open(filename)
    error = _apply_slide_edit(prs, slide_number, title, bullets)
    if error:
        return f"エラー: {error}"
    atomic_save(prs.save, path)
    return f"スライド{slide_number}を更新しました"


@tool
def ppt_batch_edit(filename: str, edits: list[dict]) -> str:
    """複数のスライドをまとめて書き換える。2枚以上を直すときはppt_edit_slideを繰り返さず必ずこちらを使う。
    editsの各要素は {"slide_number": スライド番号(1始まり), "title": "新タイトル", "bullets": ["箇条書き", ...]}。
    titleとbulletsは指定したものだけが置き換えられる。一部が失敗しても残りは適用される。"""
    prs, path = _open(filename)
    ok_count = 0
    failures: list[str] = []
    for e in edits:
        error = _apply_slide_edit(prs, e.get("slide_number"), e.get("title") or "", e.get("bullets"))
        if error:
            failures.append(error)
        else:
            ok_count += 1
    if ok_count:
        atomic_save(prs.save, path)
    if not ok_count:
        return "エラー: 1枚も更新できませんでした\n" + "\n".join(failures)
    result = f"{filename} の{ok_count}枚のスライドを更新しました"
    if failures:
        result += "\n更新できなかったもの:\n" + "\n".join(failures)
    return result


PPT_TOOLS = [ppt_create, ppt_add_slide, ppt_read, ppt_edit_slide, ppt_batch_edit]
