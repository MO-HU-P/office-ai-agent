"""Word (.docx) 操作ツール群"""
import json

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from langchain_core.tools import tool

from ...atomic import atomic_save
from ...config import resolve_workspace_path

STYLE_MAP = {
    "normal": None,
    "title": "Title",
    "h1": "Heading 1",
    "h2": "Heading 2",
    "h3": "Heading 3",
    "h4": "Heading 4",
    "bullet": "List Bullet",
    "number": "List Number",
    "quote": "Quote",
}


def _open(filename: str) -> tuple[Document, str]:
    path = resolve_workspace_path(filename, must_exist=True)
    if path.suffix != ".docx":
        raise ValueError("Word文書は .docx ファイルを指定してください")
    return Document(str(path)), str(path)


@tool
def word_create(filename: str, title: str = "") -> str:
    """新しいWord文書(.docx)を作成する。filenameは必ず.docxで終わること。titleを指定すると表題段落を追加する。"""
    path = resolve_workspace_path(filename)
    if path.suffix != ".docx":
        return "エラー: filenameは .docx で終わる必要があります"
    doc = Document()
    if title:
        doc.add_heading(title, level=0)
    atomic_save(doc.save, path)
    return f"{filename} を作成しました"


@tool
def word_read(filename: str, mode: str = "full") -> str:
    """Word文書の内容を読む。段落番号・スタイル付きで全段落と表を返す。編集前に必ず呼ぶこと。
    mode="outline"にすると表題・見出しの段落だけを返す。長い文書はまずoutlineで構造を把握し、
    必要な箇所だけをfullで読むとよい。"""
    doc, _ = _open(filename)
    if mode == "outline":
        lines = []
        for i, p in enumerate(doc.paragraphs):
            style = p.style.name if p.style else "Normal"
            if style == "Title" or style.startswith("Heading"):
                lines.append(f"[{i}] ({style}) {p.text.strip()}")
        lines.append(f"(全{len(doc.paragraphs)}段落, 表{len(doc.tables)}個)")
        return "\n".join(lines)
    lines = []
    for i, p in enumerate(doc.paragraphs):
        style = p.style.name if p.style else "Normal"
        text = p.text.strip()
        lines.append(f"[{i}] ({style}) {text}" if text else f"[{i}] ({style}) <空行>")
    for ti, table in enumerate(doc.tables):
        lines.append(f"--- 表{ti} ({len(table.rows)}行 x {len(table.columns)}列) ---")
        for row in table.rows[:20]:
            lines.append(" | ".join(c.text.strip() for c in row.cells))
    return "\n".join(lines) if lines else "(空の文書)"


@tool
def word_append(filename: str, text: str, style: str = "normal") -> str:
    """Word文書の末尾に段落を追加する。textに改行(\\n)を含めると複数段落として追加される。
    style: normal / title / h1 / h2 / h3 / h4 / bullet(箇条書き) / number(番号付き) / quote(引用)"""
    doc, path = _open(filename)
    style_name = STYLE_MAP.get(style, None)
    count = 0
    for line in text.split("\n"):
        if style == "normal" and not line.strip():
            doc.add_paragraph("")
            continue
        p = doc.add_paragraph(line)
        if style_name:
            try:
                p.style = style_name
            except KeyError:
                pass
        count += 1
    atomic_save(doc.save, path)
    return f"{filename} に{count}段落を追加しました (style={style})"


@tool
def word_find(filename: str, query: str = "", style: str = "") -> str:
    """Word文書から条件に合う段落を探し、段落番号を返す。長い文書で修正箇所を探すときに
    全文をword_readで読む代わりに使う。見つかった番号はそのままword_batch_editに渡せる。
    query: 探す文字列(部分一致)。style: スタイル指定(normal/h1/h2/h3/h4/bullet/number/quote または実際のスタイル名)。
    両方指定すると両方に合う段落だけを返す。少なくとも一方は指定すること。"""
    doc, _ = _open(filename)
    if not query and not style:
        return "エラー: query(探す文字列)かstyle(スタイル)の少なくとも一方を指定してください"
    style_name = STYLE_MAP.get(style, style) if style else None
    hits = []
    for i, p in enumerate(doc.paragraphs):
        p_style = p.style.name if p.style else "Normal"
        if style and p_style != style_name:
            continue
        if query and query not in p.text:
            continue
        text = p.text.strip()
        hits.append(f"[{i}] ({p_style}) {text[:60]}" if text else f"[{i}] ({p_style}) <空行>")
    if not hits:
        return "条件に合う段落は見つかりませんでした"
    lines = hits[:50]
    if len(hits) > 50:
        lines.append(f"...ほか{len(hits) - 50}件")
    lines.append(f"(計{len(hits)}段落が該当)")
    return "\n".join(lines)


@tool
def word_edit_paragraph(filename: str, index: int, new_text: str = "", delete: bool = False) -> str:
    """既存段落を編集する。indexはword_readで表示される段落番号。delete=Trueでその段落を削除。"""
    doc, path = _open(filename)
    if index < 0 or index >= len(doc.paragraphs):
        return f"エラー: 段落番号 {index} は範囲外です (0〜{len(doc.paragraphs) - 1})"
    p = doc.paragraphs[index]
    if delete:
        p._element.getparent().remove(p._element)
        atomic_save(doc.save, path)
        return f"段落 [{index}] を削除しました"
    # 既存の書式(スタイル)を保ちつつテキストだけ差し替える
    for run in list(p.runs):
        run._element.getparent().remove(run._element)
    p.add_run(new_text)
    atomic_save(doc.save, path)
    return f"段落 [{index}] を更新しました"


@tool
def word_batch_edit(filename: str, edits: list[dict]) -> str:
    """複数の段落をまとめて編集・削除する。2箇所以上を直すときはword_edit_paragraphを繰り返さず必ずこちらを使う。
    editsの各要素は {"index": 段落番号, "new_text": "新しい本文"} または {"index": 段落番号, "delete": true}。
    indexはword_readで表示される番号(編集前の番号のままでよい)。一部が失敗しても残りは適用される。"""
    doc, path = _open(filename)
    paras = list(doc.paragraphs)  # 削除で番号がずれないよう、編集前の番号で対象を確定しておく
    done: set[int] = set()
    ok_count = 0
    failures: list[str] = []
    for e in edits:
        idx = e.get("index")
        if not isinstance(idx, int) or idx < 0 or idx >= len(paras):
            failures.append(f"[{idx}] 段落番号が範囲外です (0〜{len(paras) - 1})")
            continue
        if idx in done:
            failures.append(f"[{idx}] 同じ段落が2回指定されています")
            continue
        p = paras[idx]
        if e.get("delete"):
            p._element.getparent().remove(p._element)
        elif e.get("new_text") is not None:
            for run in list(p.runs):
                run._element.getparent().remove(run._element)
            p.add_run(str(e["new_text"]))
        else:
            failures.append(f"[{idx}] new_text か delete を指定してください")
            continue
        done.add(idx)
        ok_count += 1
    if ok_count:
        atomic_save(doc.save, path)
    if not ok_count:
        return "エラー: 1件も適用できませんでした\n" + "\n".join(failures)
    result = f"{filename} の{ok_count}段落を更新しました"
    if failures:
        result += "\n適用できなかったもの:\n" + "\n".join(failures)
    return result


_ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}
_ALIGN_NAMES = {v: k for k, v in _ALIGN_MAP.items()}


def _dump_one_style(style) -> dict:
    """スタイル1つ分の書式を辞書にする。未設定の属性は継承元(base_style)をたどって解決する。"""
    d: dict = {}
    seen = []
    s = style
    while s is not None and s not in seen and len(seen) < 8:
        seen.append(s)
        f = s.font
        if "font" not in d:
            # 日本語フォント(eastAsia)を優先し、なければ欧文フォント名を使う
            rpr = s.element.rPr
            ea = rpr.rFonts.get(qn("w:eastAsia")) if rpr is not None and rpr.rFonts is not None else None
            if ea or f.name:
                d["font"] = ea or f.name
        if "size" not in d and f.size is not None:
            d["size"] = f.size.pt
        if "bold" not in d and f.bold is not None:
            d["bold"] = f.bold
        if "italic" not in d and f.italic is not None:
            d["italic"] = f.italic
        if "color" not in d and f.color is not None and f.color.rgb is not None:
            d["color"] = f"#{f.color.rgb}"
        pf = getattr(s, "paragraph_format", None)
        if pf is not None:
            if "align" not in d and pf.alignment is not None and pf.alignment in _ALIGN_NAMES:
                d["align"] = _ALIGN_NAMES[pf.alignment]
            if "space_before" not in d and pf.space_before is not None:
                d["space_before"] = pf.space_before.pt
            if "space_after" not in d and pf.space_after is not None:
                d["space_after"] = pf.space_after.pt
        s = s.base_style
    return d


@tool
def word_dump_style(filename: str) -> str:
    """お手本のWord文書から、見出し構造(outline)とスタイル設定(styles: フォント・サイズ・太字・色・
    配置・段落前後の間隔)をJSONで取り出す。「この文書と同じ体裁で作って」と言われたら、
    まずこれでお手本を調べ、styles部分をword_apply_styleで新しい文書に適用してから本文を書くこと。"""
    doc, _ = _open(filename)
    used: list[str] = []
    outline: list[str] = []
    for p in doc.paragraphs:
        name = p.style.name if p.style else "Normal"
        if name not in used:
            used.append(name)
        if name == "Title" or name.startswith("Heading"):
            outline.append(f"({name}) {p.text.strip()}")
    styles = {}
    for name in used:
        try:
            styles[name] = _dump_one_style(doc.styles[name])
        except KeyError:
            continue
    if len(outline) > 50:
        outline = outline[:50] + [f"...ほか{len(outline) - 50}件"]
    return json.dumps({"styles": styles, "outline": outline}, ensure_ascii=False, indent=1)


@tool
def word_apply_style(filename: str, styles: dict) -> str:
    """word_dump_styleで取り出したスタイル設定をWord文書に適用し、お手本と同じ体裁にする。
    stylesはスタイル名→設定の辞書。設定できる項目: font(フォント名) / size(pt) / bold / italic /
    color("#RRGGBB") / align(left・center・right・justify) / space_before・space_after(段落前後の間隔pt)。
    例: {"Heading 1": {"font": "游ゴシック", "size": 14, "bold": true, "color": "#1A73E8"}}
    適用後は既存段落にも、word_appendで追加する段落にも同じ体裁が効く。"""
    doc, path = _open(filename)
    ok_count = 0
    failures: list[str] = []
    for name, conf in styles.items():
        if not isinstance(conf, dict):
            failures.append(f"「{name}」の設定が辞書形式ではありません")
            continue
        try:
            style = doc.styles[name]
        except KeyError:
            failures.append(f"「{name}」というスタイルはこの文書にありません")
            continue
        try:
            f = style.font
            if conf.get("font"):
                font_name = str(conf["font"])
                f.name = font_name
                # 日本語文字にはeastAsiaフォントが使われるため、両方に同じフォントを設定する
                rpr = style.element.get_or_add_rPr()
                rfonts = rpr.rFonts if rpr.rFonts is not None else rpr.get_or_add_rFonts()
                rfonts.set(qn("w:eastAsia"), font_name)
            if conf.get("size") is not None:
                f.size = Pt(float(conf["size"]))
            if conf.get("bold") is not None:
                f.bold = bool(conf["bold"])
            if conf.get("italic") is not None:
                f.italic = bool(conf["italic"])
            if conf.get("color"):
                f.color.rgb = RGBColor.from_string(str(conf["color"]).lstrip("#").upper())
            pf = style.paragraph_format
            if conf.get("align") in _ALIGN_MAP:
                pf.alignment = _ALIGN_MAP[conf["align"]]
            if conf.get("space_before") is not None:
                pf.space_before = Pt(float(conf["space_before"]))
            if conf.get("space_after") is not None:
                pf.space_after = Pt(float(conf["space_after"]))
            ok_count += 1
        except (ValueError, TypeError) as e:
            failures.append(f"「{name}」の適用に失敗しました: {e}")
    if ok_count:
        atomic_save(doc.save, path)
    if not ok_count:
        return "エラー: 1つもスタイルを適用できませんでした\n" + "\n".join(failures)
    result = f"{filename} に{ok_count}個のスタイルを適用しました"
    if failures:
        result += "\n適用できなかったもの:\n" + "\n".join(failures)
    return result


@tool
def word_add_table(filename: str, rows: list[list[str]], header: bool = True) -> str:
    """Word文書の末尾に表を追加する。rowsは2次元配列(行のリスト)。header=Trueなら1行目を太字ヘッダーにする。"""
    doc, path = _open(filename)
    if not rows or not rows[0]:
        return "エラー: rowsが空です"
    n_cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=n_cols)
    table.style = "Light Grid Accent 1"
    for ri, row in enumerate(rows):
        for ci in range(n_cols):
            val = str(row[ci]) if ci < len(row) else ""
            cell = table.cell(ri, ci)
            cell.text = val
            if header and ri == 0:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.bold = True
                        run.font.size = Pt(10.5)
    atomic_save(doc.save, path)
    return f"{filename} に {len(rows)}x{n_cols} の表を追加しました"


WORD_TOOLS = [
    word_create,
    word_read,
    word_find,
    word_append,
    word_edit_paragraph,
    word_batch_edit,
    word_add_table,
    word_dump_style,
    word_apply_style,
]
