"""Word (.docx) 操作ツール群"""
from docx import Document
from docx.shared import Pt
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


WORD_TOOLS = [word_create, word_read, word_append, word_edit_paragraph, word_batch_edit, word_add_table]
