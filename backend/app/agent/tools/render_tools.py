"""文書のページ画像化ツール(ビジョン対応モデルのときだけ有効化される)。

Ollamaのtoolロールは画像を運べないため、このツールは画像パスをマーカー付き文字列で返し、
loop.py がそれを検出して base64 画像入りの HumanMessage をLLMへ注入する。
"""
from langchain_core.tools import tool

from ...config import resolve_workspace_path
from ...services.preview import render_page_png

# loop.py がこのprefixを検出して画像を添付する。"画像パス|LLM向けの結果文" の形式
IMAGE_RESULT_PREFIX = "__RENDER_IMAGE__:"


@tool
def render_page(filename: str, page: int = 1) -> str:
    """Word文書(.docx)またはPowerPoint(.pptx)の指定ページ/スライドを画像化し、
    自分の目でレイアウトを確認する。テキストでは分からない見た目の問題
    (文字のはみ出し・図形の重なり・不自然な余白や改行)を見つけるために使う。
    画像は次のメッセージに添付されてくるので、それを見て問題があれば編集ツールで直し、
    直したらもう一度これで確認する。pageはスライド/ページ番号(1始まり)。"""
    path = resolve_workspace_path(filename, must_exist=True)
    if path.suffix not in (".docx", ".pptx"):
        return "エラー: 画像化できるのは .docx / .pptx だけです(Excelは excel_read で内容を確認してください)"
    try:
        png, total = render_page_png(path, page)
    except ValueError as e:
        return f"エラー: {e}"
    except Exception:
        return "エラー: 画像化に失敗しました。ファイルが壊れていないか確認してください"
    return f"{IMAGE_RESULT_PREFIX}{png}|{filename} の {page}ページ目(全{total}ページ)を画像化しました。次のメッセージに画像が添付されます。"


RENDER_TOOLS = [render_page]
