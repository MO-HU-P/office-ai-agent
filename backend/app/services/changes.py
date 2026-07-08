"""「AIが最後にこのファイルをどう変えたか」の差分ビューを組み立てる。

history(自動バックアップ)の「最後の変更前」の世代と現在のファイルを比較し、
フロントエンドがそのまま色分け表示できる構造化された差分を返す。
Wordは段落単位、Excelはセル単位(シート名!A1)、PowerPointはスライド内テキスト単位。
"""
import difflib
import logging

from .. import config
from . import history
from .textextract import CHANGE_VIEW_EXTRACTORS

logger = logging.getLogger(__name__)

# 表示しきれない巨大差分を返さないための上限
_MAX_LINES = 400


def _no_changes(reason: str) -> dict:
    return {"available": False, "reason": reason}


def build_changes(filename: str) -> dict:
    """現在のファイルと「最後の変更前」バックアップの差分を返す。

    返り値(available=Trueのとき):
      base_time: 比較対象のバックアップが取られた時刻(=その編集が行われた時刻)
      lines: [{"op": "add"|"del"|"ctx"|"skip", "text": str}, ...]
      added / removed: 追加・削除行数, truncated: 上限で打ち切ったか
    """
    path = config.resolve_workspace_path(filename, must_exist=True)
    extractor = CHANGE_VIEW_EXTRACTORS.get(path.suffix.lower())
    if extractor is None:
        return _no_changes("このファイル形式は変更箇所の表示に対応していません(.docx / .xlsx / .pptx / .csv / .txt)")
    base = history.last_change_base(filename)
    if base is None:
        return _no_changes("このファイルにはまだバックアップがありません(新規作成後、編集されると表示できます)")

    try:
        old_lines = extractor(str(base.path))
        new_lines = extractor(str(path))
    except Exception:
        logger.exception("差分の抽出に失敗: %s", filename)
        return _no_changes("ファイルの読み取りに失敗したため、変更箇所を表示できません")

    lines: list[dict] = []
    added = removed = 0
    for line in difflib.unified_diff(old_lines, new_lines, lineterm="", n=1):
        if line.startswith(("---", "+++")):
            continue
        if line.startswith("@@"):
            lines.append({"op": "skip", "text": "……"})
        elif line.startswith("+"):
            added += 1
            lines.append({"op": "add", "text": line[1:]})
        elif line.startswith("-"):
            removed += 1
            lines.append({"op": "del", "text": line[1:]})
        else:
            lines.append({"op": "ctx", "text": line[1:] if line.startswith(" ") else line})

    if added == 0 and removed == 0:
        return _no_changes("最後のバックアップから内容の変更はありません(書式だけの変更は表示できません)")

    truncated = len(lines) > _MAX_LINES
    return {
        "available": True,
        "filename": filename,
        "base_version": base.id,
        "base_time": base.timestamp.isoformat(),
        "base_label": base.label,
        "added": added,
        "removed": removed,
        "truncated": truncated,
        "lines": lines[:_MAX_LINES],
    }
