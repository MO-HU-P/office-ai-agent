"""巻き戻し(自動バックアップからの復元)ツール。

ファイルが編集・削除されるたびに編集前の状態が自動保存されている(services/history)。
「さっきの編集を取り消して」「元に戻して」に、このツールで応える。
"""
from langchain_core.tools import tool

from ...services import history

# ラベル(t3 / manual / restore)を人が読める説明にする
_LABEL_TEXTS = {"manual": "アップロード等の操作前", "restore": "巻き戻し実行前"}


def _describe_label(label: str) -> str:
    if label.startswith("t") and label[1:].isdigit():
        return "AI編集前"
    return _LABEL_TEXTS.get(label, label)


@tool
def list_file_versions(filename: str) -> str:
    """ファイルの自動バックアップ(編集前の状態)の一覧を新しい順に返す。
    どの時点に戻すか選ぶために使う。各行の先頭のIDを restore_file の version に渡せる。"""
    versions = history.list_versions(filename)
    if not versions:
        return f"「{filename}」のバックアップはまだありません(新規作成後、編集されると自動で作られます)"
    lines = [
        f"{v.id}  {v.timestamp.strftime('%Y-%m-%d %H:%M:%S')}  ({_describe_label(v.label)}, {v.size:,} bytes)"
        for v in versions
    ]
    return f"「{filename}」のバックアップ({len(versions)}件、新しい順):\n" + "\n".join(lines)


@tool
def restore_file(filename: str, version: str = "") -> str:
    """ファイルを自動バックアップの状態に巻き戻す。「元に戻して」「さっきの編集を取り消して」に使う。
    versionを省略すると「最後にファイルを変更した操作(AIの1回の依頼など)の直前」に戻る。
    特定の時点に戻すときは list_file_versions で確認したIDを version に指定する。
    巻き戻し前の状態も自動バックアップされるため、巻き戻し自体も取り消せる。"""
    try:
        used = history.restore(filename, version or None)
    except FileNotFoundError as e:
        return f"エラー: {e}"
    return (
        f"{filename} を {used.timestamp.strftime('%Y-%m-%d %H:%M:%S')} 時点"
        f"({_describe_label(used.label)})の状態に戻しました"
    )


HISTORY_TOOLS = [list_file_versions, restore_file]
