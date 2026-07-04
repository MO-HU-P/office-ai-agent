"""ワークスペースの変更検出(ツール実行前後のスナップショット比較)"""
from ..config import WORKSPACE_DIR


def snapshot_workspace() -> dict[str, float]:
    snap = {}
    try:
        for p in WORKSPACE_DIR.iterdir():
            if p.is_file() and not p.name.startswith("."):
                snap[p.name] = p.stat().st_mtime
    except FileNotFoundError:
        pass
    return snap


def diff_snapshots(before: dict[str, float], after: dict[str, float]) -> tuple[list[str], list[str]]:
    """(新規作成または更新されたファイル名, 削除されたファイル名) を返す。"""
    changed = sorted(name for name, mtime in after.items() if before.get(name) != mtime)
    deleted = sorted(name for name in before if name not in after)
    return changed, deleted
