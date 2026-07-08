"""自動バックアップ(services/history)と atomic_save フックのテスト。"""
import pytest

from app import config
from app.atomic import atomic_save
from app.services import history


def _write(path, text: str):
    """atomic_save経由でテキストを書く(本番のツール保存と同じ経路)。"""
    atomic_save(lambda p: open(p, "w", encoding="utf-8").write(text), path)


def test_new_file_has_no_backup(ws):
    _write(ws / "a.txt", "v1")
    assert history.list_versions("a.txt") == []


def test_overwrite_creates_backup_of_previous_content(ws):
    _write(ws / "a.txt", "v1")
    _write(ws / "a.txt", "v2")
    versions = history.list_versions("a.txt")
    assert len(versions) == 1
    assert versions[0].path.read_text(encoding="utf-8") == "v1"
    assert (ws / "a.txt").read_text(encoding="utf-8") == "v2"


def test_unchanged_file_is_not_backed_up_twice(ws):
    _write(ws / "a.txt", "v1")
    history.record_before_change(ws / "a.txt")
    history.record_before_change(ws / "a.txt")  # 変更なしの2回目はスキップ
    assert len(history.list_versions("a.txt")) == 1


def test_backup_all_skips_hidden_and_unchanged(ws):
    _write(ws / "a.txt", "v1")
    (ws / ".secret").write_text("x", encoding="utf-8")
    history.backup_all()
    history.backup_all()
    assert len(history.list_versions("a.txt")) == 1
    assert history.list_versions(".secret") == []


def test_last_change_base_groups_turn_saves(ws):
    _write(ws / "a.txt", "A")             # 作成(バックアップなし)
    history.begin_turn()
    _write(ws / "a.txt", "B")             # 編集前A がバックアップされる
    _write(ws / "a.txt", "C")             # 編集前B がバックアップされる
    history.end_turn()
    base = history.last_change_base("a.txt")
    # 同一ターン内の複数保存はまとめられ、ターン開始前(A)が基準になる
    assert base is not None
    assert base.path.read_text(encoding="utf-8") == "A"


def test_restore_returns_to_pre_turn_state_and_is_reversible(ws):
    _write(ws / "a.txt", "A")
    history.begin_turn()
    _write(ws / "a.txt", "B")
    _write(ws / "a.txt", "C")
    history.end_turn()

    history.restore("a.txt")  # 既定=最後の変更前(A)へ
    assert (ws / "a.txt").read_text(encoding="utf-8") == "A"
    # 巻き戻し前の状態(C)も退避されているので、もう一度restoreすると戻せる
    history.restore("a.txt")
    assert (ws / "a.txt").read_text(encoding="utf-8") == "C"


def test_restore_specific_version(ws):
    _write(ws / "a.txt", "A")
    history.begin_turn()
    _write(ws / "a.txt", "B")
    history.end_turn()
    history.begin_turn()
    _write(ws / "a.txt", "C")
    history.end_turn()
    # 一番古い世代(A)を指定して戻す
    oldest = history.list_versions("a.txt")[-1]
    assert oldest.path.read_text(encoding="utf-8") == "A"
    history.restore("a.txt", oldest.id)
    assert (ws / "a.txt").read_text(encoding="utf-8") == "A"


def test_restore_without_backup_raises(ws):
    _write(ws / "a.txt", "A")
    with pytest.raises(FileNotFoundError):
        history.restore("a.txt")
    with pytest.raises(FileNotFoundError):
        history.restore("a.txt", "20000101-000000-000000__t1")


def test_prune_keeps_latest_versions(ws):
    for i in range(history.MAX_VERSIONS_PER_FILE + 5):
        _write(ws / "a.txt", f"v{i}")
    versions = history.list_versions("a.txt")
    assert len(versions) == history.MAX_VERSIONS_PER_FILE
    # 最新のバックアップは直前の内容
    assert versions[0].path.read_text(encoding="utf-8") == f"v{history.MAX_VERSIONS_PER_FILE + 3}"


def test_delete_tool_backs_up_before_unlink(ws):
    from app.agent.tools.file_tools import delete_file

    _write(ws / "a.txt", "A")
    result = delete_file.invoke({"filename": "a.txt"})
    assert "削除しました" in result
    assert not (ws / "a.txt").exists()
    history.restore("a.txt")
    assert (ws / "a.txt").read_text(encoding="utf-8") == "A"


def test_hidden_paths_are_rejected(ws):
    with pytest.raises(ValueError):
        config.resolve_workspace_path(".history/a.txt/x.txt")
    with pytest.raises(ValueError):
        config.resolve_workspace_path(".env")


def test_restore_tool_and_list_versions_tool(ws):
    from app.agent.tools.history_tools import list_file_versions, restore_file

    _write(ws / "a.txt", "A")
    assert "まだありません" in list_file_versions.invoke({"filename": "a.txt"})

    history.begin_turn()
    _write(ws / "a.txt", "B")
    history.end_turn()

    listing = list_file_versions.invoke({"filename": "a.txt"})
    assert "AI編集前" in listing

    result = restore_file.invoke({"filename": "a.txt"})
    assert "戻しました" in result
    assert (ws / "a.txt").read_text(encoding="utf-8") == "A"

    assert "エラー" in restore_file.invoke({"filename": "nai.txt"})
