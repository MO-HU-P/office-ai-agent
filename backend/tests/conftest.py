"""テスト共通フィクスチャ。

ワークスペースをテストごとの一時ディレクトリに差し替え、
本物の workspace / .history に触れないようにする。
"""
import pytest

from app import config
from app.services import history


@pytest.fixture
def ws(tmp_path, monkeypatch):
    """一時ワークスペース。config.WORKSPACE_DIR を差し替えて返す。"""
    monkeypatch.setattr(config, "WORKSPACE_DIR", tmp_path)
    history.end_turn()  # 前のテストのターン状態を持ち越さない
    yield tmp_path
    history.end_turn()
