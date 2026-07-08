"""ワークスペースの自動バックアップ(世代管理)と巻き戻し。

AIや上書きアップロードがファイルを変更・削除する直前の状態を
workspace/.history/<ファイル名>/ に自動保存する。これにより
「AIの編集をあとから差分で確認する」「編集前の状態に戻す」ができる。

- 保存の入口は atomic_save(全ツール・アップロードが通る)と delete 系のみで、
  ツール側は何も意識しなくてよい。
- バックアップの粒度は「変更のたび」だが、巻き戻し・差分の既定の単位は
  「ターン(=ユーザーの1回の依頼)」。エージェント実行開始時に begin_turn() で
  ターン番号を進め、そのターン中の最初のバックアップ=編集前の状態、とみなす。
- ファイル名先頭が「.」のもの・ワークスペース外は対象外(履歴自身や一時ファイルを守る)。
"""
import logging
import os
import re
import shutil
import threading
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .. import config

logger = logging.getLogger(__name__)

# 1ファイルあたり保持する世代数。古いものから自動削除する(運用の手間を増やさない)。
MAX_VERSIONS_PER_FILE = 20

# バックアップの由来ラベル。tN=エージェントのターン、manual=アップロード等のUI操作、
# restore=巻き戻し実行時に退避した「戻す前」の状態。
_TURN_PREFIX = "t"
_MANUAL_LABEL = "manual"

_lock = threading.Lock()
_turn_counter = 0
_current_label = _MANUAL_LABEL

# バージョンID: "20260708-153012-123456__t3" (時刻__ラベル)。ファイル名に拡張子を足して保存する。
_VERSION_RE = re.compile(r"^(\d{8}-\d{6}-\d{6})__([A-Za-z0-9_-]+)$")


def _history_dir() -> Path:
    return config.WORKSPACE_DIR / ".history"


@dataclass(frozen=True)
class Version:
    """1つのバックアップ世代。idはバージョンID、pathは実体、labelはターン等の由来。"""
    id: str
    path: Path
    timestamp: datetime
    label: str
    size: int


def begin_turn() -> str:
    """エージェントの新しいターンを開始し、以後のバックアップにそのターン番号を付ける。"""
    global _turn_counter, _current_label
    with _lock:
        _turn_counter += 1
        _current_label = f"{_TURN_PREFIX}{_turn_counter}"
        return _current_label


def end_turn() -> None:
    """ターン終了。以後のバックアップ(UI操作等)は manual 扱いに戻す。"""
    global _current_label
    with _lock:
        _current_label = _MANUAL_LABEL


def current_label() -> str:
    with _lock:
        return _current_label


def _is_trackable(path: Path) -> bool:
    """バックアップ対象か(ワークスペース直下の通常ファイルで、隠しファイルでない)。"""
    try:
        rel = path.resolve().relative_to(config.WORKSPACE_DIR)
    except ValueError:
        return False
    return not any(part.startswith(".") for part in rel.parts) and path.is_file()


def _version_dir(filename: str) -> Path:
    return _history_dir() / filename


def record_before_change(path: Path, label: str | None = None) -> None:
    """ファイルが変更・削除される直前に呼び、現在の内容を1世代として保存する。

    直前のバックアップからファイルが変わっていなければ何もしない(重複を作らない)。
    バックアップの失敗で本来の保存処理を止めないよう、例外はログに留める。
    """
    try:
        if not _is_trackable(path):
            return
        newest = _newest_version(path.name)
        stat = path.stat()
        if newest is not None:
            nstat = newest.path.stat()
            # 内容が前回バックアップから変わっていない(サイズ+mtimeで判定)ならスキップ。
            # copy2がナノ秒までmtimeを保つため、ns精度で比較できる
            if nstat.st_size == stat.st_size and nstat.st_mtime_ns == stat.st_mtime_ns:
                return
        vdir = _version_dir(path.name)
        vdir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        vid = f"{ts}__{label or current_label()}"
        tmp = vdir / f".{uuid.uuid4().hex}.tmp"
        try:
            shutil.copy2(path, tmp)  # copy2でmtimeも保つ(スキップ判定に使う)
            os.replace(tmp, vdir / f"{vid}{path.suffix}")
        finally:
            tmp.unlink(missing_ok=True)
        _prune(path.name)
    except OSError:
        logger.exception("バックアップの保存に失敗: %s", path)


def backup_all(label: str | None = None) -> None:
    """ワークスペースの全対象ファイルをバックアップする(run_python など、
    どのファイルを書き換えるか事前に分からないツールの実行前に呼ぶ)。
    変更のないファイルは record_before_change 側でスキップされるため軽い。"""
    try:
        for p in config.WORKSPACE_DIR.iterdir():
            if p.is_file() and not p.name.startswith("."):
                record_before_change(p, label)
    except OSError:
        logger.exception("全体バックアップに失敗")


def list_versions(filename: str) -> list[Version]:
    """新しい順のバックアップ世代一覧。"""
    vdir = _version_dir(filename)
    if not vdir.is_dir():
        return []
    versions: list[Version] = []
    for p in vdir.iterdir():
        if not p.is_file() or p.name.startswith("."):
            continue
        m = _VERSION_RE.match(p.stem)
        if not m:
            continue
        ts_text, label = m.groups()
        try:
            ts = datetime.strptime(ts_text, "%Y%m%d-%H%M%S-%f")
        except ValueError:
            continue
        versions.append(Version(id=p.stem, path=p, timestamp=ts, label=label, size=p.stat().st_size))
    versions.sort(key=lambda v: v.id, reverse=True)
    return versions


def _newest_version(filename: str) -> Version | None:
    versions = list_versions(filename)
    return versions[0] if versions else None


def last_change_base(filename: str) -> Version | None:
    """「最後にこのファイルを変えた一連の操作」の直前の状態を返す(差分・巻き戻しの既定)。

    エージェントのターン(tN)は、同一ターン内の複数保存をひとまとめにし、
    そのターンで最初に取られたバックアップ=ターン開始前の状態を採用する。
    manual/restore はバックアップ1つを1操作として扱う。
    """
    versions = list_versions(filename)
    if not versions:
        return None
    newest = versions[0]
    if not newest.label.startswith(_TURN_PREFIX):
        return newest
    # 同じターンのうち最も古いもの(=ターン突入前の状態)
    same_turn = [v for v in versions if v.label == newest.label]
    return same_turn[-1]


def find_version(filename: str, version_id: str) -> Version | None:
    for v in list_versions(filename):
        if v.id == version_id:
            return v
    return None


def restore(filename: str, version_id: str | None = None) -> Version:
    """ファイルを指定世代(省略時は「最後の変更前」)に戻す。

    戻す直前の現在の状態も restore ラベルでバックアップするため、巻き戻し自体も取り消せる。
    戻り値は復元に使った世代。該当世代が無ければ FileNotFoundError。
    """
    target = find_version(filename, version_id) if version_id else last_change_base(filename)
    if target is None:
        raise FileNotFoundError(
            f"「{filename}」に戻せるバックアップがありません"
            + (f" (指定: {version_id})" if version_id else "")
        )
    dest = config.resolve_workspace_path(filename)
    if dest.exists():
        record_before_change(dest, label="restore")
    tmp = dest.with_name(f".{dest.name}.{uuid.uuid4().hex}.tmp")
    try:
        shutil.copy2(target.path, tmp)
        os.replace(tmp, dest)
    finally:
        tmp.unlink(missing_ok=True)
    logger.info("巻き戻し: %s ← %s", filename, target.id)
    return target


def _prune(filename: str) -> None:
    """古い世代を削除して MAX_VERSIONS_PER_FILE 世代までに保つ。"""
    for v in list_versions(filename)[MAX_VERSIONS_PER_FILE:]:
        try:
            v.path.unlink()
        except OSError:
            pass
