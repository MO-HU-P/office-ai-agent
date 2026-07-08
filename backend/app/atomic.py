"""ファイルのアトミック保存。

保存中にプロセスが落ちても、元ファイルが壊れないようにする。
一時ファイル(先頭ドットで隠しファイル扱い)へ書き込み、成功したら os.replace で
本来のパスへ差し替える。os.replace は同一ファイルシステム内ではアトミックに動作するため、
「途中まで書けた壊れたファイル」が本来のパスに残ることがない。

さらに、既存ファイルを上書きする場合は差し替え直前に services.history へ
「編集前の状態」を自動バックアップする。全ツール・アップロードがここを通るため、
この1箇所のフックで「差分確認」「巻き戻し」が全操作に効く。
"""
import os
import uuid
from pathlib import Path
from typing import Callable

from .services import history


def atomic_save(save_to: Callable[[str], None], final_path) -> None:
    """save_to(path) を一時ファイルに対して実行し、成功したら最終パスへ差し替える。

    save_to には doc.save / wb.save / prs.save のような「パスを1つ受け取る保存関数」を渡す。
    """
    final = Path(final_path)
    # 一時ファイルは同一ディレクトリに置く(os.replace のアトミック性は同一FS内が条件)。
    # 先頭ドットで一覧・変更検知から除外される。
    tmp = final.with_name(f".{final.name}.{uuid.uuid4().hex}.tmp")
    try:
        save_to(str(tmp))
        if final.exists():
            history.record_before_change(final)
        os.replace(str(tmp), str(final))
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
