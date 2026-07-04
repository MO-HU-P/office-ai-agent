"""ファイル管理・Python実行ツール"""
import subprocess
import sys
from datetime import datetime

from langchain_core.tools import tool

from ...config import WORKSPACE_DIR, resolve_workspace_path


@tool
def list_files() -> str:
    """ワークスペース内のファイル一覧を返す(ファイル名・サイズ・更新日時)。"""
    entries = []
    for p in sorted(WORKSPACE_DIR.iterdir()):
        if p.is_file() and not p.name.startswith("."):
            mtime = datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
            entries.append(f"{p.name}  ({p.stat().st_size:,} bytes, {mtime})")
    return "\n".join(entries) if entries else "(ワークスペースは空です)"


@tool
def delete_file(filename: str) -> str:
    """ワークスペース内のファイルを削除する。ユーザーから明示的に削除を依頼された場合のみ使うこと。"""
    path = resolve_workspace_path(filename, must_exist=True)
    path.unlink()
    return f"{filename} を削除しました"


@tool
def run_python(code: str) -> str:
    """Pythonコードを実行し、標準出力を返す(データ分析・計算・CSV処理などに使う)。
    pandas, openpyxl, numpy が利用可能。カレントディレクトリはワークスペースなので、
    ファイルは相対パス(例: "data.csv")でそのまま読み書きできる。
    結果は必ずprint()で出力すること。Excelの集計値を計算してから書き込む用途にも使える。"""
    result = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        timeout=60,
        cwd=str(WORKSPACE_DIR),
    )
    out = result.stdout.strip()
    err = result.stderr.strip()
    if result.returncode != 0:
        return f"エラー (exit {result.returncode}):\n{err[-2000:]}"
    response = out[-4000:] if out else "(出力なし — 結果を見たい場合はprint()を使ってください)"
    if err:
        response += f"\n[stderr] {err[-500:]}"
    return response


FILE_TOOLS = [list_files, delete_file, run_python]
