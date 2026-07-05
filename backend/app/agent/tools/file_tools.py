"""ファイル管理・Python実行ツール"""
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from langchain_core.tools import tool

from ...config import WORKSPACE_DIR, resolve_workspace_path

# グラフ描画の既定設定(ヘッドレス動作 + 日本語フォント)。LLMがコード内で設定しなくても効く
MATPLOTLIBRC = Path(__file__).with_name("matplotlibrc")


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
def copy_file(source: str, dest: str) -> str:
    """ワークスペース内のファイルを複製する。書式・レイアウトを含め完全に同じコピーができる。
    「元のファイルを残したまま編集したい」「バックアップを作って」と頼まれたら、まずこれでコピーし、
    コピーしたほうを編集する。読み取って書き写す方法は書式が失われるので使わないこと。"""
    src = resolve_workspace_path(source, must_exist=True)
    dst = resolve_workspace_path(dest)
    if dst.exists():
        return f"エラー: {dest} は既に存在します。別の名前を指定してください"
    shutil.copy2(src, dst)
    return f"{source} を {dest} にコピーしました"


@tool
def rename_file(old_name: str, new_name: str) -> str:
    """ワークスペース内のファイル名を変更する。中身はそのまま、名前だけが変わる。
    拡張子(.docx等)は変更できない。"""
    src = resolve_workspace_path(old_name, must_exist=True)
    dst = resolve_workspace_path(new_name)
    if dst.exists():
        return f"エラー: {new_name} は既に存在します。別の名前を指定してください"
    if src.suffix != dst.suffix:
        return f"エラー: 拡張子は変更できません({src.suffix} のままにしてください)"
    src.rename(dst)
    return f"{old_name} を {new_name} に変更しました"


@tool
def delete_file(filename: str) -> str:
    """ワークスペース内のファイルを削除する。ユーザーから明示的に削除を依頼された場合のみ使うこと。"""
    path = resolve_workspace_path(filename, must_exist=True)
    path.unlink()
    return f"{filename} を削除しました"


@tool
def run_python(code: str) -> str:
    """Pythonコードを実行し、標準出力を返す(データ分析・統計解析・グラフ作成・CSV処理などに使う)。
    利用可能: pandas / numpy / scipy(t検定・カイ二乗検定・分布フィットなどの基本統計) /
    statsmodels(統計モデリング: 回帰分析の詳細レポート summary()・分散分析 anova_lm・
    ロジスティック回帰・GLM・時系列 ARIMA。smf.ols("y ~ x1 + x2", data=df) のように
    R風の式でモデルを指定でき、決定係数・係数のp値・信頼区間まで一度に出せる) /
    seaborn・matplotlib(グラフ描画。日本語フォント設定済みなのでラベルは日本語でよい) /
    openpyxl・python-docx・python-pptx(グラフ画像の文書への貼り込みに使える)。
    カレントディレクトリはワークスペースなので、ファイルは相対パス(例: "data.csv")でそのまま読み書きできる。
    結果は必ずprint()で出力すること。Excelの集計値を計算してから書き込む用途にも使える。
    グラフは plt.savefig("グラフ.png") でPNG保存してから add_picture 等で文書に貼り込み、
    貼り込みが済んだ一時PNGは os.remove で削除する。"""
    # APIキー等のシークレットは実行コードから見えないようにする(生成コードが誤って文書へ書き出すのを防ぐ)
    env = {k: v for k, v in os.environ.items() if "KEY" not in k.upper() and "TOKEN" not in k.upper() and "SECRET" not in k.upper()}
    result = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        text=True,
        timeout=60,
        cwd=str(WORKSPACE_DIR),
        env={**env, "MPLBACKEND": "Agg", "MATPLOTLIBRC": str(MATPLOTLIBRC)},
    )
    out = result.stdout.strip()
    err = result.stderr.strip()
    if result.returncode != 0:
        return f"エラー (exit {result.returncode}):\n{err[-2000:]}"
    response = out[-4000:] if out else "(出力なし — 結果を見たい場合はprint()を使ってください)"
    if err:
        response += f"\n[stderr] {err[-500:]}"
    return response


FILE_TOOLS = [list_files, copy_file, rename_file, delete_file, run_python]
