"""エージェント本体: ChatOllama + ツールによるReActループ。

LangGraphのprebuiltではなく素のループを実装している。ローカルLLMは挙動の揺れが
大きく、ストリーミング・イベント発行・エラー回復を細かく制御したいため。
"""
import asyncio
import base64
import json
import logging
from pathlib import Path
from typing import Any, Awaitable, Callable

from langchain_core.messages import AIMessage, BaseMessage, HumanMessage, SystemMessage, ToolMessage
from langchain_ollama import ChatOllama
from pydantic import ValidationError

from .. import config
from ..services.model_admin import model_supports_vision
from ..services.workspace_watch import snapshot_workspace, diff_snapshots
from .tools.check_tools import CHECK_TOOLS
from .tools.excel_tools import EXCEL_TOOLS
from .tools.file_tools import FILE_TOOLS
from .tools.ppt_tools import PPT_TOOLS
from .tools.render_tools import IMAGE_RESULT_PREFIX, RENDER_TOOLS
from .tools.template_tools import TEMPLATE_TOOLS
from .tools.word_tools import WORD_TOOLS

logger = logging.getLogger(__name__)

ALL_TOOLS = FILE_TOOLS + EXCEL_TOOLS + WORD_TOOLS + PPT_TOOLS + TEMPLATE_TOOLS + CHECK_TOOLS
TOOL_MAP = {t.name: t for t in ALL_TOOLS}

SYSTEM_PROMPT = """あなたはOfficeファイル(Word/Excel/PowerPoint)を操作するAIアシスタントです。
ユーザーのワークスペース内のファイルをツールで直接読み書きできます。

ルール:
- ファイルを変更する依頼には、説明だけで終わらせず必ずツールを呼び出して実際に実行する。
- 既存ファイルを編集する前には、まず read 系ツール(word_read / excel_read / ppt_read)で現状を確認する。
- ファイル名には必ず拡張子(.docx / .xlsx / .pptx)を付ける。日本語のファイル名も使用可能。
- Excelで集計・統計などの計算結果を書く場合は、run_python で計算してから値を書き込むか、Excel数式("=SUM(B2:B10)"など)を書き込む。複雑な計算はrun_pythonを優先する。
- 統計解析(検定・回帰など)やグラフは run_python で scipy / seaborn を使う。グラフはPNGに保存して python-docx / python-pptx / openpyxl で文書に貼り込み、貼り込んだら一時PNGを削除する。Excel上で編集できるグラフを求められたときだけ openpyxl.chart を使う。PowerPointに入れる単純な棒・折れ線・円グラフは ppt_add_chart を優先する。
- 「〜の行」「〜と書いてある段落」など条件に合う場所を探すときは、全体を読まず excel_query / word_find で番号を特定してから編集・書式設定する。
- 「この文書と同じ体裁で」と言われたら、word_dump_style でお手本の体裁を調べ、word_apply_style で新しい文書に適用する。
- 表データはexcel_write_rowsで一括書き込みし、ヘッダー行にはexcel_formatで太字と背景色(#1a73e8の背景+白文字など)を付けて見やすくする。
- 複数の段落・スライドを直すときは word_batch_edit / ppt_batch_edit で一括修正する(1件ずつ繰り返さない)。
- 長い文書・大きなブックを読むときは、まず mode="outline"(Excelは mode="summary")で全体構造を把握する。
- 作業が終わったら、何をしたかを簡潔に(2〜4文で)日本語で報告する。長い前置きや箇条書きの乱用はしない。
- ツールがエラーを返したら、原因を考えて引数を修正し再試行する。同じ失敗を3回以上繰り返さない。
"""

# ビジョン対応モデルのときだけSYSTEM_PROMPTに追記するルール
VISION_RULE = """- あなたは画像を見ることができる。PowerPointやWordを作成・編集したら、render_page で主要なページを画像化して
文字のはみ出し・図形の重なり・不自然なレイアウトがないか自分の目で確認し、問題があれば修正してもう一度確認する。
"""

EmitFn = Callable[[dict[str, Any]], Awaitable[None]]


def _is_model_not_found(e: Exception) -> bool:
    """「モデルが存在しない」エラーか判定する(提供終了・入力ミス時。リトライ無意味)。"""
    if getattr(e, "status_code", None) == 404:
        return True
    text = str(e).lower()
    return "model" in text and "not found" in text


def _is_plan_restricted(e: Exception) -> bool:
    """「現在のOllamaプランでは使えないモデル」エラーか判定する(リトライ無意味)。
    Ollama Cloudの無料プランで有料限定モデルを選ぶと、403で
    "this model requires a subscription, upgrade for access" が返る。"""
    if getattr(e, "status_code", None) == 403:
        return True
    return "requires a subscription" in str(e).lower()


def build_llm() -> ChatOllama:
    # 依頼のたびに現在の設定を読むため、設定UIでの変更が再起動なしで反映される
    s = config.get_settings()
    kwargs: dict[str, Any] = dict(
        model=s.model,
        base_url=s.base_url,
        temperature=0.1,
    )
    if s.mode == "cloud":
        kwargs["client_kwargs"] = {"headers": s.headers()}
    else:
        kwargs["num_ctx"] = s.num_ctx
    if s.reasoning in ("true", "false"):
        kwargs["reasoning"] = s.reasoning == "true"
    elif s.reasoning in ("low", "medium", "high"):
        # gpt-oss等のレベル対応モデル向け。boolean のみ対応のモデルには "auto"/"true"/"false" を使う
        kwargs["reasoning"] = s.reasoning
    return ChatOllama(**kwargs)


class ThinkFilter:
    """コンテンツ中の <think>...</think> ブロックをストリームから除去する保険。"""

    def __init__(self):
        self._in_think = False
        self._buf = ""

    def feed(self, text: str) -> str:
        self._buf += text
        out = []
        while self._buf:
            if self._in_think:
                end = self._buf.find("</think>")
                if end == -1:
                    self._buf = self._buf[-8:]  # タグ跨ぎ検出用に末尾のみ保持
                    break
                self._buf = self._buf[end + len("</think>"):]
                self._in_think = False
            else:
                start = self._buf.find("<think>")
                if start == -1:
                    # 部分タグの可能性がある末尾は保留する
                    safe_len = len(self._buf)
                    for k in range(1, min(7, len(self._buf)) + 1):
                        if "<think>".startswith(self._buf[-k:]):
                            safe_len = len(self._buf) - k
                            break
                    out.append(self._buf[:safe_len])
                    self._buf = self._buf[safe_len:]
                    break
                out.append(self._buf[:start])
                self._buf = self._buf[start + len("<think>"):]
                self._in_think = True
        return "".join(out)


def _shorten(value: Any, limit: int = 300) -> str:
    text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False, default=str)
    return text if len(text) <= limit else text[:limit] + "…"


# pydanticのエラー種別(bool_parsing等)の先頭部分 → 期待する値の日本語表現
_TYPE_HINTS = {
    "bool": "true または false",
    "int": "整数",
    "float": "数値",
    "dict": "辞書({\"キー\": 値} の形式)",
    "model": "辞書({\"キー\": 値} の形式)",
    "list": "リスト([...] の形式)",
    "string": "文字列",
}


def _format_validation_error(name: str, e: ValidationError) -> str:
    """引数の型エラーを、LLMが引数を直してリトライしやすい平易な日本語にする。
    pydanticの英語メッセージをそのまま返すと、詳細を開いたユーザーにも読めない。"""
    hints = []
    for err in e.errors():
        field = ".".join(str(p) for p in err.get("loc", ())) or "(不明な引数)"
        etype = err.get("type", "")
        if etype == "missing":
            hints.append(f"{field}: 必須の引数が指定されていません")
            continue
        expect = _TYPE_HINTS.get(etype.split("_")[0])
        if expect:
            hints.append(f"{field}: {expect} を指定してください(使わない引数は空文字にせず省略する)")
        else:
            hints.append(f"{field}: 値の形式が正しくありません")
    return f"エラー: ツール「{name}」の引数が正しくありません。\n" + "\n".join(f"- {h}" for h in hints)


def _image_message(png_paths: list[str]) -> HumanMessage:
    """render_pageの結果画像をLLMへ渡すメッセージを作る。
    Ollamaのtoolロールは画像を運べないため、HumanMessageとして注入する。"""
    content: list[dict[str, Any]] = [{
        "type": "text",
        "text": "render_pageの結果画像です(システムが自動添付)。はみ出し・重なり・不自然なレイアウトがないか確認してください。",
    }]
    for p in png_paths:
        b64 = base64.b64encode(Path(p).read_bytes()).decode()
        content.append({"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}})
    return HumanMessage(content=content)


async def run_agent(user_message: str, history: list[BaseMessage], emit: EmitFn) -> list[BaseMessage]:
    """1ターン分のエージェント実行。更新後の履歴を返す。"""
    s = config.get_settings()
    # 画像を見られるモデルのときだけ render_page を有効化する
    # (テキスト専用モデルに画像入りメッセージを送るとエラーになるため)
    vision = await model_supports_vision(s.mode, s.model)
    tools = ALL_TOOLS + RENDER_TOOLS if vision else ALL_TOOLS
    tool_map = {t.name: t for t in tools}
    system_prompt = (SYSTEM_PROMPT + VISION_RULE) if vision else SYSTEM_PROMPT
    llm = build_llm().bind_tools(tools)
    messages: list[BaseMessage] = [SystemMessage(system_prompt), *history, HumanMessage(user_message)]

    for _step in range(config.MAX_AGENT_STEPS):
        think_filter = ThinkFilter()
        gathered = None
        failed = False
        # 一時的なサーバーエラー(Ollama Cloudの500など)に備え、
        # トークンをまだ出力していない段階での失敗のみリトライする
        for attempt in range(3):
            emitted = False
            gathered = None
            try:
                async for chunk in llm.astream(messages):
                    if isinstance(chunk.content, str) and chunk.content:
                        visible = think_filter.feed(chunk.content)
                        if visible:
                            await emit({"type": "token", "content": visible})
                            emitted = True
                    gathered = chunk if gathered is None else gathered + chunk
                break
            except Exception as e:
                if _is_model_not_found(e):
                    model = config.get_settings().model
                    logger.error("モデルが見つかりません (提供終了または入力ミス): %s", model)
                    await emit({
                        "type": "error",
                        "message": f"AIモデル「{model}」が見つかりません。提供終了した可能性があります。"
                                   "右上の設定（歯車アイコン）から別のモデルを選んでください。",
                    })
                    failed = True
                    break
                if _is_plan_restricted(e):
                    model = config.get_settings().model
                    logger.error("現在のOllamaプランでは利用できないモデルです: %s", model)
                    await emit({
                        "type": "error",
                        "message": f"AIモデル「{model}」は、お使いのOllamaアカウントのプラン（無料プランなど）では利用できません。"
                                   "右上の設定（歯車アイコン）から別のモデルに切り替えるか、"
                                   "Ollamaのプランのアップグレード（ollama.com）をご検討ください。",
                    })
                    failed = True
                    break
                if emitted or attempt == 2:
                    logger.exception("LLM呼び出しに失敗")
                    await emit({"type": "error", "message": "AIモデルの呼び出しに失敗しました。しばらくして再度お試しください。"})
                    failed = True
                    break
                logger.warning("LLM呼び出しに失敗、リトライします (%d/2): %s", attempt + 1, type(e).__name__)
                await asyncio.sleep(2 * (attempt + 1))
        if failed:
            break

        if gathered is None:
            await emit({"type": "error", "message": "LLMから応答がありませんでした"})
            break

        ai_msg = AIMessage(
            content=gathered.content if isinstance(gathered.content, str) else "",
            tool_calls=gathered.tool_calls or [],
        )
        messages.append(ai_msg)

        if not ai_msg.tool_calls:
            break  # 最終回答

        pending_images: list[str] = []
        for tc in ai_msg.tool_calls:
            name, args, call_id = tc["name"], tc.get("args") or {}, tc.get("id") or ""
            await emit({"type": "tool_start", "name": name, "args": _shorten(args, 500)})
            before = snapshot_workspace()
            tool_fn = tool_map.get(name)
            if tool_fn is None:
                result = f"エラー: ツール「{name}」は存在しません。利用可能: {', '.join(tool_map)}"
            else:
                try:
                    result = await asyncio.to_thread(tool_fn.invoke, args)
                except ValidationError as e:
                    result = _format_validation_error(name, e)
                except Exception as e:
                    result = f"エラー: {type(e).__name__}: {e}"
            # render_pageの結果は「画像パス|結果文」のマーカー形式。画像は後でHumanMessageとして注入する
            if isinstance(result, str) and result.startswith(IMAGE_RESULT_PREFIX):
                img_path, result = result[len(IMAGE_RESULT_PREFIX):].split("|", 1)
                pending_images.append(img_path)
            changed, deleted = diff_snapshots(before, snapshot_workspace())
            is_error = isinstance(result, str) and result.startswith("エラー")
            await emit({
                "type": "tool_end",
                "name": name,
                "ok": not is_error,
                "result": _shorten(result),
            })
            for fname in changed:
                await emit({"type": "doc_updated", "filename": fname})
            for fname in deleted:
                await emit({"type": "doc_deleted", "filename": fname})
            messages.append(ToolMessage(content=str(result), tool_call_id=call_id, name=name))
        if pending_images:
            # ToolMessage列の直後にまとめて注入する(列の途中に挟むとAPIエラーになるため)
            try:
                messages.append(_image_message(pending_images))
            except OSError:
                logger.exception("画像の読み込みに失敗")
                messages.append(HumanMessage("(画像の読み込みに失敗しました。テキスト情報だけで作業を続けてください)"))
    else:
        await emit({"type": "error", "message": f"ステップ上限({config.MAX_AGENT_STEPS})に達したため中断しました"})

    # システムプロンプトを除いた履歴を返し、直近のみ保持する
    new_history = messages[1:]
    if len(new_history) > config.MAX_HISTORY_MESSAGES:
        new_history = new_history[-config.MAX_HISTORY_MESSAGES:]
        # 先頭が ToolMessage だと次ターンでAPIエラーになり得るため取り除く
        while new_history and isinstance(new_history[0], ToolMessage):
            new_history.pop(0)
    # 注入した画像(base64)は巨大なので、次ターン以降の履歴ではプレースホルダに置き換える
    # (ユーザー入力のHumanMessageは常にstrなので、contentがリストのものは注入分だけ)
    return [
        HumanMessage("(このターンでrender_pageの画像を確認済み。画像データは履歴から省略)")
        if isinstance(m, HumanMessage) and isinstance(m.content, list) else m
        for m in new_history
    ]
