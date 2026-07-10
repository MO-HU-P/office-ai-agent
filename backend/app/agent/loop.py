"""エージェント本体: LLM(providersが出し分け) + ツールによるReActループ。

LangGraphのprebuiltではなく素のループを実装している。ローカルLLMは挙動の揺れが
大きく、ストリーミング・イベント発行・エラー回復を細かく制御したいため。
LLMの生成はプロバイダー非依存(agent/providers.py)に集約している。
"""
import asyncio
import base64
import json
import logging
import re
from pathlib import Path
from typing import Any, Awaitable, Callable

from langchain_core.messages import AIMessage, BaseMessage, HumanMessage, SystemMessage, ToolMessage
from pydantic import ValidationError

from .. import config
from . import providers
from ..services import history as history_mod  # run_agentの引数history(会話履歴)と衝突するため別名
from ..services.workspace_watch import snapshot_workspace, diff_snapshots
from .tools.check_tools import CHECK_TOOLS
from .tools.excel_tools import EXCEL_TOOLS
from .tools.file_tools import FILE_TOOLS
from .tools.history_tools import HISTORY_TOOLS
from .tools.pii_tools import PII_TOOLS
from .tools.ppt_tools import PPT_TOOLS
from .tools.render_tools import IMAGE_RESULT_PREFIX, RENDER_TOOLS
from .tools.review_tools import REVIEW_TOOLS
from .tools.template_tools import TEMPLATE_TOOLS
from .tools.word_tools import WORD_TOOLS

logger = logging.getLogger(__name__)

ALL_TOOLS = (FILE_TOOLS + EXCEL_TOOLS + WORD_TOOLS + PPT_TOOLS + TEMPLATE_TOOLS
             + CHECK_TOOLS + REVIEW_TOOLS + PII_TOOLS + HISTORY_TOOLS)
TOOL_MAP = {t.name: t for t in ALL_TOOLS}

SYSTEM_PROMPT = """あなたはOfficeファイル(Word/Excel/PowerPoint)を操作するAIアシスタントです。
ユーザーのワークスペース内のファイルをツールで直接読み書きできます。

ルール:
- ファイルを変更する依頼には、説明だけで終わらせず必ずツールを呼び出して実際に実行する。
- 既存ファイルを編集する前には、まず read 系ツール(word_read / excel_read / ppt_read)で現状を確認する。
- ファイル名には必ず拡張子(.docx / .xlsx / .pptx)を付ける。日本語のファイル名も使用可能。
- Excelで集計・統計などの計算結果を書く場合は、run_python で計算してから値を書き込むか、Excel数式("=SUM(B2:B10)"など)を書き込む。複雑な計算はrun_pythonを優先する。
- 統計解析やグラフは run_python を使う。単発の検定は scipy、回帰分析の詳細レポート・分散分析(ANOVA)・ロジスティック回帰・GLM・時系列(ARIMA)は statsmodels("y ~ x1 + x2" のR風の式が使える)。結果を資料にするときはsummary()の生テキストを貼らず、主要な数値(決定係数・係数・p値・信頼区間)を表にして、意味を平易な日本語で説明する。グラフはPNGに保存して python-docx / python-pptx / openpyxl で文書に貼り込み、貼り込んだら一時PNGを削除する。Excel上で編集できるグラフを求められたときだけ openpyxl.chart を使う。PowerPointに入れる単純な棒・折れ線・円グラフは ppt_add_chart を優先する。
- 「〜の行」「〜と書いてある段落」など条件に合う場所を探すときは、全体を読まず excel_query / word_find で番号を特定してから編集・書式設定する。
- 「この文書と同じ体裁で」と言われたら、word_dump_style でお手本の体裁を調べ、word_apply_style で新しい文書に適用する。
- 表データはexcel_write_rowsで一括書き込みし、ヘッダー行にはexcel_formatで太字と背景色(#1a73e8の背景+白文字など)を付けて見やすくする。
- 計算結果など新しい表を既存データがあるシートに追加するときは、excel_read(mode="summary")で使用範囲を確認し、
  既存データの最終行より下(1行以上空ける)か新しいシートに書く。既存のセルを上書きして消さないこと。
  書き込み結果に「⚠️ 上書きしました」と出たら、意図した上書きか確認し、誤りなら restore_file で戻してやり直す。
- 複数の段落・スライドを直すときは word_batch_edit / ppt_batch_edit で一括修正する(1件ずつ繰り返さない)。
- 長い文書・大きなブックを読むときは、まず mode="outline"(Excelは mode="summary")で全体構造を把握する。
- 作業が終わったら、何をしたかを簡潔に(2〜4文で)日本語で報告する。長い前置きや箇条書きの乱用はしない。
- ツールがエラーを返したら、原因を考えて引数を修正し再試行する。同じ失敗を3回以上繰り返さない。
- ファイルの編集・削除の直前の状態は自動でバックアップされている。「元に戻して」「さっきの編集を取り消して」と
  言われたら restore_file を使う(省略時は直前の変更前に戻る)。どの時点に戻すか迷うときは list_file_versions で
  一覧を確認してから選ぶ。自分の編集が意図とずれていたと気づいたときも restore_file でやり直せる。
- メッセージ冒頭に「（対象箇所: …）」と付いていたら、ユーザーがプレビュー上でマウス選択した場所を指す。
  作業はその箇所に限定する。Excelのセル範囲(例: B2:D5)はそのまま excel_write_cells / excel_format の引数に使い、
  スライド番号は ppt_edit_slide 等に、Wordの「…」の文言は word_find で段落番号を特定してから編集に使う。
- 複数の工程がかかる依頼(例: データ読込→集計→グラフ→レポート)は、最初のツールを呼ぶ前に
  「進め方: ①…→②…→③…」と1行で示してから作業する(どこまで進んだか・どこで失敗したかを分かりやすくするため)。
- 複数のファイルにまたがる作業は1ファイルずつ順に処理し、途中のエラーで全体を止めない(失敗したファイルは
  スキップして次へ進む)。最後に「✅完了したファイル / ⚠️失敗したファイルと原因」を分けて報告し、
  失敗分だけを再依頼すればよいことを伝える。

校閲(レビュー)作業の進め方:
- 校正(誤字脱字・文法・不自然な言い回しの修正): まず read 系ツールで全体を読み、意味と書式を変えずに文章だけを直す。複数箇所は word_batch_edit / ppt_batch_edit でまとめて直し、最後に「どこをどう直したか」を数点にまとめて報告する。原文を残したい指示があれば copy_file でコピーしてからコピー側を直す。
- 変更履歴(見え消し)での提案が求められたら(「変更履歴で」「見え消しで」「提案として」等)、Wordでは word_edit_paragraph で確定させず word_suggest_edits を使う。元の文に取り消し線・提案文に下線が付き、ユーザーがWord上で承認/却下できる。1段落に二重には提案できない。
- コメント(吹き出し)でのレビューが求められたら、本文は変えず word_add_comments で指摘・確認事項を該当段落に付ける(「本文は直さずコメントだけ」「指摘して」等)。変更履歴とコメントはどちらもWord(.docx)専用。
- 要約: 対象を read で読み、要点がわかるように日本語でまとめてチャットで返す。「ファイルにまとめて」と言われたときだけ新しいWord文書を作る。
- 匿名化: anonymize_file を使う(自分で本文を書き換えない)。メール・電話番号・URL・郵便番号・マイナンバー・カード番号を、AIに送らずローカルで確実にマスクした「(匿名化)元の名前」コピーを作り、元ファイルは残す。氏名・住所・会社/学校名などパターンの無い情報は自動マスクできないので、その点を必ず添えて報告する。
- 比較・差分: 「AとBの違い」「前の版から何が変わったか」を聞かれたら doc_diff で差分を取り、記号(－/＋)はそのまま見せずに、変更点を平易な日本語で要約して報告する。

翻訳の進め方:
- 進め方の確認: 翻訳を頼まれたら、着手する前に一度だけ「原文を残して訳文を併記する」か「原文を訳文で上書きする」かをチャットで確認してから作業する(この確認のときだけツールを呼ばず先に質問してよい)。依頼文にどちらか(「上書きで」「原文を残して」「対訳で」等)が既に書かれていれば、確認せずそれに従う。まず read 系ツールで全体を読み、書式・レイアウト・数値・数式は変えず文章だけを訳文にする。固有名詞・専門用語の訳語を通して揃え、終わったら報告する。
- Word: word_read で全段落を読み、word_batch_edit で各段落を訳文に一括置換する(併記のときは各段落の直後に訳文の段落を挿入する)。見出し・箇条書きなどのスタイルは保つ。
- PowerPoint: ppt_read で各スライドを読み、ppt_batch_edit でタイトル・本文を訳す(併記のときは訳文をノートに入れる)。訳で文字数が変わり枠から溢れやすいので、render_page で溢れ・重なりを確認し、あれば直す。
- Excel: excel_read で読み、文字列セルだけを excel_write_cells で訳文に置き換える。数値・日付・数式("="で始まるセル)は翻訳も変更もしない(集計が壊れる)。併記のときは訳文を隣の新しい列に書く。訳すべきか迷うセル(商品コード等)は残し、その旨を報告する。
"""

# ビジョン対応モデルのときだけSYSTEM_PROMPTに追記するルール
VISION_RULE = """- あなたは画像を見ることができる。PowerPointやWordを作成・編集したら、render_page で主要なページを画像化して
文字のはみ出し・図形の重なり・不自然なレイアウトがないか自分の目で確認し、問題があれば修正してもう一度確認する。
"""

# PPTX関連の依頼のときだけSYSTEM_PROMPTに追記するデザインガイド。
# 常駐させないのは、Word/Excelだけの依頼でトークンを消費しないため
# (特にローカルモデルのコンテキストを圧迫しない)。
_DESIGN_GUIDE_PATH = Path(__file__).parent / "design_guide.md"
_PPTX_HINT_RE = re.compile(r"\.pptx|スライド|パワポ|パワーポイント|プレゼン|powerpoint", re.IGNORECASE)


def _load_design_guide() -> str:
    """design_guide.md を読む(都度読み。ファイル編集だけでガイドを調整できる)。"""
    try:
        return _DESIGN_GUIDE_PATH.read_text(encoding="utf-8")
    except OSError:
        logger.exception("design_guide.md の読み込みに失敗")
        return ""


def _wants_design_guide(user_message: str, history: list[BaseMessage]) -> bool:
    """PPTXに関わる依頼か判定する。「続きを作って」のような継続依頼にも効くよう、
    直近の履歴に .pptx が現れる場合も対象にする。"""
    if _PPTX_HINT_RE.search(user_message):
        return True
    return any(".pptx" in str(m.content).lower() for m in history[-8:])

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


def _describe_error(e: Exception) -> str:
    """例外を、HTTPステータス付きの短い文字列にする(ログでの原因切り分け用)。
    Ollama CloudのResponseErrorは .status_code と .error(サーバーメッセージ)を持つ。
    これを記録しておくと、429(レート制限)・500(容量不足)・502等を区別できる。
    ※シークレットは log_setup の RedactSecretsFilter が墨消しするので、そのまま出してよい。"""
    parts = [type(e).__name__]
    status = getattr(e, "status_code", None)
    if status is not None:
        parts.append(f"status={status}")
    # ollamaのResponseErrorはサーバー由来メッセージを .error に持つ(無ければ str(e))
    msg = getattr(e, "error", None) or str(e)
    if msg:
        parts.append(_shorten(msg, 200))
    return " ".join(parts)


def _chunk_text(content: Any) -> str:
    """ストリームチャンクの content から可視テキストだけを取り出す。
    Ollama/OpenAIは文字列で届くが、Gemini 3系はコンテンツブロックのリスト
    ([{'type':'text','text':...}] や thinking ブロック等)で届くため、
    text ブロックのみを連結する(thinking等の非表示ブロックは可視化しない)。"""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for block in content:
            if isinstance(block, str):
                parts.append(block)
            elif isinstance(block, dict) and block.get("type") == "text":
                parts.append(block.get("text") or "")
        return "".join(parts)
    return ""


def build_llm():
    """現在の設定に対応するLLMを生成する(providersがOllama/OpenAI等を出し分ける)。
    依頼のたびに現在の設定を読むため、設定UIでの変更が再起動なしで反映される。"""
    return providers.build_chat_model(config.get_settings())


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


_TOOL_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def _drop_broken_tool_calls(tool_calls: list[dict]) -> list[dict]:
    """ツール名が識別子の形をしていない呼び出しを取り除く(Ollama Cloudが不安定なとき、
    「...」等の壊れた名前の呼び出しが正常な呼び出しに混ざって届くことがある)。
    正常な呼び出しが1つも残らない場合は元のまま返し、通常のエラー応答でモデルに修正を促す。"""
    kept = [tc for tc in tool_calls if _TOOL_NAME_RE.match(tc.get("name") or "")]
    if kept and len(kept) < len(tool_calls):
        dropped = [tc.get("name") or "(無名)" for tc in tool_calls if tc not in kept]
        logger.warning("壊れたツール呼び出しを無視します: %s", dropped)
    return kept or tool_calls


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
    vision = await providers.supports_vision(s)
    tools = ALL_TOOLS + RENDER_TOOLS if vision else ALL_TOOLS
    tool_map = {t.name: t for t in tools}
    system_prompt = (SYSTEM_PROMPT + VISION_RULE) if vision else SYSTEM_PROMPT
    if _wants_design_guide(user_message, history):
        guide = _load_design_guide()
        if guide:
            system_prompt += "\n\n" + guide
    llm = build_llm().bind_tools(tools)
    messages: list[BaseMessage] = [SystemMessage(system_prompt), *history, HumanMessage(user_message)]
    # この依頼(ターン)中のバックアップをひとまとめにする。「元に戻して」「変更箇所の表示」は
    # ターン単位(=依頼の直前の状態)で働く
    history_mod.begin_turn()

    for _step in range(config.MAX_AGENT_STEPS):
        gathered = None
        failed = False
        # 一時的なサーバーエラー(Ollama Cloudの500など)に備え、
        # トークンをまだ出力していない段階での失敗のみリトライする
        for attempt in range(config.LLM_MAX_ATTEMPTS):
            emitted = False
            gathered = None
            think_filter = ThinkFilter()  # リトライで最初から流し直すため、途中状態を持ち越さない
            try:
                # Ollama Cloudはまれに応答を返さないまま無言で止まる/推論だけ延々流し続けるため、
                # 「次のチャンクが来ない(idle)」と「応答全体が長すぎる(step)」の両方で打ち切りリトライに回す
                stream = aiter(llm.astream(messages))
                deadline = asyncio.get_running_loop().time() + config.LLM_STEP_TIMEOUT
                while True:
                    remaining = deadline - asyncio.get_running_loop().time()
                    if remaining <= 0:
                        raise TimeoutError(f"LLM応答が{config.LLM_STEP_TIMEOUT:.0f}秒を超えました")
                    try:
                        chunk = await asyncio.wait_for(anext(stream), timeout=min(config.LLM_IDLE_TIMEOUT, remaining))
                    except StopAsyncIteration:
                        break
                    text = _chunk_text(chunk.content)
                    if text:
                        visible = think_filter.feed(text)
                        if visible:
                            await emit({"type": "token", "content": visible})
                            emitted = True
                    gathered = chunk if gathered is None else gathered + chunk
                break
            except Exception as e:
                if _is_model_not_found(e):
                    s = config.get_settings()
                    model = s.model
                    logger.error("モデルが見つかりません (提供終了または入力ミス): %s", model)
                    if s.provider != "ollama":
                        # OpenAI/Gemini等はモデル名を自由入力できるため、まず入力ミスを疑ってもらう
                        message = (f"AIモデル「{model}」が見つかりません。モデル名が正しいか確認して、"
                                   "右上の設定（歯車アイコン）から選び直してください。")
                    else:
                        message = (f"AIモデル「{model}」が見つかりません。提供終了した可能性があります。"
                                   "右上の設定（歯車アイコン）から別のモデルを選んでください。")
                    await emit({"type": "error", "message": message})
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
                if emitted or attempt == config.LLM_MAX_ATTEMPTS - 1:
                    logger.exception("LLM呼び出しに失敗 (%s)", _describe_error(e))
                    await emit({"type": "error", "message": "AIモデルの呼び出しに失敗しました。しばらくして再度お試しください。"})
                    failed = True
                    break
                # 500の波が続くことがあるため指数バックオフ(上限あり)で間隔を空けて撃ち直す
                backoff = min(2 * 2 ** attempt, config.LLM_RETRY_BACKOFF_CAP)
                logger.warning(
                    "LLM呼び出しに失敗、%.0f秒後にリトライします (%d/%d): %s",
                    backoff, attempt + 1, config.LLM_MAX_ATTEMPTS - 1, _describe_error(e),
                )
                await asyncio.sleep(backoff)
        if failed:
            break

        if gathered is None:
            await emit({"type": "error", "message": "LLMから応答がありませんでした"})
            break

        # contentは文字列(Ollama/OpenAI)またはコンテンツブロックのリスト(Gemini 3系)。
        # リストはそのまま保持する。潰すと最終回答が履歴から消え、Gemini 3の思考署名
        # (signatureブロック)も失われてツール呼び出しの多ターン継続が壊れるため。
        ai_content = gathered.content if isinstance(gathered.content, (str, list)) else ""
        ai_msg = AIMessage(
            content=ai_content,
            tool_calls=_drop_broken_tool_calls(gathered.tool_calls or []),
        )
        messages.append(ai_msg)

        if not ai_msg.tool_calls:
            break  # 最終回答

        pending_images: list[str] = []
        for tc in ai_msg.tool_calls:
            name, args, call_id = tc["name"], tc.get("args") or {}, tc.get("id") or ""
            await emit({"type": "tool_start", "name": name, "args": _shorten(args, 500)})
            before = snapshot_workspace()
            if name == "run_python":
                # run_pythonは(pandas等で)atomic_saveを通らずファイルを書き換えうるため、
                # 実行前に全ファイルをバックアップしておく(変更が無いファイルはスキップされ軽い)
                await asyncio.to_thread(history_mod.backup_all)
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
