"""LLMプロバイダーの集約。

Ollamaが主役(唯一のゼロ設定・クレカ不要の入口)だが、APIキーを持つユーザー向けに
OpenAI等の外部プロバイダーも同じ口から呼べるようにする。LLM生成(build_chat_model)・
reasoningの解釈・vision判定・モデル候補をここに集約し、下流(agent/loop.py)は
プロバイダーを意識せず LangChain の BaseChatModel を受け取る。

新しいプロバイダーを足すときは _build_xxx を書いて _BUILDERS に登録し、
config.VALID_PROVIDERS と設定UIに選択肢を追加する。
"""
import logging
import re
from typing import Any

from langchain_core.language_models.chat_models import BaseChatModel

from .. import config
from ..services import model_admin

logger = logging.getLogger(__name__)

# OpenAIの推奨モデル候補は config.toml の llm.openai_models で管理する(config.OPENAI_PRESET_MODELS)。
# モデルは順次廃止されるため、コードに埋め込まずデータ化している。ここに無い名前も設定UIから
# 自由入力でき、そちらは settings.json に保存される(list_preset_models 参照)。

# 推論モデル(o系・gpt-5系)の判定。temperature非対応で、深さは reasoning_effort で指定。
# gpt-4o系・gpt-4.1系など従来型は temperature を受け付ける(この正規表現に含めない)。
_OPENAI_REASONING_RE = re.compile(r"^(?:o\d|gpt-5)")


def _build_ollama(s: config.LLMSettings) -> BaseChatModel:
    from langchain_ollama import ChatOllama

    kwargs: dict[str, Any] = dict(model=s.model, base_url=s.base_url, temperature=0.1)
    if s.mode == "cloud":
        kwargs["client_kwargs"] = {"headers": s.headers()}
    else:
        kwargs["num_ctx"] = s.num_ctx
    if s.reasoning in ("true", "false"):
        kwargs["reasoning"] = s.reasoning == "true"
    elif s.reasoning in ("low", "medium", "high"):
        # gpt-oss等のレベル対応モデル向け。boolean型のモデルには "auto"/"true"/"false" を使う
        kwargs["reasoning"] = s.reasoning
    return ChatOllama(**kwargs)


def _build_openai(s: config.LLMSettings) -> BaseChatModel:
    from langchain_openai import ChatOpenAI

    if not config.OPENAI_API_KEY:
        raise RuntimeError("OpenAIのAPIキーが未設定です (.env の OPENAI_API_KEY)")
    kwargs: dict[str, Any] = dict(model=s.model, api_key=config.OPENAI_API_KEY, streaming=True)
    if _OPENAI_REASONING_RE.match(s.model):
        # 推論モデル(o系・gpt-5系)は temperature を受け付けず、送ると400エラーになる。
        # 深さは reasoning_effort で指定する(reasoning="auto" のときは送らずモデル既定)。
        if s.reasoning in ("low", "medium", "high"):
            kwargs["reasoning_effort"] = s.reasoning
    else:
        # gpt-4o系・gpt-4.1系など従来型モデルのみ temperature を送る。
        kwargs["temperature"] = 0.1
    return ChatOpenAI(**kwargs)


_BUILDERS = {"ollama": _build_ollama, "openai": _build_openai}


def build_chat_model(s: config.LLMSettings) -> BaseChatModel:
    """現在の設定に対応する LangChain チャットモデルを生成する。
    依頼のたびに呼ばれるため、設定UIでの変更が再起動なしで反映される。"""
    return _BUILDERS.get(s.provider, _build_ollama)(s)


async def supports_vision(s: config.LLMSettings) -> bool:
    """選択中モデルが画像入力(vision)に対応しているか。
    ollamaは /api/show のcapabilitiesで動的判定する。
    openaiは現行のチャットモデルがすべてテキスト+画像入力に対応するため常にTrue
    (接頭辞リストで判定するとgpt-6等の新モデルで即陳腐化し、対応済みモデルなのに
    自己レビュー描画が無効化されてしまうため、リストは持たない)。テキスト専用モデルを
    自由入力欄で指定した場合のみ、画像添付時にAPIエラーになる点は許容する。"""
    if s.provider == "openai":
        return True
    return await model_admin.model_supports_vision(s.mode, s.model)


def list_preset_models(provider: str) -> list[dict]:
    """設定UIのモデル一覧用。一覧APIを持たないプロバイダーの推奨候補を返す
    ([{name, size, vision}] 形式。sizeはダウンロード概念が無いのでNone)。
    OpenAIの候補は config.toml(llm.openai_models)由来。現行OpenAIのチャットモデルは
    すべて画像入力対応のため vision=True 固定。"""
    if provider == "openai":
        return [{"name": name, "size": None, "vision": True} for name in config.OPENAI_PRESET_MODELS]
    return []
