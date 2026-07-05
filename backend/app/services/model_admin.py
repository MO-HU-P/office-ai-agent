"""Ollamaモデルの一覧・ダウンロード・削除。

設定UIの裏側として、Ollama HTTP APIを叩く薄いラッパー。
- 一覧はモードに応じてローカルOllama / Ollama Cloudへ問い合わせる。
- ダウンロード(pull)と削除はローカルOllamaのみが対象(Cloudはダウンロード不要)。
- クライアントへ返すのはモデル名とサイズのみ。認証情報やURL等の内部情報は返さない。
"""
import asyncio
import json
import logging
from typing import AsyncIterator

import httpx

from .. import config

logger = logging.getLogger(__name__)


class OllamaUnavailable(Exception):
    """Ollamaに接続できない(コンテナ停止・ネットワーク断など)。"""


def _endpoint(mode: str) -> tuple[str, dict[str, str]]:
    if mode == "cloud":
        headers = {"Authorization": f"Bearer {config.OLLAMA_API_KEY}"} if config.OLLAMA_API_KEY else {}
        return config.OLLAMA_CLOUD_URL, headers
    return config.OLLAMA_LOCAL_URL, {}


async def list_models(mode: str) -> list[dict]:
    """モデル一覧を [{name, size, vision}] で返す。sizeはバイト数(不明ならNone)。
    visionは画像入力対応か(設定UIの「画像対応」バッジ用。判定結果はキャッシュされるため、
    2回目以降の一覧取得では追加のAPI呼び出しは発生しない)。"""
    base, headers = _endpoint(mode)
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            res = await client.get(f"{base}/api/tags", headers=headers)
    except httpx.HTTPError as e:
        raise OllamaUnavailable(f"Ollamaに接続できません ({mode})") from e
    if res.status_code != 200:
        raise OllamaUnavailable(f"Ollamaがエラーを返しました ({mode}: HTTP {res.status_code})")
    models = res.json().get("models", [])
    infos = [(m["name"], m.get("size")) for m in models if m.get("name")]
    sem = asyncio.Semaphore(8)  # /api/show をモデル数ぶん叩くので同時数を抑える

    async def _vision(name: str) -> bool:
        async with sem:
            return await model_supports_vision(mode, name)

    visions = await asyncio.gather(*(_vision(name) for name, _ in infos))
    return [
        {"name": name, "size": size, "vision": vision}
        for (name, size), vision in zip(infos, visions)
    ]


_vision_cache: dict[tuple[str, str], bool] = {}


async def model_supports_vision(mode: str, model: str) -> bool:
    """モデルが画像入力(vision)に対応しているかを /api/show のcapabilitiesで判定する。
    判定できないとき(接続失敗等)はFalseを返す(画像ツールを無効化する安全側)。
    確定した結果だけをキャッシュし、依頼のたびのAPI呼び出しを避ける。"""
    key = (mode, model)
    if key in _vision_cache:
        return _vision_cache[key]
    base, headers = _endpoint(mode)
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            res = await client.post(f"{base}/api/show", json={"model": model}, headers=headers)
    except httpx.HTTPError:
        return False
    if res.status_code != 200:
        return False
    try:
        ok = "vision" in res.json().get("capabilities", [])
    except (ValueError, AttributeError):
        return False
    _vision_cache[key] = ok
    return ok


async def stream_pull(name: str) -> AsyncIterator[str]:
    """ローカルOllamaへモデルをpullし、進捗をNDJSON行で逐次yieldする。

    Ollamaの応答から進捗表示に必要なフィールドだけを通す(内部情報のパススルーを避ける)。
    """
    async with httpx.AsyncClient(timeout=None) as client:
        try:
            async with client.stream(
                "POST", f"{config.OLLAMA_LOCAL_URL}/api/pull", json={"model": name}
            ) as res:
                if res.status_code != 200:
                    yield json.dumps({"error": f"ダウンロードを開始できませんでした (HTTP {res.status_code})"}) + "\n"
                    return
                async for line in res.aiter_lines():
                    if not line:
                        continue
                    try:
                        data = json.loads(line)
                    except json.JSONDecodeError:
                        continue
                    if data.get("error"):
                        logger.warning("モデルpullエラー %s: %s", name, data["error"])
                        yield json.dumps({"error": "ダウンロードに失敗しました。モデル名を確認してください。"}) + "\n"
                        return
                    out = {"status": data.get("status", "")}
                    if isinstance(data.get("total"), int):
                        out["total"] = data["total"]
                    if isinstance(data.get("completed"), int):
                        out["completed"] = data["completed"]
                    yield json.dumps(out, ensure_ascii=False) + "\n"
        except httpx.HTTPError:
            logger.exception("モデルpull中に接続エラー: %s", name)
            yield json.dumps({"error": "ローカルOllamaに接続できません。"}) + "\n"


async def delete_model(name: str) -> None:
    """ローカルOllamaからモデルを削除する。存在しない場合はFileNotFoundError。"""
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            res = await client.request(
                "DELETE", f"{config.OLLAMA_LOCAL_URL}/api/delete", json={"model": name}
            )
    except httpx.HTTPError as e:
        raise OllamaUnavailable("ローカルOllamaに接続できません") from e
    if res.status_code == 404:
        raise FileNotFoundError(f"モデルが見つかりません: {name}")
    if res.status_code != 200:
        raise OllamaUnavailable(f"削除に失敗しました (HTTP {res.status_code})")
