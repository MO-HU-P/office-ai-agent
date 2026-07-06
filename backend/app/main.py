import asyncio
import json
import logging
from contextlib import asynccontextmanager
from datetime import datetime
from pathlib import Path

import httpx
from fastapi import FastAPI, HTTPException, UploadFile, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from pydantic import BaseModel

from . import config
from .agent import providers
from .agent.loop import run_agent
from .atomic import atomic_save
from .log_setup import configure_logging
from .services import model_admin
from .services.preview import excel_preview, pptx_preview

configure_logging()
logger = logging.getLogger(__name__)

# バックグラウンドタスクの参照を保持する(参照ゼロのタスクはGCに途中で回収されうるため)
_background_tasks: set[asyncio.Task] = set()


async def _auto_pull_local_model(model: str):
    """localモード時、設定モデルが未ダウンロードならバックグラウンドで自動pullする"""
    try:
        async with httpx.AsyncClient(timeout=None) as client:
            for _ in range(30):  # ollama起動待ち(最大60秒)
                try:
                    res = await client.get(f"{config.OLLAMA_LOCAL_URL}/api/tags", timeout=2)
                    if res.status_code == 200:
                        break
                except Exception:
                    pass
                await asyncio.sleep(2)
            else:
                logger.warning("Ollamaに接続できないため自動pullをスキップします")
                return
            models = [m["name"] for m in res.json().get("models", [])]
            if any(m == model or m.split(":")[0] == model for m in models):
                logger.info("モデル %s は準備済みです", model)
                return
            logger.info("モデル %s をダウンロードします…", model)
            last_status = ""
            async with client.stream(
                "POST", f"{config.OLLAMA_LOCAL_URL}/api/pull", json={"model": model}
            ) as r:
                async for line in r.aiter_lines():
                    if not line:
                        continue
                    status = json.loads(line).get("status", "")
                    if status != last_status:
                        logger.info("pull %s: %s", model, status)
                        last_status = status
            logger.info("モデル %s の準備が完了しました", model)
    except Exception:
        logger.exception("モデルの自動pullに失敗しました")


@asynccontextmanager
async def lifespan(_app: FastAPI):
    s = config.get_settings()
    if s.provider == "ollama" and s.mode == "local":
        task = asyncio.create_task(_auto_pull_local_model(s.model))
        _background_tasks.add(task)
        task.add_done_callback(_background_tasks.discard)
    else:
        logger.info("LLM: provider=%s (model=%s)", s.provider, s.model)
    yield


app = FastAPI(title="Office AI Agent", lifespan=lifespan)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

FILE_TYPES = {".docx": "word", ".xlsx": "excel", ".pptx": "powerpoint", ".csv": "csv"}


@app.get("/api/health")
async def health():
    s = config.get_settings()
    if s.provider != "ollama":
        # 外部プロバイダー(openai等)は鍵の有無だけを見る(毎回の疎通pingで課金しない)
        key_missing = not _provider_key_configured(s.provider)
        return {
            "status": "ok",
            "provider": s.provider,
            "mode": s.mode,
            "backend_ok": not key_missing,
            "key_missing": key_missing,
            "model": s.model,
            "model_ready": not key_missing,
        }
    key_missing = s.mode == "cloud" and not config.OLLAMA_API_KEY
    ollama_ok = False
    models: list[str] = []
    if not key_missing:
        try:
            async with httpx.AsyncClient(timeout=5) as client:
                res = await client.get(f"{s.base_url}/api/tags", headers=s.headers())
                if res.status_code == 200:
                    ollama_ok = True
                    models = [m["name"] for m in res.json().get("models", [])]
        except Exception:
            pass
    model_found = any(m == s.model or m.startswith(s.model + ":") for m in models) or (
        ":" not in s.model and any(m.startswith(s.model) for m in models)
    )
    if s.mode == "cloud":
        # Cloudはダウンロード不要だが、選択中モデルが提供終了していないか一覧と突き合わせる
        model_ready = ollama_ok and model_found
    else:
        model_ready = model_found
    return {
        "status": "ok",
        "provider": "ollama",
        "mode": s.mode,
        "backend_ok": ollama_ok,
        "key_missing": key_missing,
        "model": s.model,
        "model_ready": model_ready,
    }


# ---------- 設定・モデル管理 (設定UIから利用) ----------
# セキュリティ方針: APIキーはこれらのAPIで受け取らず・返さず、.envでのみ管理する。


class SettingsUpdate(BaseModel):
    provider: str | None = None
    mode: str | None = None
    model_local: str | None = None
    model_cloud: str | None = None
    model_openai: str | None = None
    openai_custom_models: list[str] | None = None
    reasoning: str | None = None


class PullRequest(BaseModel):
    name: str


def _provider_key_configured(provider: str) -> bool:
    """そのプロバイダーのAPIキーが .env に設定済みか(真偽値のみ。キー本体は扱わない)。"""
    return {"openai": bool(config.OPENAI_API_KEY)}.get(provider, False)


def _settings_response(s: config.LLMSettings) -> dict:
    return {
        "provider": s.provider,
        "mode": s.mode,
        "model": s.model,
        "model_local": s.model_local,
        "model_cloud": s.model_cloud,
        "model_openai": s.model_openai,
        "openai_custom_models": list(s.openai_custom_models),
        "reasoning": s.reasoning,
        # キー本体は返さない。「設定済みかどうか」だけをUIに知らせる
        "cloud_key_configured": bool(config.OLLAMA_API_KEY),
        "openai_key_configured": bool(config.OPENAI_API_KEY),
    }


@app.get("/api/settings")
async def get_settings():
    return _settings_response(config.get_settings())


@app.put("/api/settings")
async def put_settings(update: SettingsUpdate):
    try:
        s = config.update_settings(update.model_dump(exclude_unset=True))
    except ValueError as e:
        raise HTTPException(400, str(e))
    return _settings_response(s)


@app.get("/api/models")
async def list_models(source: str | None = None):
    """指定ソース(local / cloud / openai)のモデル一覧を返す。
    省略時は現在の実効ソース。local/cloudはOllama、openaiは推奨候補の固定リスト。"""
    s = config.get_settings()
    source = (source or ("openai" if s.provider == "openai" else s.mode)).lower()
    # 外部プロバイダー(openai等): ダウンロード概念が無いため推奨候補を返す
    if source not in config.VALID_MODES:
        if source not in config.VALID_PROVIDERS:
            raise HTTPException(400, "source は local / cloud / openai のいずれかです")
        if not _provider_key_configured(source):
            return {"source": source, "models": [], "unavailable": f"{source} のAPIキーが未設定です (.env)"}
        return {"source": source, "models": providers.list_preset_models(source)}
    # Ollama (local / cloud)
    if source == "cloud" and not config.OLLAMA_API_KEY:
        return {"source": source, "models": [], "unavailable": "APIキーが未設定です (.env)"}
    try:
        models = await model_admin.list_models(source)
    except model_admin.OllamaUnavailable as e:
        return {"source": source, "models": [], "unavailable": str(e)}
    return {"source": source, "models": models}


@app.post("/api/models/pull")
async def pull_model(req: PullRequest):
    """ローカルOllamaへモデルをダウンロードし、進捗をNDJSONでストリームする。"""
    try:
        name = config.validate_model_name(req.name)
    except ValueError as e:
        raise HTTPException(400, str(e))
    return StreamingResponse(model_admin.stream_pull(name), media_type="application/x-ndjson")


@app.delete("/api/models/{name:path}")
async def delete_model(name: str):
    """ローカルOllamaからモデルを削除する(使用中のモデルは削除不可)。"""
    try:
        name = config.validate_model_name(name)
    except ValueError as e:
        raise HTTPException(400, str(e))
    s = config.get_settings()
    if name == s.model_local or name.split(":")[0] == s.model_local:
        raise HTTPException(400, "使用中のモデルは削除できません。先に別のモデルへ切り替えてください。")
    try:
        await model_admin.delete_model(name)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))
    except model_admin.OllamaUnavailable as e:
        raise HTTPException(502, str(e))
    logger.info("モデルを削除しました: %s", name)
    return {"deleted": name}


@app.get("/api/files")
async def list_files():
    files = []
    for p in sorted(config.WORKSPACE_DIR.iterdir()):
        if p.is_file() and not p.name.startswith("."):
            files.append({
                "name": p.name,
                "size": p.stat().st_size,
                "mtime": datetime.fromtimestamp(p.stat().st_mtime).isoformat(),
                "type": FILE_TYPES.get(p.suffix.lower(), "other"),
            })
    return files


@app.post("/api/files/upload")
async def upload_file(file: UploadFile):
    name = Path(file.filename or "upload").name
    dest = config.resolve_workspace_path(name)
    content = await file.read()
    # 書き込み途中で失敗しても壊れたファイルが残らないようにする
    atomic_save(lambda p: Path(p).write_bytes(content), dest)
    return {"name": name, "size": len(content)}


@app.get("/api/files/{filename}/raw")
async def get_raw(filename: str):
    path = _existing(filename)
    return FileResponse(str(path), filename=path.name)


@app.get("/api/files/{filename}/preview")
async def get_preview(filename: str):
    path = _existing(filename)
    suffix = path.suffix.lower()
    try:
        if suffix == ".xlsx":
            # 大きいブックの解析でイベントループ(他のAPI応答)を止めないようスレッドで実行
            return await asyncio.to_thread(excel_preview, path)
        if suffix == ".pptx":
            return await pptx_preview(path)
        if suffix == ".docx":
            return {"type": "docx"}  # フロント側でrawをdocx-preview描画
        if suffix == ".csv":
            text = path.read_text(encoding="utf-8", errors="replace")
            return {"type": "csv", "content": text[:200_000]}
    except Exception:
        logger.exception("プレビュー生成に失敗: %s", filename)
        return JSONResponse(status_code=500, content={"detail": "プレビューの生成に失敗しました"})
    return {"type": "unsupported"}


@app.delete("/api/files/{filename}")
async def delete_file(filename: str):
    path = _existing(filename)
    path.unlink()
    return {"deleted": filename}


@app.get("/api/preview_cache/{cache_name}/{image_name}")
async def preview_image(cache_name: str, image_name: str):
    path = (config.PREVIEW_CACHE_DIR / cache_name / image_name).resolve()
    if not path.is_relative_to(config.PREVIEW_CACHE_DIR) or not path.exists():
        raise HTTPException(404)
    return FileResponse(str(path), media_type="image/png")


def _existing(filename: str) -> Path:
    try:
        return config.resolve_workspace_path(filename, must_exist=True)
    except FileNotFoundError:
        raise HTTPException(404, f"ファイルが見つかりません: {filename}")
    except ValueError as e:
        raise HTTPException(400, str(e))


# 会話の記憶はプロセス全体で1つ持つ(このアプリはlocalhostのシングルユーザー前提)。
# 一時的な切断→再接続で記憶が消えないようにする。ページの再読み込み時は
# フロントが reset を送ってくるため、新しい会話として始まる。
_chat_history: list = []
# エージェント実行を直列化する(複数タブから同時に依頼されたとき、
# 2つのエージェントが同じファイルを同時編集して壊すのを防ぐ)
_agent_lock = asyncio.Lock()


@app.websocket("/ws")
async def websocket_endpoint(ws: WebSocket):
    await ws.accept()

    async def emit(event: dict):
        await ws.send_json(event)

    async def run_and_watch(content: str, history: list) -> list:
        """エージェントを実行しつつ切断を監視する。切断されたら実行を即キャンセルする
        (LLM応答待ちの間に切断されると、送信も受信もしないままタスクが残り続けるため)。"""
        agent_task = asyncio.create_task(run_agent(content, history, emit))
        watch_task = asyncio.create_task(ws.receive_text())  # 実行中の受信は切断検知にだけ使う
        try:
            done, _ = await asyncio.wait({agent_task, watch_task}, return_when=asyncio.FIRST_COMPLETED)
            if agent_task in done:
                return agent_task.result()
            # 実行中に受信イベントが来た = 切断(UIは実行中の送信を禁止している)
            watch_task.result()  # 切断なら WebSocketDisconnect が上がる
            raise WebSocketDisconnect(code=1000)
        finally:
            for t in (agent_task, watch_task):
                if not t.done():
                    t.cancel()

    global _chat_history
    try:
        while True:
            data = await ws.receive_json()
            if data.get("type") == "reset":
                _chat_history = []
                await emit({"type": "reset_done"})
                continue
            if data.get("type") != "chat":
                continue
            content = str(data.get("content", "")).strip()
            if not content:
                continue
            await emit({"type": "start"})
            try:
                async with _agent_lock:
                    _chat_history = await run_and_watch(content, _chat_history)
            except WebSocketDisconnect:
                logger.info("実行中にクライアントが切断したため、エージェントを中断しました")
                raise
            except Exception:
                logger.exception("エージェント実行エラー")
                await emit({"type": "error", "message": "内部エラーが発生しました。しばらくして再度お試しください。"})
            await emit({"type": "done"})
    except WebSocketDisconnect:
        pass
