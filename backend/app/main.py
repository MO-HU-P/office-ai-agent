import asyncio
import json
import logging
from datetime import datetime
from pathlib import Path

import httpx
from fastapi import FastAPI, HTTPException, UploadFile, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from pydantic import BaseModel

from . import config
from .agent.loop import run_agent
from .log_setup import configure_logging
from .services import model_admin
from .services.preview import excel_preview, pptx_preview

configure_logging()
logger = logging.getLogger(__name__)

app = FastAPI(title="Office AI Agent")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

FILE_TYPES = {".docx": "word", ".xlsx": "excel", ".pptx": "powerpoint", ".csv": "csv"}


@app.on_event("startup")
async def auto_pull_local_model():
    """localモード時、設定モデルが未ダウンロードならバックグラウンドで自動pullする"""
    s = config.get_settings()
    if s.mode != "local":
        logger.info("LLM: cloudモード (model=%s, base=%s)", s.model, s.base_url)
        return
    model = s.model

    async def _pull():
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

    asyncio.create_task(_pull())


@app.get("/api/health")
async def health():
    s = config.get_settings()
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
        "mode": s.mode,
        "ollama": ollama_ok,
        "key_missing": key_missing,
        "model": s.model,
        "model_ready": model_ready,
    }


# ---------- 設定・モデル管理 (設定UIから利用) ----------
# セキュリティ方針: APIキーはこれらのAPIで受け取らず・返さず、.envでのみ管理する。


class SettingsUpdate(BaseModel):
    mode: str | None = None
    model_local: str | None = None
    model_cloud: str | None = None
    reasoning: str | None = None


class PullRequest(BaseModel):
    name: str


def _settings_response(s: config.LLMSettings) -> dict:
    return {
        "mode": s.mode,
        "model": s.model,
        "model_local": s.model_local,
        "model_cloud": s.model_cloud,
        "reasoning": s.reasoning,
        # キー本体は返さない。「設定済みかどうか」だけをUIに知らせる
        "cloud_key_configured": bool(config.OLLAMA_API_KEY),
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
async def list_models(mode: str | None = None):
    """指定モード(省略時は現在のモード)のモデル一覧を返す。"""
    mode = (mode or config.get_settings().mode).lower()
    if mode not in config.VALID_MODES:
        raise HTTPException(400, "mode は local / cloud のいずれかです")
    if mode == "cloud" and not config.OLLAMA_API_KEY:
        return {"mode": mode, "models": [], "unavailable": "APIキーが未設定です (.env)"}
    try:
        models = await model_admin.list_models(mode)
    except model_admin.OllamaUnavailable as e:
        return {"mode": mode, "models": [], "unavailable": str(e)}
    return {"mode": mode, "models": models}


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
    dest.write_bytes(content)
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
            return excel_preview(path)
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
    if not str(path).startswith(str(config.PREVIEW_CACHE_DIR)) or not path.exists():
        raise HTTPException(404)
    return FileResponse(str(path), media_type="image/png")


def _existing(filename: str) -> Path:
    try:
        return config.resolve_workspace_path(filename, must_exist=True)
    except FileNotFoundError:
        raise HTTPException(404, f"ファイルが見つかりません: {filename}")
    except ValueError as e:
        raise HTTPException(400, str(e))


@app.websocket("/ws")
async def websocket_endpoint(ws: WebSocket):
    await ws.accept()
    history = []

    async def emit(event: dict):
        await ws.send_json(event)

    try:
        while True:
            data = await ws.receive_json()
            if data.get("type") == "reset":
                history = []
                await emit({"type": "reset_done"})
                continue
            if data.get("type") != "chat":
                continue
            content = str(data.get("content", "")).strip()
            if not content:
                continue
            await emit({"type": "start"})
            try:
                history = await run_agent(content, history, emit)
            except Exception:
                logger.exception("エージェント実行エラー")
                await emit({"type": "error", "message": "内部エラーが発生しました。しばらくして再度お試しください。"})
            await emit({"type": "done"})
    except WebSocketDisconnect:
        pass
