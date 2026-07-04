import json
import logging
import os
import re
import threading
import tomllib
from dataclasses import dataclass, replace
from pathlib import Path

logger = logging.getLogger(__name__)

WORKSPACE_DIR = Path(os.environ.get("WORKSPACE_DIR", "/workspace")).resolve()
PREVIEW_CACHE_DIR = Path(os.environ.get("PREVIEW_CACHE_DIR", "/tmp/preview_cache")).resolve()

# --- config.toml (デフォルト設定。上級者向けの編集口) ---
CONFIG_PATH = Path(os.environ.get("CONFIG_PATH", "/app/config.toml"))
# --- 設定UIからの変更の保存先 (config.toml より優先される) ---
DATA_DIR = Path(os.environ.get("DATA_DIR", "/app/data"))
SETTINGS_PATH = DATA_DIR / "settings.json"

_cfg: dict = {}
if CONFIG_PATH.exists():
    with open(CONFIG_PATH, "rb") as f:
        _cfg = tomllib.load(f)
else:
    logger.warning("config.toml が見つかりません (%s)。デフォルト設定で起動します", CONFIG_PATH)

_llm = _cfg.get("llm", {})
_agent = _cfg.get("agent", {})

MAX_AGENT_STEPS = int(_agent.get("max_steps", 15))
MAX_HISTORY_MESSAGES = int(_agent.get("max_history_messages", 40))

# --- 環境変数 (インフラ・シークレット。実行時変更しない) ---
OLLAMA_API_KEY = os.environ.get("OLLAMA_API_KEY", "").strip()
OLLAMA_CLOUD_URL = os.environ.get("OLLAMA_CLOUD_URL", "https://ollama.com")
OLLAMA_LOCAL_URL = os.environ.get("OLLAMA_BASE_URL", "http://localhost:11434")

DEFAULT_LOCAL_MODEL = "qwen3.5:9b"
DEFAULT_CLOUD_MODEL = "gpt-oss:120b"

VALID_MODES = {"local", "cloud"}
# "auto"=パラメータを送らずモデル既定に任せる / true・false=思考のオンオフ(qwen3等)
# low・medium・high=思考の深さ(gpt-oss等のレベル対応モデル)
VALID_REASONING = {"auto", "true", "false", "low", "medium", "high"}

# Ollamaのモデル名: 例 qwen3:8b, gpt-oss:120b, library/model:tag
_MODEL_NAME_RE = re.compile(
    r"^[A-Za-z0-9][A-Za-z0-9._-]*(?:/[A-Za-z0-9][A-Za-z0-9._-]*)?(?::[A-Za-z0-9._-]+)?$"
)


def validate_model_name(name: str) -> str:
    name = (name or "").strip()
    if not name or len(name) > 128 or not _MODEL_NAME_RE.match(name):
        raise ValueError(f"モデル名の形式が正しくありません: {name[:64]}")
    return name


@dataclass(frozen=True)
class LLMSettings:
    """実行時に変更可能なLLM設定。設定UIまたはconfig.tomlから供給される。"""

    mode: str            # "local" | "cloud"
    model_local: str     # localモードで使うモデル
    model_cloud: str     # cloudモードで使うモデル
    reasoning: str       # VALID_REASONING のいずれか
    num_ctx: int         # localモードのみ有効 (config.tomlでのみ変更)

    @property
    def model(self) -> str:
        return self.model_cloud if self.mode == "cloud" else self.model_local

    @property
    def base_url(self) -> str:
        return OLLAMA_CLOUD_URL if self.mode == "cloud" else OLLAMA_LOCAL_URL

    def headers(self) -> dict[str, str]:
        """Ollama APIへのリクエストヘッダー(Cloudモード時はAPIキーを付与)"""
        if self.mode == "cloud" and OLLAMA_API_KEY:
            return {"Authorization": f"Bearer {OLLAMA_API_KEY}"}
        return {}


def _load_initial_settings() -> LLMSettings:
    """config.toml を土台に、settings.json(UIからの変更)があれば上書きして読み込む。"""
    mode = str(_llm.get("mode", "local")).lower()
    if mode not in VALID_MODES:
        mode = "local"
    toml_model = str(_llm.get("model", "")).strip()
    reasoning = str(_llm.get("reasoning", "auto")).lower()
    if reasoning not in VALID_REASONING:
        reasoning = "auto"

    # config.tomlのmodelは「そのmode用のモデル」として扱い、反対側は既定値で埋める
    model_local = toml_model if mode == "local" and toml_model else DEFAULT_LOCAL_MODEL
    model_cloud = toml_model if mode == "cloud" and toml_model else DEFAULT_CLOUD_MODEL

    settings = LLMSettings(
        mode=mode,
        model_local=model_local,
        model_cloud=model_cloud,
        reasoning=reasoning,
        num_ctx=int(_llm.get("num_ctx", 8192)),
    )

    if SETTINGS_PATH.exists():
        try:
            saved = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            settings = _apply_changes(settings, saved)
            logger.info("設定UIの保存値を読み込みました (%s)", SETTINGS_PATH)
        except (ValueError, OSError) as e:
            logger.warning("settings.json の読み込みに失敗したため config.toml の値を使います: %s", e)
    return settings


def _apply_changes(settings: LLMSettings, changes: dict) -> LLMSettings:
    """検証しながら設定変更を適用する。不正な値は ValueError。"""
    if "mode" in changes and changes["mode"] is not None:
        mode = str(changes["mode"]).lower()
        if mode not in VALID_MODES:
            raise ValueError(f"mode は local / cloud のいずれかです: {mode[:32]}")
        settings = replace(settings, mode=mode)
    if "model_local" in changes and changes["model_local"] is not None:
        settings = replace(settings, model_local=validate_model_name(str(changes["model_local"])))
    if "model_cloud" in changes and changes["model_cloud"] is not None:
        settings = replace(settings, model_cloud=validate_model_name(str(changes["model_cloud"])))
    if "reasoning" in changes and changes["reasoning"] is not None:
        reasoning = str(changes["reasoning"]).lower()
        if reasoning not in VALID_REASONING:
            raise ValueError(f"reasoning の値が不正です: {reasoning[:32]}")
        settings = replace(settings, reasoning=reasoning)
    return settings


_settings_lock = threading.Lock()
_settings = _load_initial_settings()


def get_settings() -> LLMSettings:
    """現在のLLM設定のスナップショットを返す(イミュータブル)。"""
    with _settings_lock:
        return _settings


def update_settings(changes: dict) -> LLMSettings:
    """設定を検証・適用し、settings.json に永続化して新しい設定を返す。"""
    global _settings
    with _settings_lock:
        new_settings = _apply_changes(_settings, changes)
        _persist(new_settings)
        _settings = new_settings
        logger.info(
            "設定を更新しました: mode=%s model=%s reasoning=%s",
            new_settings.mode, new_settings.model, new_settings.reasoning,
        )
        return new_settings


def _persist(settings: LLMSettings) -> None:
    """一時ファイル+os.replaceでアトミックに保存(書きかけファイルを残さない)。"""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "mode": settings.mode,
        "model_local": settings.model_local,
        "model_cloud": settings.model_cloud,
        "reasoning": settings.reasoning,
    }
    tmp = SETTINGS_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(tmp, SETTINGS_PATH)


OFFICE_EXTENSIONS = {".docx", ".xlsx", ".pptx", ".csv"}

WORKSPACE_DIR.mkdir(parents=True, exist_ok=True)
PREVIEW_CACHE_DIR.mkdir(parents=True, exist_ok=True)


def resolve_workspace_path(filename: str, must_exist: bool = False) -> Path:
    """ワークスペース内のパスに解決する。外側へのトラバーサルは拒否。"""
    path = (WORKSPACE_DIR / filename).resolve()
    if not str(path).startswith(str(WORKSPACE_DIR)):
        raise ValueError(f"ワークスペース外のパスは指定できません: {filename}")
    if must_exist and not path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {filename}")
    return path
