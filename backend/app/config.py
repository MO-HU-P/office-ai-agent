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
# LLMのストリーミング中、次のチャンクをこの秒数待っても届かなければ打ち切る
LLM_IDLE_TIMEOUT = float(_agent.get("llm_idle_timeout", 180))
# 1回のLLM応答“全体”の上限秒数。チャンクは届き続けるのに(推論を延々流し続ける等)
# 可視出力もツール呼び出しも出ないまま止まる「無言ハング」を打ち切るための総タイムアウト。
LLM_STEP_TIMEOUT = float(_agent.get("llm_step_timeout", 240))
# 一時的なサーバーエラー(Ollama Cloudの500など)へのリトライ設定。
# 500の波が数十秒続くことがあるため、回数を確保しつつ指数バックオフで待つ。
LLM_MAX_ATTEMPTS = int(_agent.get("llm_max_attempts", 5))
LLM_RETRY_BACKOFF_CAP = float(_agent.get("llm_retry_backoff_cap", 30))

# --- 環境変数 (インフラ・シークレット。実行時変更しない) ---
OLLAMA_API_KEY = os.environ.get("OLLAMA_API_KEY", "").strip()
OLLAMA_CLOUD_URL = os.environ.get("OLLAMA_CLOUD_URL", "https://ollama.com")
OLLAMA_LOCAL_URL = os.environ.get("OLLAMA_BASE_URL", "http://localhost:11434")
# 鍵を持つユーザー向けの追加プロバイダー(任意)。Ollamaが主役・唯一のゼロ設定入口で、
# ここは.envに鍵を入れた人だけのopt-in。鍵は.envのみで管理し、APIには出さない。
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()

DEFAULT_LOCAL_MODEL = "qwen3.5:9b"
DEFAULT_CLOUD_MODEL = "gpt-oss:120b"
# DEFAULT_OPENAI_MODEL は OpenAIプリセット(config.toml)の先頭から後段で決める

# provider: どのLLMプロバイダーを使うか。"ollama" は従来どおり mode(local/cloud)で
# 実行場所を切り替える。"openai"/"gemini" 等の外部プロバイダーは mode を持たない(常にクラウド)。
VALID_PROVIDERS = {"ollama", "openai", "gemini"}
VALID_MODES = {"local", "cloud"}
# "auto"=パラメータを送らずモデル既定に任せる / true・false=思考のオンオフ(qwen3等)
# low・medium・high=思考の深さ(gpt-oss等のレベル対応モデル)
VALID_REASONING = {"auto", "true", "false", "low", "medium", "high"}

# Ollamaのモデル名: 例 qwen3:8b, gpt-oss:120b, library/model:tag
_MODEL_NAME_RE = re.compile(
    r"^[A-Za-z0-9][A-Za-z0-9._-]*(?:/[A-Za-z0-9][A-Za-z0-9._-]*)?(?::[A-Za-z0-9._-]+)?$"
)


# APIキーは "sk-"(OpenAI等) や "AIza"(Google) で始まる。どのプロバイダーのモデル名としても
# 正当でないため、すべてのモデル名欄(local/cloud/openai/gemini)で弾き、秘密情報を
# settings.json に保存させない(キーは.envのみで管理)。
_API_KEY_LIKE_RE = re.compile(r"^(?:sk-|AIza)", re.IGNORECASE)


def validate_model_name(name: str) -> str:
    name = (name or "").strip()
    if _API_KEY_LIKE_RE.match(name):
        raise ValueError("APIキーのような値は入力できません。モデル名を入力してください。")
    if not name or len(name) > 128 or not _MODEL_NAME_RE.match(name):
        raise ValueError(f"モデル名の形式が正しくありません: {name[:64]}")
    return name


# 外部プロバイダーで設定UIに出す推奨モデル候補(config.toml の llm.openai_models /
# llm.gemini_models で編集可)。モデルは順次廃止されるため、コードに埋め込まずデータで持つ。
# 廃止時はconfig.tomlを直せばよい。
_DEFAULT_OPENAI_PRESET = ("gpt-4o-mini",)
_DEFAULT_GEMINI_PRESET = ("gemini-2.5-flash",)


def _load_presets(raw, toml_key: str, fallback: tuple[str, ...]) -> list[str]:
    names: list[str] = []
    for item in raw if isinstance(raw, list) else ():
        try:
            names.append(validate_model_name(str(item)))
        except ValueError:
            logger.warning("config.toml の %s に不正なモデル名: %s", toml_key, str(item)[:64])
    return names or list(fallback)


OPENAI_PRESET_MODELS = _load_presets(_llm.get("openai_models"), "openai_models", _DEFAULT_OPENAI_PRESET)
GEMINI_PRESET_MODELS = _load_presets(_llm.get("gemini_models"), "gemini_models", _DEFAULT_GEMINI_PRESET)
# 既定モデルは各プリセットの先頭(config.tomlで制御可能)
DEFAULT_OPENAI_MODEL = OPENAI_PRESET_MODELS[0]
DEFAULT_GEMINI_MODEL = GEMINI_PRESET_MODELS[0]


@dataclass(frozen=True)
class LLMSettings:
    """実行時に変更可能なLLM設定。設定UIまたはconfig.tomlから供給される。"""

    provider: str        # "ollama" | "openai" | "gemini" (VALID_PROVIDERS)
    mode: str            # ollama時のみ有効: "local" | "cloud"
    model_local: str     # ollama localモードで使うモデル
    model_cloud: str     # ollama cloudモードで使うモデル
    model_openai: str    # openaiプロバイダーで使うモデル
    model_gemini: str    # geminiプロバイダーで使うモデル
    reasoning: str       # VALID_REASONING のいずれか
    num_ctx: int         # ollama localモードのみ有効 (config.tomlでのみ変更)
    # 設定UIで自由入力し保存した外部プロバイダーのモデル候補(settings.jsonに永続化)。
    # プリセット(config.tomlの*_models)とは別で、こちらは削除可能。イミュータブルにtupleで保持。
    openai_custom_models: tuple[str, ...] = ()
    gemini_custom_models: tuple[str, ...] = ()

    @property
    def model(self) -> str:
        """現在のプロバイダー/モードでの実効モデル名。"""
        if self.provider == "openai":
            return self.model_openai
        if self.provider == "gemini":
            return self.model_gemini
        return self.model_cloud if self.mode == "cloud" else self.model_local

    @property
    def is_cloud(self) -> bool:
        """外部クラウド(ネット越し)を使うか。ヘッダー表示や鍵チェックの判定に使う。"""
        return self.provider != "ollama" or self.mode == "cloud"

    @property
    def base_url(self) -> str:
        """Ollama用のベースURL(provider=='ollama'のときのみ意味を持つ)。"""
        return OLLAMA_CLOUD_URL if self.mode == "cloud" else OLLAMA_LOCAL_URL

    def headers(self) -> dict[str, str]:
        """Ollama APIへのリクエストヘッダー(Cloudモード時はAPIキーを付与)"""
        if self.mode == "cloud" and OLLAMA_API_KEY:
            return {"Authorization": f"Bearer {OLLAMA_API_KEY}"}
        return {}


def _load_initial_settings() -> LLMSettings:
    """config.toml を土台に、settings.json(UIからの変更)があれば上書きして読み込む。"""
    provider = str(_llm.get("provider", "ollama")).lower()
    if provider not in VALID_PROVIDERS:
        provider = "ollama"
    mode = str(_llm.get("mode", "local")).lower()
    if mode not in VALID_MODES:
        mode = "local"
    toml_model = str(_llm.get("model", "")).strip()
    reasoning = str(_llm.get("reasoning", "auto")).lower()
    if reasoning not in VALID_REASONING:
        reasoning = "auto"

    # config.tomlのmodelは「その(provider,mode)用のモデル」として扱い、他は既定値で埋める
    model_local = toml_model if provider == "ollama" and mode == "local" and toml_model else DEFAULT_LOCAL_MODEL
    model_cloud = toml_model if provider == "ollama" and mode == "cloud" and toml_model else DEFAULT_CLOUD_MODEL
    model_openai = toml_model if provider == "openai" and toml_model else DEFAULT_OPENAI_MODEL
    model_gemini = toml_model if provider == "gemini" and toml_model else DEFAULT_GEMINI_MODEL

    settings = LLMSettings(
        provider=provider,
        mode=mode,
        model_local=model_local,
        model_cloud=model_cloud,
        model_openai=model_openai,
        model_gemini=model_gemini,
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
    if "provider" in changes and changes["provider"] is not None:
        provider = str(changes["provider"]).lower()
        if provider not in VALID_PROVIDERS:
            raise ValueError(f"provider は {' / '.join(sorted(VALID_PROVIDERS))} のいずれかです: {provider[:32]}")
        settings = replace(settings, provider=provider)
    if "mode" in changes and changes["mode"] is not None:
        mode = str(changes["mode"]).lower()
        if mode not in VALID_MODES:
            raise ValueError(f"mode は local / cloud のいずれかです: {mode[:32]}")
        settings = replace(settings, mode=mode)
    for model_field in ("model_local", "model_cloud", "model_openai", "model_gemini"):
        if model_field in changes and changes[model_field] is not None:
            settings = replace(settings, **{model_field: validate_model_name(str(changes[model_field]))})
    for custom_field in ("openai_custom_models", "gemini_custom_models"):
        if custom_field in changes and changes[custom_field] is not None:
            raw = changes[custom_field]
            if not isinstance(raw, list):
                raise ValueError(f"{custom_field} はリスト形式で指定してください")
            cleaned: list[str] = []
            for item in raw:
                name = validate_model_name(str(item))
                if name not in cleaned:
                    cleaned.append(name)
            settings = replace(settings, **{custom_field: tuple(cleaned)})
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
            "設定を更新しました: provider=%s mode=%s model=%s reasoning=%s",
            new_settings.provider, new_settings.mode, new_settings.model, new_settings.reasoning,
        )
        return new_settings


def _persist(settings: LLMSettings) -> None:
    """一時ファイル+os.replaceでアトミックに保存(書きかけファイルを残さない)。"""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "provider": settings.provider,
        "mode": settings.mode,
        "model_local": settings.model_local,
        "model_cloud": settings.model_cloud,
        "model_openai": settings.model_openai,
        "model_gemini": settings.model_gemini,
        "openai_custom_models": list(settings.openai_custom_models),
        "gemini_custom_models": list(settings.gemini_custom_models),
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
    # 文字列の前方一致では /workspace2 のような「名前が同じで始まる別ディレクトリ」を通してしまう
    if path != WORKSPACE_DIR and not path.is_relative_to(WORKSPACE_DIR):
        raise ValueError(f"ワークスペース外のパスは指定できません: {filename}")
    if must_exist and not path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {filename}")
    return path
