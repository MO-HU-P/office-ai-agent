"""ログ設定とシークレット墨消し。

セキュリティ方針:
- 詳細な例外・スタックトレースはサーバー内部ログにのみ出す(ブラウザには汎用文言のみ返す)。
- APIキー等のシークレットは、万一ログ文字列に混入しても伏せ字に置換する。
- 外部ライブラリ(httpx等)のリクエストURLログはノイズかつ情報漏えい源になるため抑制する。
"""
import logging
import os
import re

from . import config

_BEARER_RE = re.compile(r"(Bearer\s+)[A-Za-z0-9._\-]+", re.IGNORECASE)


class RedactSecretsFilter(logging.Filter):
    """ログレコードからAPIキーやBearerトークンを伏せ字にする。"""

    def __init__(self, secrets: list[str]):
        super().__init__()
        # 短すぎる値の誤置換を避けるため一定長以上のみ対象
        self._secrets = sorted({s for s in secrets if s and len(s) >= 8}, key=len, reverse=True)

    def _scrub(self, text: str) -> str:
        for s in self._secrets:
            if s in text:
                text = text.replace(s, "***REDACTED***")
        return _BEARER_RE.sub(r"\1***REDACTED***", text)

    def filter(self, record: logging.LogRecord) -> bool:
        try:
            msg = record.getMessage()
        except Exception:
            return True
        scrubbed = self._scrub(msg)
        if scrubbed != msg:
            record.msg = scrubbed
            record.args = ()
        # 例外のスタックトレースは getMessage() に含まれないため個別に墨消しする。
        # logger.exception(...) の例外文字列(SDKがURLやヘッダーを載せることがある)は
        # 混入経路として最も可能性が高い。exc_text を埋めておくと Formatter は
        # formatException() を呼ばずこの墨消し済み文字列を使う
        if record.exc_info:
            if not record.exc_text:
                record.exc_text = logging.Formatter().formatException(record.exc_info)
            record.exc_text = self._scrub(record.exc_text)
        if record.stack_info:
            record.stack_info = self._scrub(record.stack_info)
        return True


def configure_logging() -> None:
    level_name = os.environ.get("LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)

    root = logging.getLogger()
    root.setLevel(level)
    if not root.handlers:
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        root.addHandler(handler)

    # 全ハンドラにシークレット墨消しフィルタを付与(全プロバイダーの鍵を対象)
    redactor = RedactSecretsFilter([config.OLLAMA_API_KEY, config.OPENAI_API_KEY, config.GEMINI_API_KEY])
    for h in root.handlers:
        h.addFilter(redactor)
    for name in ("uvicorn", "uvicorn.error", "uvicorn.access"):
        lg = logging.getLogger(name)
        for h in lg.handlers:
            h.addFilter(redactor)

    # 外部HTTPライブラリ/SDKの詳細ログ(リクエストURL・ヘッダー等)を抑制する。
    # 万一漏れても上のRedactSecretsFilterが墨消しするが、そもそも出力させない多重防御。
    # openai SDKはDEBUG時にAuthorizationヘッダーを含むリクエスト詳細を出しうるため必ず抑制する。
    # google-genai SDK(google_genai)も同様にリクエスト詳細ログを出しうるため抑制する。
    for noisy in ("httpx", "httpcore", "openai", "openai._base_client", "google_genai", "google_genai.models"):
        logging.getLogger(noisy).setLevel(logging.WARNING)
