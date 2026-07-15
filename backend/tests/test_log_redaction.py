"""ログのシークレット墨消し(log_setup.RedactSecretsFilter)のテスト。

APIキーをログに出さないことは本アプリのセキュリティ方針の要。特に例外の
スタックトレースは getMessage() に含まれず素通ししやすいので、回帰しないよう
テストで固定する。
"""
import io
import logging

from app.log_setup import RedactSecretsFilter

FAKE_KEY = "sk-FAKEKEY1234567890abcdef"  # 本物ではないダミー


def _log(fn, secrets=(FAKE_KEY,)) -> str:
    """フィルタ付きハンドラでfn(logger)を実行し、出力された文字列を返す。"""
    buf = io.StringIO()
    handler = logging.StreamHandler(buf)
    handler.setFormatter(logging.Formatter("%(message)s"))
    handler.addFilter(RedactSecretsFilter(list(secrets)))
    logger = logging.getLogger(f"test_redaction_{id(fn)}")
    logger.handlers = [handler]
    logger.setLevel(logging.INFO)
    logger.propagate = False
    fn(logger)
    return buf.getvalue()


def test_secret_in_message_is_redacted():
    out = _log(lambda lg: lg.warning("鍵: %s", FAKE_KEY))
    assert FAKE_KEY not in out
    assert "***REDACTED***" in out


def test_secret_in_traceback_is_redacted():
    """logger.exception のスタックトレースに混入した鍵も墨消しされること。"""
    def emit(lg):
        try:
            raise RuntimeError(f"request failed url=https://x/?key={FAKE_KEY}")
        except Exception:
            lg.exception("LLM呼び出しに失敗")

    out = _log(emit)
    assert FAKE_KEY not in out
    assert "***REDACTED***" in out


def test_secret_in_chained_traceback_is_redacted():
    """「During handling of the above exception」で連結された例外も対象。"""
    def emit(lg):
        try:
            raise ValueError("内側")
        except Exception:
            try:
                raise RuntimeError(f"外側 key={FAKE_KEY}")
            except Exception:
                lg.exception("ネストした例外")

    out = _log(emit)
    assert FAKE_KEY not in out


def test_bearer_token_is_redacted_even_if_unknown():
    """未知のトークンでも Authorization: Bearer 形式なら墨消しされること。"""
    out = _log(lambda lg: lg.warning("headers: Bearer abcdefGHIJKL0123456789"), secrets=())
    assert "abcdefGHIJKL0123456789" not in out
    assert "Bearer ***REDACTED***" in out


def test_short_values_are_not_redacted():
    """短い値まで置換すると通常のログが壊れるため、8文字未満は対象外。"""
    out = _log(lambda lg: lg.info("シート名は abc です"), secrets=("abc",))
    assert "abc" in out


def test_normal_log_is_untouched():
    out = _log(lambda lg: lg.info("モデル %s は準備済みです", "gpt-oss:120b"))
    assert "モデル gpt-oss:120b は準備済みです" in out
