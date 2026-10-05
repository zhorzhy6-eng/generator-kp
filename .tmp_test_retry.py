# -*- coding: utf-8 -*-
"""Временная проверка логики повтора (retry) в llm_provider."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(r"E:\Prpgrammy\Шаблоны") / "src"))
import llm_provider as lp  # noqa: E402


class ConnectTimeout(Exception):
    pass


class ResponseError(Exception):
    pass


class SSLCertVerificationError(Exception):
    pass


class _Client:
    """Клиент-заглушка: падает заданное число раз, потом отдаёт текст."""

    def __init__(self, failures, exc):
        self.calls = 0
        self.failures = failures
        self.exc = exc

    def chat(self, payload):
        self.calls += 1
        if self.calls <= self.failures:
            raise self.exc
        msg = type("M", (), {"content": "готовый текст"})()
        return type("R", (), {"choices": [type("C", (), {"message": msg})()]})()


print("=== 1. Классификация ошибок ===")
cases = [
    (ConnectTimeout("[WinError 10060] Попытка установить соединение была безуспешной"), True),
    (Exception("ConnectTimeout: [WinError 10060] ..."), True),
    (ResponseError("401 Unauthorized"), False),
    (SSLCertVerificationError("certificate verify failed"), False),
    (Exception("certificate verify failed: self-signed certificate"), False),
]
for err, expected in cases:
    got = lp.is_connect_timeout_error(err)
    mark = "OK " if got == expected else "ПРОВАЛ"
    print(f"  [{mark}] {type(err).__name__:28} -> retry={got} (ждали {expected})")

print("\n=== 2. Повтор при сетевом таймауте (успех со 2-й попытки) ===")
c = _Client(failures=1, exc=ConnectTimeout("[WinError 10060]"))
text = lp._generate_gigachat_sync_with_retry(c, "sys", "user", 0.85, 150)
print(f"  попыток: {c.calls} (ждали 2) | текст: {text!r}")

print("\n=== 3. Повтор НЕ помогает — ошибка уходит наружу ===")
c = _Client(failures=99, exc=ConnectTimeout("[WinError 10060]"))
try:
    lp._generate_gigachat_sync_with_retry(c, "sys", "user", 0.85, 150)
    print("  ПРОВАЛ: исключение не выброшено")
except ConnectTimeout:
    print(f"  OK: исключение выброшено после {c.calls} попыток (ждали 2)")

print("\n=== 4. Ошибку авторизации НЕ повторяем ===")
c = _Client(failures=99, exc=ResponseError("401 Unauthorized"))
try:
    lp._generate_gigachat_sync_with_retry(c, "sys", "user", 0.85, 150)
except ResponseError:
    print(f"  OK: 401 не повторялся, попыток: {c.calls} (ждали 1)")

print("\n=== 5. SSL-ошибку НЕ повторяем ===")
c = _Client(failures=99, exc=SSLCertVerificationError("certificate verify failed"))
try:
    lp._generate_gigachat_sync_with_retry(c, "sys", "user", 0.85, 150)
except SSLCertVerificationError:
    print(f"  OK: SSL не повторялся, попыток: {c.calls} (ждали 1)")

print("\n=== 6. Успех с первой попытки ===")
c = _Client(failures=0, exc=ConnectTimeout("нет"))
text = lp._generate_gigachat_sync_with_retry(c, "sys", "user", 0.85, 150)
print(f"  попыток: {c.calls} (ждали 1) | текст: {text!r}")

print("\n=== 7. Настройки по умолчанию ===")
import inspect  # noqa: E402

sig = inspect.signature(lp.generate_gigachat)
print(f"  use_stream по умолчанию: {sig.parameters['use_stream'].default} (ждали False)")
print(f"  DEFAULT_GIGACHAT_TIMEOUT: {lp.DEFAULT_GIGACHAT_TIMEOUT} (ждали 60.0)")
print(f"  DEFAULT_GIGACHAT_BASE_URL: {lp.DEFAULT_GIGACHAT_BASE_URL}")
