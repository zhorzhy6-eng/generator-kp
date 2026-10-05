# -*- coding: utf-8 -*-
"""Временная E2E-проверка: реальный generate_gigachat() с новыми настройками."""
import os
import sys
import time
from pathlib import Path

ROOT = Path(r"E:\Prpgrammy\Шаблоны")
sys.path.insert(0, str(ROOT / "src"))

from dotenv import load_dotenv  # noqa: E402

env_path = ROOT / "dist" / "Генератор_КП_GigaChat" / ".env"
load_dotenv(env_path, override=True)

import llm_provider as lp  # noqa: E402

print("=== КОНФИГУРАЦИЯ ===")
print("timeout из .env      :", os.environ.get("GIGACHAT_TIMEOUT"))
print("base_url из .env     :", os.environ.get("GIGACHAT_BASE_URL"))
print("_resolve_timeout()   :", lp._resolve_timeout())
print("_resolve_base_url()  :", lp._resolve_base_url())
print("DEFAULT_GIGACHAT_TIMEOUT:", lp.DEFAULT_GIGACHAT_TIMEOUT)
print("use_stream по умолч. :", lp.generate_gigachat.__defaults__)
lp.setup_ca_bundle()

# Системный промпт и промпт — ровно как в генераторе КП
from генератор_кп_gigachat import (  # noqa: E402
    AIGuard,
    AI_MAX_TOKENS,
    AI_TEMPERATURE,
    GIGACHAT_SYSTEM_PROMPT,
    _build_prompt,
)

print("\n=== ПАРАМЕТРЫ ГЕНЕРАЦИИ ===")
print("AI_MAX_TOKENS  :", AI_MAX_TOKENS)
print("AI_TEMPERATURE :", AI_TEMPERATURE)
print("AI_TIMEOUT     :", __import__("генератор_кп_gigachat").AI_TIMEOUT)

print("\n=== ПРОГРЕВ ===")
t = time.time()
ok = lp.warmup_gigachat()
print(f"warmup -> {ok} за {time.time() - t:.2f} сек")

print("\n=== ГЕНЕРАЦИЯ (боевой путь, use_stream по умолчанию) ===")
prompt = _build_prompt("Юга", ["Ростов-на-Дону", "Краснодар", "Новороссийск"])
t = time.time()
text = lp.generate_gigachat(
    prompt=prompt,
    system_prompt=GIGACHAT_SYSTEM_PROMPT,
    temperature=AI_TEMPERATURE,
    max_tokens=AI_MAX_TOKENS,
)
elapsed = time.time() - t
print(f"generate_gigachat -> {'ТЕКСТ' if text else 'None'} за {elapsed:.2f} сек")

if text:
    is_valid, reason = AIGuard.validate_ai_response(text)
    print("валидация AIGuard:", is_valid, "|", reason)
    print("символов:", len(text))
    sys.stdout.buffer.write(("---ТЕКСТ---\n" + text + "\n---КОНЕЦ---\n").encode("utf-8"))
else:
    print("!!! ТЕКСТА НЕТ — будет шаблонный fallback")

print("\n=== ИТОГ ===")
print("WinError 10060 в этом запуске:", lp.has_ssl_error())
print(f"Время: {elapsed:.2f} сек | Лимит 20 сек: {'OK' if elapsed < 20 else 'ПРЕВЫШЕН'}")
