# -*- coding: utf-8 -*-
"""Временная диагностика GigaChat: сравнение хостов + проверка sync-запроса."""
import os
import sys
import time
from pathlib import Path

ROOT = Path(r"E:\Prpgrammy\Шаблоны")
sys.path.insert(0, str(ROOT / "src"))

# Читаем .env из dist (тот, что реально использует EXE)
from dotenv import load_dotenv  # noqa: E402

env_path = ROOT / "dist" / "Генератор_КП_GigaChat" / ".env"
load_dotenv(env_path, override=True)
print(f"ENV: {env_path} -> exists={env_path.exists()}")
print("GIGACHAT_SCOPE =", os.environ.get("GIGACHAT_SCOPE"))
print("GIGACHAT_MODEL =", os.environ.get("GIGACHAT_MODEL"))
print("GIGACHAT_TIMEOUT =", os.environ.get("GIGACHAT_TIMEOUT"))

import llm_provider as lp  # noqa: E402

print("\n=== CA bundle ===")
lp.setup_ca_bundle()
print("GIGACHAT_CA_BUNDLE_FILE =", os.environ.get("GIGACHAT_CA_BUNDLE_FILE"))

CANDIDATES = [
    ("default (api.giga.chat)", None),
    ("sberbank legacy", "https://gigachat.devices.sberbank.ru/api/v1"),
]

from gigachat import GigaChat  # noqa: E402
from gigachat.models import Chat, Messages, MessagesRole  # noqa: E402

SYSTEM = lp.DEFAULT_SYSTEM_PROMPT
PROMPT = (
    "Напиши коммерческое предложение по автоперевозкам. "
    "Маршрут: Москва → юг России (Ростов, Краснодар, Новороссийск). "
    "Города: Ростов-на-Дону, Краснодар, Новороссийск. "
    "Упомяни: машины в наличии, оплата безналом, работа по договору. "
    "3 предложения. В конце призыв к действию."
)

for label, base_url in CANDIDATES:
    print(f"\n=== {label} | base_url={base_url} ===")
    kwargs = dict(
        credentials=lp.get_gigachat_credentials(),
        scope=os.environ.get("GIGACHAT_SCOPE", "GIGACHAT_API_PERS"),
        model=os.environ.get("GIGACHAT_MODEL", "GigaChat"),
        timeout=60.0,
        verify_ssl_certs=True,
        ca_bundle_file=os.environ.get("GIGACHAT_CA_BUNDLE_FILE") or None,
    )
    if base_url:
        kwargs["base_url"] = base_url

    try:
        client = GigaChat(**kwargs)
    except Exception as e:
        print(f"  CLIENT FAIL: {type(e).__name__}: {e}")
        continue

    t = time.time()
    try:
        client.get_token()
        print(f"  get_token OK  {time.time() - t:.2f}s")
    except Exception as e:
        print(f"  get_token FAIL {time.time() - t:.2f}s {type(e).__name__}: {str(e)[:200]}")
        continue

    payload = Chat(
        messages=[
            Messages(role=MessagesRole.SYSTEM, content=SYSTEM),
            Messages(role=MessagesRole.USER, content=PROMPT),
        ],
        temperature=0.85,
        max_tokens=150,
    )

    t = time.time()
    try:
        resp = client.chat(payload)
        text = resp.choices[0].message.content
        print(f"  CHAT SYNC OK  {time.time() - t:.2f}s | chars={len(text)}")
        print("  ---TEXT---")
        print(text)
        print("  ---END---")
    except Exception as e:
        print(f"  CHAT SYNC FAIL {time.time() - t:.2f}s {type(e).__name__}: {str(e)[:300]}")

    try:
        client.close()
    except Exception:
        pass
