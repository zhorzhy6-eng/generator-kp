# -*- coding: utf-8 -*-
"""Временный тест: симуляция frozen-режима для llm_provider."""
import os
import sys
from pathlib import Path

# --- Симулируем собранный EXE ДО импорта llm_provider ---
FAKE_EXE_DIR = Path(r"E:\Prpgrammy\Шаблоны\_frozen_sim")
FAKE_EXE_DIR.mkdir(parents=True, exist_ok=True)
sys.frozen = True
sys.executable = str(FAKE_EXE_DIR / "Генератор_КП_GigaChat.exe")

SRC = Path(r"E:\Prpgrammy\Шаблоны\src")
sys.path.insert(0, str(SRC))

import llm_provider  # noqa: E402

print("IS_FROZEN          :", llm_provider.IS_FROZEN)
print("get_env_write_path :", llm_provider.get_env_write_path())
print("search paths       :")
for p in llm_provider._env_search_paths():
    print("   -", p)
print("find_env_file      :", llm_provider.find_env_file())

# --- Проверяем запись ключа в правильное место ---
write_path = llm_provider.get_env_write_path()
write_path.parent.mkdir(parents=True, exist_ok=True)

# Имитируем текущую логику записи из GUI (_save_credentials_to_env)
lines = []
if write_path.exists():
    lines = [
        l.rstrip("\n") for l in write_path.read_text(encoding="utf-8-sig").splitlines()
        if not l.startswith("GIGACHAT_CREDENTIALS")
    ]
if not lines:
    lines = [
        "GIGACHAT_SCOPE=GIGACHAT_API_PERS",
        "GIGACHAT_MODEL=GigaChat",
        "GIGACHAT_TIMEOUT=30",
        "GIGACHAT_VERIFY_SSL_CERTS=true",
        "GIGACHAT_CA_BUNDLE_FILE=",
    ]
lines.append("GIGACHAT_CREDENTIALS=TEST123")
write_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

print("\nПосле записи файл создан:", write_path)
print("Содержимое:")
print(write_path.read_text(encoding="utf-8"))
print("\nfind_env_file после записи:", llm_provider.find_env_file())

# Очистка временной папки
import shutil
shutil.rmtree(FAKE_EXE_DIR, ignore_errors=True)
print("\nВременная папка очищена:", not FAKE_EXE_DIR.exists())
