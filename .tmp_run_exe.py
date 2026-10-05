# -*- coding: utf-8 -*-
"""Временная проверка: запускаем собранный EXE на 30 сек и читаем его лог."""
import subprocess
import time
from pathlib import Path

BASE = Path(r"E:\Prpgrammy\Шаблоны\dist\Генератор_КП_GigaChat")
EXE = BASE / "Генератор_КП_GigaChat.exe"
LOG = BASE / "logs" / "generator_20261005.log"

before = LOG.read_text(encoding="utf-8", errors="replace").splitlines() if LOG.is_file() else []
print(f"Строк в логе до запуска: {len(before)}")

print(f"Запускаю: {EXE}")
proc = subprocess.Popen(
    [str(EXE)],
    cwd=str(BASE),
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
    stdin=subprocess.DEVNULL,
)
print(f"PID = {proc.pid}")

WAIT = 30
for i in range(WAIT):
    time.sleep(1)
    if proc.poll() is not None:
        print(f"!!! Процесс завершился сам через {i + 1} сек, код {proc.returncode}")
        break
else:
    print(f"Процесс жив через {WAIT} сек — закрываю")

if proc.poll() is None:
    proc.terminate()
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.wait(timeout=10)
print("Процесс закрыт")

time.sleep(1)
after = LOG.read_text(encoding="utf-8", errors="replace").splitlines() if LOG.is_file() else []
new = after[len(before):]

print(f"\n===== НОВЫЕ СТРОКИ ЛОГА ({len(new)}) =====")
for line in new:
    print(line)

print("\n===== ПРОВЕРКИ =====")
joined = "\n".join(new)
checks = [
    ("Прогрев выполнен", "Прогрев GigaChat выполнен" in joined),
    ("НЕТ WinError 10060", "10060" not in joined),
    ("НЕТ 'Потоковая генерация'", "Потоковая генерация" not in joined),
    ("Использован ИИ-текст", "Использован ИИ-текст" in joined),
    ("НЕТ fallback на шаблон", "шаблонный текст (fallback)" not in joined),
    ("НЕТ 'Ошибка GigaChat'", "Ошибка GigaChat" not in joined),
    ("НЕТ SSL-ошибки", "SSL-ошибк" not in joined and "SSL: не удалось" not in joined),
]
for name, ok in checks:
    print(f"  [{'OK  ' if ok else 'ПРОВАЛ'}] {name}")

import re  # noqa: E402

m = re.search(r"GigaChat: ответ получен за ([\d.]+) сек", joined)
if m:
    secs = float(m.group(1))
    print(f"\n  Время генерации: {secs} сек | < 20 сек: {'OK' if secs < 20 else 'ПРЕВЫШЕН'}")

m = re.search(r"Прогрев GigaChat выполнен за ([\d.]+) сек", joined)
if m:
    secs = float(m.group(1))
    print(f"  Время прогрева: {secs} сек | < 2 сек: {'OK' if secs < 2 else 'ПРЕВЫШЕН'}")
