# -*- coding: utf-8 -*-
"""Временная инструментовка: где именно зависает EXE."""
import re
import subprocess
import time
from pathlib import Path

BASE = Path(r"E:\Prpgrammy\Шаблоны\dist\Генератор_КП_GigaChat")
EXE = BASE / "Генератор_КП_GigaChat.exe"
LOG = BASE / "logs" / "generator_20261005.log"

before = LOG.read_text(encoding="utf-8", errors="replace").splitlines()
print(f"Строк в логе до запуска: {len(before)}\n")

proc = subprocess.Popen(
    [str(EXE)],
    cwd=str(BASE),
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
    stdin=subprocess.DEVNULL,
)
pid = proc.pid
print(f"PID = {pid}\n")
print(f"{'сек':>4} | {'лог':>4} | состояние / соединения")
print("-" * 78)

seen = len(before)
for i in range(1, 91):
    time.sleep(1)
    lines = LOG.read_text(encoding="utf-8", errors="replace").splitlines()
    new = lines[seen:]
    seen = len(lines)

    # Соединения процесса
    try:
        ns = subprocess.run(
            ["netstat", "-ano"], capture_output=True, text=True, timeout=10
        ).stdout
        conns = [
            l.split()
            for l in ns.splitlines()
            if l.strip().endswith(str(pid)) and "TCP" in l
        ]
        conn_txt = "; ".join(f"{c[1]}->{c[2]} [{c[3]}]" for c in conns) or "нет TCP"
    except Exception as e:
        conn_txt = f"netstat ошибка: {e}"

    if new:
        for nl in new:
            short = re.sub(r"^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d - ", "", nl)
            short = re.sub(r" - \w+ - ", " | ", short)
            print(f"{i:>4} | +{len(new)} | {short}")

    # Соединения печатаем только если менялись или есть что показать
    if conns if "conns" in dir() else False:
        pass

    if proc.poll() is not None:
        print(f"\n!!! Процесс завершился через {i} сек, код {proc.returncode}")
        break
else:
    print(f"\n{i} сек: процесс жив — закрываю")

print(f"\nСоединения на момент конца: {conn_txt}")
if proc.poll() is None:
    proc.terminate()
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.wait(timeout=10)

time.sleep(1)
after = LOG.read_text(encoding="utf-8", errors="replace").splitlines()
print(f"\n===== ВСЕ НОВЫЕ СТРОКИ ЛОГА ({len(after) - len(before)}) =====")
for line in after[len(before):]:
    print(line)
