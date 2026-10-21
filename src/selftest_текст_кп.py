#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Самопроверка генерации текста КП — без запуска графического интерфейса.

Зачем отдельный файл
--------------------
Обе программы собирают окно Tkinter прямо при импорте модуля (глобальные
виджеты), поэтому обычный `import генератор_кп` открыл бы окно и повис бы на
главном цикле. Для проверки это не нужно: берём из файлов ТОЛЬКО функции
сборки текста и Word (через разбор исходника) и подставляем им зависимости.
Проверяются настоящие строки кода, а не их копии в тесте.

Что проверяется
---------------
1. 10 текстов подряд в одном регионе — все разные.
2. Оплата «обсуждается» — в призыве к действию есть упоминание оплаты.
3. Три варианта Word — файлы создаются и заметно отличаются.
4. Заголовок и маршрут содержат города лота и не теряют их.

Запуск:  python src/selftest_текст_кп.py
"""

import ast
import os
import shutil
import sys
import types
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import text_variants as tv  # noqa: E402

TEMPLATE_APP = SRC_DIR / "генератор_кп.py"
AI_APP = SRC_DIR / "генератор_кп_gigachat.py"

FAILURES = []


def check(condition: bool, message: str) -> bool:
    """Печатает результат проверки и копит провалы."""
    mark = "✅" if condition else "❌"
    print(f"  {mark} {message}")
    if not condition:
        FAILURES.append(message)
    return condition


def load_functions(path: Path, names: list, extra_globals: dict) -> dict:
    """
    Достаёт перечисленные функции из файла, не выполняя остальной модуль.

    Зависимости (logger, пулы фраз и прочее) берутся из extra_globals: так
    тест не тянет за собой ни Tkinter, ни GigaChat.
    """
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    wanted = {}
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name in names:
            wanted[node.name] = node

    missing = set(names) - set(wanted)
    if missing:
        raise RuntimeError(f"{path.name}: не найдены функции {sorted(missing)}")

    module = types.ModuleType(f"probe_{path.stem}")
    module.__dict__.update(extra_globals)
    code = ast.Module(body=list(wanted.values()), type_ignores=[])
    ast.fix_missing_locations(code)
    exec(compile(code, str(path), "exec"), module.__dict__)
    return module.__dict__


class SilentLogger:
    """Заглушка логгера: самопроверка не должна писать в рабочий лог."""

    def _noop(self, *args, **kwargs):
        return None

    debug = info = warning = error = critical = exception = _noop


# ============================================
# 1. ВЕРСИЯ С GIGACHAT: ШАБЛОННЫЙ ПУТЬ (fallback)
# ============================================

CURRENT_REGION = {"value": "Юга"}

giga_globals = {
    "logger": SilentLogger(),
    "CREATIVE_MODE": tv.CREATIVE_MODE,
    "DEFAULT_PAYMENT_KEY": tv.DEFAULT_PAYMENT_KEY,
    "LEGALS": tv.LEGALS,
    "_PHRASES": tv._history,
    "assemble_text": tv.assemble_text,
    "build_route_block": tv.build_route_block,
    "build_lot_header": tv.build_lot_header,
    "payment_line": tv.payment_line,
    "build_cta": tv.build_cta,
    "pick_opener": tv.pick_opener,
    "get_current_region": lambda: CURRENT_REGION["value"],
    "get_selected_cities": lambda: [],
    "List": list,
    "Optional": object,
}

giga = load_functions(
    AI_APP,
    ["format_cities_list", "_guess_city_count", "build_template_text"],
    giga_globals,
)

print("\n=== 1. Версия с ИИ (шаблонный fallback): 10 текстов подряд ===")
texts = []
for _ in range(10):
    route_cities = ["Ростов-на-Дону", "Краснодар", "Адыгея"]
    cities_text = giga["format_cities_list"](route_cities)
    texts.append(
        giga["build_template_text"](cities_text, "beznal_nds", len(route_cities))
    )

check(len(set(texts)) == 10, f"все 10 текстов разные (уникальных: {len(set(texts))})")
check(
    all("автовоз" in t.lower() for t in texts),
    "в каждом тексте есть контекст автовоза",
)
check(
    all("Ростов-на-Дону" in t and "Адыгея" in t for t in texts),
    "города лота присутствуют в каждом тексте",
)
check(
    all(t.count("💰") == 1 for t in texts),
    "строка оплаты в тексте ровно одна",
)
# Город встречается дважды законно: в заголовке («Москва → … → Адыгея») и в
# блоке маршрута. Дублирование ВНУТРИ одной строки — уже ошибка оформления.
check(
    all(
        all(line.count("Ростов-на-Дону") <= 1 for line in t.splitlines())
        for t in texts
    ),
    "внутри одной строки город не повторяется",
)

# ============================================
# 2. ОПЛАТА «ОБСУЖДАЕТСЯ» — В ПРИЗЫВЕ ЕСТЬ УПОМИНАНИЕ ОПЛАТЫ
# ============================================

print("\n=== 2. Оплата «обсуждается»: призыв упоминает оплату ===")
tv._history.clear()
discuss_texts = [
    giga["build_template_text"]("Ростов-на-Дону", "discuss", 1) for _ in range(8)
]
CTA_WORDS = ("оплат", "условия", "договоримся")
check(
    all(any(word in t.lower() for word in CTA_WORDS) for t in discuss_texts),
    "в каждом тексте с оплатой «обсуждается» призыв говорит про оплату",
)
cta_lines = [t.splitlines()[-1] for t in discuss_texts]
check(
    all("оплат" in line.lower() or "услови" in line.lower() for line in cta_lines),
    "именно последняя строка (призыв) упоминает оплату",
)
check(
    all("обсужда" in t.lower() for t in discuss_texts),
    "формулировка «оплата обсуждается» осталась в тексте",
)
check(len(set(discuss_texts)) > 1, "варианты с «обсуждается» тоже различаются")

# ============================================
# 3. ШАБЛОННАЯ ВЕРСИЯ: 10 ТЕКСТОВ ПОДРЯД
# ============================================

print("\n=== 3. Шаблонная версия: 10 текстов подряд ===")

tpl_globals = {
    "logger": SilentLogger(),
    "CREATIVE_MODE": tv.CREATIVE_MODE,
    "DEFAULT_PAYMENT_KEY": tv.DEFAULT_PAYMENT_KEY,
    "PAYMENT_KEY": "beznal_nds",
    "CURRENT_REGION": "Юга",
    "LEGALS": tv.LEGALS,
    "STATUSES": tv.STATUSES,
    "_PHRASES": tv._history,
    "assemble_text": tv.assemble_text,
    "build_route_block": tv.build_route_block,
    "build_lot_header": tv.build_lot_header,
    "build_cta": tv.build_cta,
    "payment_line": tv.payment_line,
    "payment_phrase": lambda: tv.payment_line(tpl_globals["PAYMENT_KEY"]),
    "INTROS_SOUTH_ONE": tv.INTROS_SOUTH_ONE,
    "INTROS_SOUTH_MULTI": tv.INTROS_SOUTH_MULTI,
    "INTROS_VLADIVOSTOK_ONE": tv.INTROS_VLADIVOSTOK_ONE,
    "INTROS_VLADIVOSTOK_MULTI": tv.INTROS_VLADIVOSTOK_MULTI,
    "INTROS_EAST_ONE": tv.INTROS_EAST_ONE,
    "INTROS_EAST_MULTI": tv.INTROS_EAST_MULTI,
    "CREATIVE_ROUTE_EAST": ["📍 Маршрут: Москва → Восток"],
    "get_selected_cities": lambda: ["Ростов-на-Дону", "Краснодар", "Адыгея"],
    "get_current_region": lambda: "Юга",
    "List": list,
    "Dict": dict,
}

template = load_functions(
    TEMPLATE_APP,
    ["_pick_intro", "generate_creative_text"],
    tpl_globals,
)

tv._history.clear()
tpl_texts = [template["generate_creative_text"]() for _ in range(10)]
check(len(set(tpl_texts)) == 10, f"все 10 текстов разные (уникальных: {len(set(tpl_texts))})")
check(
    all("автовоз" in t.lower() for t in tpl_texts),
    "в каждом тексте есть контекст автовоза",
)
check(
    all("Ростов-на-Дону" in t and "Адыгея" in t for t in tpl_texts),
    "города лота присутствуют в каждом тексте",
)
check(
    all(
        all(line.count("Ростов-на-Дону") <= 1 for line in t.splitlines())
        for t in tpl_texts
    ),
    "внутри одной строки город не повторяется",
)
check(
    all(t.count("💰") == 1 for t in tpl_texts),
    "строка оплаты в тексте ровно одна",
)

# ============================================
# 4. ТРИ ВАРИАНТА WORD
# ============================================

print("\n=== 4. Три варианта Word ===")

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from datetime import datetime

    # Папка для файлов Word: временная внутри рабочей папки проекта. Писать
    # на рабочий стол при проверке нельзя — программа теста его не трогает.
    temp_dir = str(SRC_DIR / "_selftest_out")
    Path(temp_dir).mkdir(parents=True, exist_ok=True)

    word_globals = {
        "logger": SilentLogger(),
        "os": os,
        "Document": Document,
        "WD_ALIGN_PARAGRAPH": WD_ALIGN_PARAGRAPH,
        "datetime": datetime,
        "PROJECT_ROOT": str(SRC_DIR.parent),
        "__file__": str(SRC_DIR / "генератор_кп.py"),
        "_documents_dir": lambda: temp_dir,
        "generate_creative_text": template["generate_creative_text"],
    }

    word = load_functions(TEMPLATE_APP, ["_save_text_as_docx"], word_globals)

    regions = ("Юга", "Владивосток", "Восток")
    sample_cities = {
        "Юга": ["Ростов-на-Дону", "Краснодар"],
        "Владивосток": ["Владивосток → Москва"],
        "Восток": ["Екатеринбург", "Тюмень", "Омск"],
    }

    saved = []
    contents = []
    for i, region in enumerate(regions, start=1):
        tpl_globals["get_current_region"] = lambda r=region: r
        tpl_globals["get_selected_cities"] = lambda c=sample_cities[region]: list(c)
        text = template["generate_creative_text"]()
        filename = f"КП_{region}_вариант_{i}.docx"
        path = word["_save_text_as_docx"](text, filename)
        saved.append(path)
        contents.append(text)

    check(all(os.path.isfile(p) for p in saved), f"создано файлов: {len(saved)}")
    check(
        all(os.path.getsize(p) > 0 for p in saved),
        "каждый файл не пустой",
    )
    check(len(set(contents)) == 3, "тексты трёх вариантов различаются")

    # Заметность отличий: сравниваем по строкам, а не по символам
    line_sets = [set(t.splitlines()) for t in contents]
    diffs = [
        len(line_sets[i] ^ line_sets[j])
        for i in range(3)
        for j in range(i + 1, 3)
    ]
    check(
        all(d >= 3 for d in diffs),
        f"файлы отличаются минимум на 3 строки (попарно: {diffs})",
    )

    print("     Папка проверки:", temp_dir)
    for p in saved:
        print("     ", os.path.basename(p), f"({os.path.getsize(p)} байт)")

    # Проверяем и убираем за собой: рабочие файлы пользователя не трогаем
    # и мусор в проекте не оставляем. Оставляем только при провале — чтобы
    # было что посмотреть.
    if not FAILURES:
        shutil.rmtree(temp_dir, ignore_errors=True)

except ImportError as exc:  # pragma: no cover
    print(f"  ⚠️ python-docx недоступен, проверка Word пропущена: {exc}")

# ============================================
# 5. ЗАГОЛОВОК И МАРШРУТ: ГОРОДА НА МЕСТЕ
# ============================================

print("\n=== 5. Заголовок и маршрут: данные лота на месте ===")

for region, cities in (
    ("Юга", ["Ростов-на-Дону"]),
    ("Юга", ["Ростов-на-Дону", "Краснодар"]),
    ("Владивосток", ["Владивосток → Москва"]),
    ("Восток", ["Екатеринбург", "Тюмень", "Омск"]),
):
    origin = tv.route_origin(tv.normalize_cities(cities))
    ok_header = True
    ok_route = True
    for _ in range(15):
        header = tv.build_lot_header(region, cities)
        route = tv.build_route_block(cities)
        if origin not in header:
            ok_header = False
        for city in tv.normalize_cities(cities):
            if city not in route:
                ok_route = False
        if f"{origin} → {origin}" in header or f"{origin} → {origin}" in route:
            ok_header = False
    check(ok_header, f"{region} {cities}: город погрузки есть в заголовке, без повторов")
    check(ok_route, f"{region} {cities}: все города есть в блоке маршрута")

# Разнообразие заголовков
titles = {tv.build_lot_header("Юга", ["Ростов-на-Дону"]) for _ in range(20)}
check(len(titles) >= 8, f"заголовки одного лота разнообразны (вариантов: {len(titles)})")

# ============================================
# ИТОГ
# ============================================

print("\n" + "=" * 60)
if FAILURES:
    print(f"❌ ПРОВАЛЕНО ПРОВЕРОК: {len(FAILURES)}")
    for item in FAILURES:
        print("   -", item)
    raise SystemExit(1)

print("✅ ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ")
