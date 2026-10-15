#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Генератор коммерческих предложений (шаблонная версия, без ИИ)
Версия: 2.0 — интерфейс на CustomTkinter

Работает без ключа GigaChat и без интернета: текст коммерческого предложения
собирается из готовых фраз по выбранному региону и типу оплаты.

Что сохранено из прежней версии:
  * списки городов и креативный режим «Без города»;
  * типы оплаты PAYMENT_LABELS_MAP / PAYMENT_PHRASES (ключи совпадают с
    версией для GigaChat, поэтому settings.json у обеих версий один и тот же);
  * функции format_route, generate_creative_text, payment_phrase;
  * автосохранение настроек и сохранение при закрытии окна.

Что добавлено:
  * выбор городов запоминается ОТДЕЛЬНО для каждого региона (cities_by_region):
    раньше при переключении вкладки выбор предыдущего региона затирался;
  * история фраз RecentHistory — одна и та же строка не повторяется подряд;
  * открывающие фразы INTROS_*_ONE / INTROS_*_MULTI подбираются по количеству
    городов в лоте: «одно направление» и «сборный маршрут» — разные предложения;
  * списки фраз по регионам расширены.
"""

import json
import logging
import os
import random
import sys
from datetime import datetime
from typing import Dict, List, Optional, Sequence

import customtkinter as ctk
import pyperclip
import tkinter as tk
from tkinter import messagebox

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Пути проекта и настройка логирования (общий модуль с версией для GigaChat)
from logger_config import setup_logger, SETTINGS_PATH, LEGACY_LOG_FILE

# ============================================
# ОФОРМЛЕНИЕ (CustomTkinter)
# ============================================
# Тема задаётся ДО создания окна: CustomTkinter читает её при создании
# первого виджета.

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

# Палитра — та же, что в версии с GigaChat: обе программы должны выглядеть
# как одно семейство.
BG_WINDOW = "#12151a"
BG_CARD = "#1b1f27"
BG_HEADER = "#171b22"
BG_STATUS = "#0e1116"
TEXT_MAIN = "#e9ecf1"
TEXT_DIM = "#98a2b3"

COLOR_PRIMARY = "#2e7d32"
COLOR_PRIMARY_HOVER = "#256428"
COLOR_ACCENT = "#1565c0"
COLOR_ACCENT_HOVER = "#11519c"
COLOR_VIOLET = "#7b1fa2"
COLOR_VIOLET_HOVER = "#641a84"
COLOR_WARN = "#ef6c00"
COLOR_WARN_HOVER = "#c85b00"
COLOR_DANGER = "#c62828"
COLOR_DANGER_HOVER = "#a42020"
COLOR_TOOL = "#37474f"
COLOR_TOOL_HOVER = "#2c3a41"

LEFT_COLUMN_WIDTH = 380

logging.getLogger("Генератор_КП").setLevel(logging.DEBUG)
logger = setup_logger("Генератор_КП", level=logging.DEBUG)

logger.info("=" * 60)
logger.info("ЗАПУСК ПРОГРАММЫ ГЕНЕРАТОР КП (шаблонная версия, CustomTkinter)")
logger.info("=" * 60)

# ============================================
# ФАЙЛ ДЛЯ СОХРАНЕНИЯ НАСТРОЕК
# ============================================

SETTINGS_FILE = SETTINGS_PATH
LOG_FILE = LEGACY_LOG_FILE

REGIONS = ("Юга", "Владивосток", "Восток")
DEFAULT_REGION = "Юга"


def _normalize_region(value) -> str:
    """Приводит название региона из настроек к одному из REGIONS."""
    text = str(value or "").strip()
    for region in REGIONS:
        if text.lower() == region.lower():
            return region
    return DEFAULT_REGION


def load_settings():
    """Загружает настройки из JSON файла"""
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            logger.info(
                f"Настройки загружены: городов в текущем регионе="
                f"{len(data.get('cities', []))}, регион={data.get('region', 'не указан')}"
            )
            return data
    except FileNotFoundError:
        logger.warning("Файл настроек не найден. Использую настройки по умолчанию.")
        return None
    except json.JSONDecodeError as e:
        logger.error(f"Ошибка чтения JSON: {e}")
        return None
    except Exception as e:
        logger.error(f"Неизвестная ошибка при загрузке настроек: {e}")
        return None


def _read_cities_by_region(saved: Optional[dict]) -> Dict[str, List[str]]:
    """
    Выбор городов по регионам с переносом старого формата настроек.

    Новый формат: {"cities_by_region": {"Юга": [...], ...}}
    Старый формат: {"cities": [...], "region": "Юга"} — один список на всё.

    Без миграции у пользователя после обновления пропал бы его выбор городов.
    """
    result: Dict[str, List[str]] = {}
    if not saved:
        return result

    raw = saved.get("cities_by_region")
    if isinstance(raw, dict):
        for region, cities in raw.items():
            if str(region) in REGIONS and isinstance(cities, list):
                result[str(region)] = [str(city) for city in cities]
        if result:
            return result

    legacy_cities = saved.get("cities")
    if isinstance(legacy_cities, list):
        legacy_region = _normalize_region(saved.get("region"))
        result[legacy_region] = [str(city) for city in legacy_cities]
        logger.info(
            "Настройки: старый формат перенесён в cities_by_region[%s] (%d городов)",
            legacy_region,
            len(result[legacy_region]),
        )
    return result


def save_settings(selected_cities, region, payment_key=None):
    """
    Сохраняет настройки в JSON файл.

    Выбор городов пишется в двух видах:
      * cities_by_region — свой список для каждого региона (основной формат);
      * cities           — выбор текущего региона. Поле оставлено для
        совместимости: его читают уже разосланные сборки и версия с GigaChat.
    """
    if payment_key is None:
        payment_key = globals().get("PAYMENT_KEY", DEFAULT_PAYMENT_KEY)
    try:
        cities_by_region[region] = list(selected_cities)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "cities_by_region": dict(cities_by_region),
                    "cities": list(selected_cities),
                    "region": region,
                    "payment": payment_key,
                },
                f,
                ensure_ascii=False,
                indent=2,
            )
        logger.info(
            f"Настройки сохранены: городов={len(selected_cities)}, "
            f"регион={region}, оплата={payment_key}"
        )
        return True
    except Exception as e:
        logger.error(f"Ошибка сохранения настроек: {e}")
        return False


# ============================================
# СПИСКИ ГОРОДОВ
# ============================================

ALL_CITIES_SOUTH = [
    "Ростов-на-Дону",
    "Краснодар",
    "Новороссийск",
    "Адыгея",
    "Пятигорск",
    "Минеральные Воды",
]

ALL_CITIES_VLADIVOSTOK = [
    "Владивосток → Москва",
    "Владивосток → Новосибирск",
    "Владивосток → Краснодар",
    "Владивосток → Екатеринбург",
    "Владивосток → Казань",
    "Владивосток → Санкт-Петербург",
]

ALL_CITIES_EAST = [
    "Екатеринбург",
    "Тюмень",
    "Омск",
    "Новосибирск",
    "Красноярск",
    "Иркутск",
    "Улан-Удэ",
    "Чита",
    "Курган",
    "Кемерово",
    "Барнаул",
    "Абакан",
    "Благовещенск",
    "Владивосток",
    "Нижний Новгород",
    "Чебоксары",
]

CREATIVE_MODE = "✨ Без города (креативный режим)"

# Города по умолчанию для региона, который ещё не открывали. Креативный режим
# сюда не входит: это режим, а не город (раньше он отмечался вместе со всеми
# городами Востока, и в тексте получалось «Без города» плюс список городов).
DEFAULT_CITIES_BY_REGION: Dict[str, List[str]] = {
    "Юга": list(ALL_CITIES_SOUTH),
    "Владивосток": list(ALL_CITIES_VLADIVOSTOK),
    "Восток": list(ALL_CITIES_EAST),
}

TAB_TITLES: Dict[str, str] = {
    "Юга": "📍 Юга",
    "Владивосток": "📍 Владивосток",
    "Восток": "📍 Восток",
}
TAB_TITLES_BY_NAME: Dict[str, str] = {title: region for region, title in TAB_TITLES.items()}

# ============================================
# ТИП ОПЛАТЫ
# ============================================
# Ключи совпадают с llm_provider.PAYMENT_TYPES, чтобы settings.json был
# одинаковым для версии с ИИ и шаблонной версии.
PAYMENT_LABELS_MAP = {
    "beznal_nds": "Безнал с НДС",
    "beznal": "Безнал без НДС",
    "cash": "Наличные",
    "discuss": "Обсуждается",
}
PAYMENT_LABELS = list(PAYMENT_LABELS_MAP.values())
DEFAULT_PAYMENT_KEY = "beznal_nds"

# Что писать про оплату для каждого варианта: фразы «Оплата … (с НДС)»
PAYMENT_PHRASES = {
    "beznal_nds": [
        "💰 Оплата по безналу (с НДС).",
        "💰 Оплата по безналу (с НДС), сроки обсуждаем.",
        "💰 Оплата — безналичный расчёт с НДС, работаем с юрлицами.",
    ],
    "beznal": [
        "💰 Оплата по безналу (без НДС).",
        "💰 Оплата по безналу (без НДС), сроки обсуждаем.",
        "💰 Безналичный расчёт без НДС, работаем с юрлицами.",
    ],
    "cash": [
        "💰 Оплата наличными (без НДС).",
        "💰 Оплата наличными (без НДС), по факту выгрузки.",
        "💰 Оплата наличными (без НДС), договор оформляем.",
    ],
    "discuss": [
        "💰 Оплата обсуждается — безнал с НДС или наличные.",
        "💰 Условия оплаты обсуждаются: безнал с НДС или наличные.",
        "💰 Оплата обсуждается — с НДС по безналу или наличными.",
    ],
}

saved = load_settings()

# Выбор городов по регионам (миграция старого формата внутри _read_cities_by_region)
cities_by_region: Dict[str, List[str]] = _read_cities_by_region(saved)

if saved:
    CURRENT_REGION = _normalize_region(saved.get("region"))
    PAYMENT_KEY = saved.get("payment", DEFAULT_PAYMENT_KEY)
    if PAYMENT_KEY not in PAYMENT_PHRASES:
        PAYMENT_KEY = DEFAULT_PAYMENT_KEY
    logger.info(
        f"Загружены настройки пользователя: регион={CURRENT_REGION}, "
        f"городов={len(cities_by_region.get(CURRENT_REGION, []))}"
    )
else:
    CURRENT_REGION = DEFAULT_REGION
    PAYMENT_KEY = DEFAULT_PAYMENT_KEY
    logger.info("Использую настройки по умолчанию: регион=Юга")

# Регион, города которого показываются сейчас (нужен, чтобы сохранить выбор
# при переключении вкладок)
_active_region = CURRENT_REGION

if CURRENT_REGION not in cities_by_region:
    cities_by_region[CURRENT_REGION] = list(DEFAULT_CITIES_BY_REGION[CURRENT_REGION])

SELECTED_CITIES = list(cities_by_region[CURRENT_REGION])

# ============================================
# ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ДЛЯ GUI
# ============================================

city_vars = {}
notebook = None
text_preview = None

# False до конца построения окна: обработчик смены вкладки не должен
# срабатывать на программную установку вкладки при старте.
_gui_ready = {"value": False}

# ============================================
# ИСТОРИЯ ФРАЗ (чтобы текст не повторялся)
# ============================================


class RecentHistory:
    """
    Помнит последние выданные фразы и не даёт выдать ту же самую подряд.

    Зачем: random.choice из одного списка регулярно выдаёт одну и ту же фразу
    два раза кряду — со стороны это выглядит как «кнопка не работает, текст не
    меняется». Здесь у каждого списка своя история: новая фраза выбирается
    только из тех, что не выдавались последними.
    """

    def __init__(self, depth: int = 3) -> None:
        self._depth = max(1, int(depth))
        self._history: Dict[str, List[str]] = {}

    def pick(self, key: str, variants: Sequence[str]) -> str:
        """Возвращает вариант, которого не было среди последних выданных."""
        pool = [v for v in variants if v]
        if not pool:
            return ""
        if len(pool) == 1:
            return pool[0]

        recent = self._history.setdefault(key, [])
        candidates = [v for v in pool if v not in recent]
        if not candidates:
            # Все варианты уже были в недавних — начинаем круг заново.
            recent.clear()
            candidates = pool

        choice = random.choice(candidates)
        recent.append(choice)

        # История всегда короче самого списка, иначе выбор «не из недавних»
        # станет невозможным и мы будем каждый раз её чистить.
        limit = min(self._depth, max(1, len(pool) - 1))
        del recent[:-limit]
        return choice

    def clear(self) -> None:
        """Забывает всё — например, по кнопке «Новый вариант»."""
        self._history.clear()


_PHRASES = RecentHistory()

# ============================================
# ФУНКЦИИ
# ============================================


def get_selected_cities():
    """
    Возвращает список выбранных городов.

    Отметки есть только у текущего региона: при переключении вкладки галочки
    остальных регионов снимаются, а их выбор лежит в cities_by_region.
    """
    try:
        selected = [city for city, var in city_vars.items() if var.get() == 1]
        logger.debug(f"Выбрано городов: {len(selected)} - {selected}")
        return selected
    except Exception as e:
        logger.error(f"Ошибка получения выбранных городов: {e}")
        return []


def update_cities():
    """Обновляет выбранные города и сохраняет настройки"""
    logger.info("Обновление выбора городов")
    selected = get_selected_cities()
    region = get_current_region()
    if selected:
        global SELECTED_CITIES, CURRENT_REGION
        SELECTED_CITIES = selected
        CURRENT_REGION = region
        cities_by_region[region] = list(selected)
        save_settings(selected, region)
        show_preview()
        if len(selected) == 1:
            if selected[0] == CREATIVE_MODE:
                msg = "✅ Выбран креативный режим 'Без города'"
                logger.info("Выбран креативный режим")
            else:
                msg = f"✅ Выбран город:\n\n{selected[0]}"
                logger.info(f"Выбран город: {selected[0]}")
        else:
            msg = f"✅ Выбрано: {len(selected)}\n\n{', '.join(selected)}"
            logger.info(f"Выбрано {len(selected)} городов: {', '.join(selected)}")
        messagebox.showinfo("Готово!", msg)
    else:
        logger.warning("Попытка применить выбор без выбранных городов")
        messagebox.showwarning("Внимание!", "Выберите хотя бы один город!")


def get_current_region():
    """
    Определяет текущий активный регион по вкладке.

    Регион ищем по НАЗВАНИЮ вкладки, а не по её номеру: номер сбивается,
    стоит поменять порядок вкладок или добавить новую.
    """
    try:
        tab_name = notebook.get()
    except Exception:
        return CURRENT_REGION  # Возвращаем сохранённый регион
    return TAB_TITLES_BY_NAME.get(tab_name, CURRENT_REGION)


def apply_region_selection(region: str) -> None:
    """
    Показывает сохранённый выбор городов для региона.

    Для региона, который ещё не открывали, берём города целиком — так же
    вела себя прежняя версия при переключении вкладки.
    """
    if not city_vars:
        return

    selected = cities_by_region.get(region)
    if selected is None:
        selected = list(DEFAULT_CITIES_BY_REGION.get(region, []))

    chosen = set(selected)
    for city, var in city_vars.items():
        var.set(1 if city in chosen else 0)


def switch_region(event=None):
    """
    Переключает регион при смене вкладки.

    Главное отличие от прежней версии: выбор городов больше НЕ теряется.
    Перед переключением он сохраняется в cities_by_region[_active_region],
    а для нового региона восстанавливается из сохранённого.
    """
    global CURRENT_REGION, SELECTED_CITIES, _active_region

    if not _gui_ready["value"]:
        return

    region = get_current_region()
    if region == _active_region:
        return

    cities_by_region[_active_region] = get_selected_cities()
    logger.info(
        f"Регион {_active_region}: сохранён выбор городов "
        f"({len(cities_by_region[_active_region])} шт.)"
    )

    _active_region = region
    CURRENT_REGION = region
    apply_region_selection(region)

    selected = get_selected_cities()
    SELECTED_CITIES = selected
    save_settings(selected, region)
    _update_mode_label()
    show_preview()


def _on_city_toggle(city: str) -> None:
    """
    Реакция на установку галочки города.

    «Без города» — это режим, а не город: вместе со списком городов он
    противоречит сам себе. Поэтому отметка одного снимает другое.
    """
    if city == CREATIVE_MODE:
        if city_vars[CREATIVE_MODE].get() != 1:
            return
        for other, var in city_vars.items():
            if other != CREATIVE_MODE:
                var.set(0)
    else:
        if city_vars[city].get() != 1:
            return
        creative = city_vars.get(CREATIVE_MODE)
        if creative is not None:
            creative.set(0)


# ============================================
# ОФОРМЛЕНИЕ ГОРОДОВ - РАСШИРЕННЫЙ КРЕАТИВ
# ============================================


def format_route():
    """Форматирует маршрут с креативным оформлением"""
    cities = get_selected_cities()
    region = get_current_region()

    is_creative = CREATIVE_MODE in cities and len(cities) == 1

    if is_creative and region == "Восток":
        return _PHRASES.pick("route_creative", CREATIVE_ROUTE_EAST)

    if not cities:
        return "Нет выбранных городов"

    cities = [c for c in cities if c != CREATIVE_MODE]
    if not cities:
        if region == "Восток":
            return _PHRASES.pick("route_creative", CREATIVE_ROUTE_EAST)
        return "Нет выбранных городов"

    if len(cities) == 1:
        city = cities[0]
        one_city_styles = [
            f"🚛 Лот машин на автовоз: Москва → {city}!",
            f"🚛 Лот машин ждёт автовоз: Москва → {city}!",
            f"🚛 Выгрузка: {city} — ищем автовоз на этот лот!",
            f"🚛 Автовозам: лот машин Москва → {city}!",
            f"🚛 Машины ждут автовоз: Москва → {city}!",
            f"🚛 Груз для автовоза: Москва → {city}!",
            f"🚛 Лот сформирован, ищем автовоз: Москва → {city}!",
            f"🚛 Есть машины на перевозку автовозом: Москва → {city}!",
            f"🚛 Лот машин под автовоз: Москва → {city}!",
            f"🚛 Машины на отправку: Москва → {city}, нужен автовоз!",
        ]
        return _PHRASES.pick("route_one", one_city_styles)

    styles = [
        lambda: "\n".join([f"{i+1}. {city}" for i, city in enumerate(cities)]),
        lambda: "\n".join([f"{i+1}) {city}" for i, city in enumerate(cities)]),
        lambda: "\n".join([f"• {city}" for city in cities]),
        lambda: "\n".join([f"▸ {city}" for city in cities]),
        lambda: "\n".join([f"■ {city}" for city in cities]),
        lambda: "\n".join([f"✦ {city}" for city in cities]),
        lambda: "\n".join([f"{chr(1072 + i)}) {city}" for i, city in enumerate(cities)]),
        lambda: "\n".join([f"{chr(1040 + i)}) {city}" for i, city in enumerate(cities)]),
        lambda: "\n".join([f"{['I','II','III','IV','V','VI','VII','VIII','IX','X','XI','XII','XIII','XIV','XV','XVI'][i]}. {city}" for i, city in enumerate(cities)]),
        lambda: "\n".join([f"➜ {city}" for city in cities]),
        lambda: "\n".join([f"✅ {city}" for city in cities]),
        lambda: "\n".join([f"★ {city}" for city in cities]),
        lambda: ", ".join(cities),
        lambda: "\n".join([f">> {city}" for city in cities]),
        lambda: "\n".join([f"[{i+1}] {city}" for i, city in enumerate(cities)]),
        lambda: "\n".join([f"{i+1}.     {city}" for i, city in enumerate(cities)]),
    ]
    # Стиль оформления выбираем без повторов подряд: иначе нумерация списка
    # «залипает» и кажется, что новый вариант не сгенерировался.
    return _PHRASES.pick("route_style", [style() for style in styles])


# ============================================
# КРЕАТИВНЫЕ ФРАЗЫ - РАСШИРЕННЫЕ
# ============================================
#
# Специфика проекта: машины в лоте — это ГРУЗ для автовоза, а не наш транспорт.
# Поэтому в каждой фразе про машины стоит контекст «на автовоз»: без него текст
# читается как «наши авто готовы к отправке» (см. README, раздел про автовозы).
#
# Открывающая фраза зависит от количества городов в лоте: «одно направление» и
# «сборный маршрут» — разные предложения для перевозчика.

# Общие «креативные» заголовки для Востока без конкретных городов
CREATIVE_ROUTE_EAST = [
    "🚛 Лот машин на автовоз: Москва → Восточные регионы!",
    "🚛 Лот машин ждёт автовоз: Москва → Восток!",
    "🚛 Автовозам: лот машин Москва → Восточная Сибирь и ДВ!",
    "🚛 Лот машин на автовоз: Москва → Урал и Сибирь!",
    "🚛 Груз для автовоза: Москва → Урал, Сибирь, Дальний Восток!",
    "🚛 Лот машин на автовоз: Москва → Восточные города России!",
    "🚛 Лот машин ждёт автовоз: от Урала до Тихого океана!",
    "🚛 Лот машин на автовоз: Москва → Восточные регионы РФ!",
    "🚛 Груз для автовоза: Москва → Дальний Восток!",
    "🚛 Лот машин на автовоз — восточный маршрут!",
]

INTROS_SOUTH_ONE = [
    "🚛 Есть лот машин на автовоз: Москва → южный город!",
    "🚛 Свежий лот на автовоз — одно направление на юг!",
    "🚛 Лот машин ждёт автовоз: Москва → Юга, один город!",
    "🚛 Автовозам: свежий лот с Москвы на юг, один рейс!",
    "🚛 Машины на отправку — нужен автовоз на южное направление!",
    "🚛 Лот сформирован, ищем автовоз: Москва → Юга, одна точка!",
    "🚛 Груз для автовоза — Москва → южный город!",
    "🚛 Есть машины на перевозку автовозом: Москва → Юга!",
    "🚛 Новый лот машин под автовоз — один город на юге!",
    "🚛 Лот машин на юг — загрузка на автовоз, один город!",
    "🚛 Машины ждут автовоз: Москва → южное направление!",
    "🚛 Свежий лот: один южный город, ищем автовоз!",
]

INTROS_SOUTH_MULTI = [
    "🚛 Есть лот машин на автовоз: Москва → Юга, сборный маршрут!",
    "🚛 Свежий лот на автовоз — Москва → города Юга!",
    "🚛 Лот машин ждёт автовоз — Москва → Юга, несколько городов!",
    "🚛 Автовозам: сборный лот с Москвы на юг России!",
    "🚛 Машины на отправку — нужен автовоз по южному маршруту!",
    "🚛 Лот сформирован, ищем автовоз: Москва → Юга, сборный рейс!",
    "🚛 Груз для автовоза — Москва → южные города!",
    "🚛 Есть машины на перевозку автовозом: Москва → Юга, сборный!",
    "🚛 Новый лот машин под автовоз — забрать из городов Юга!",
    "🚛 Лот машин на юг — загрузка на автовоз, несколько городов!",
    "🚛 Машины ждут автовоз: Москва → города Юга!",
    "🚛 Сборный лот на юг — нужен автовоз на весь маршрут!",
]

INTROS_VLADIVOSTOK_ONE = [
    "🚛 Лот машин на автовоз: Владивосток → один город!",
    "🚛 Свежий лот на автовоз — из Владивостока в один город!",
    "🚛 Автовозам: лот машин с Владивостока, одно направление!",
    "🚛 Лот машин ждёт автовоз — Владивосток → город ниже!",
    "🚛 Машины на отправку с Владивостока — нужен автовоз!",
    "🚛 Лот сформирован, ищем автовоз: Владивосток → одна точка!",
    "🚛 Груз для автовоза — из Владивостока в один город!",
    "🚛 Есть машины на перевозку автовозом: Владивосток → город!",
    "🚛 Новый лот машин под автовоз — старт во Владивостоке!",
    "🚛 Лот машин с Владивостока — загрузка на автовоз!",
    "🚛 Машины ждут автовоз: Владивосток → одно направление!",
    "🚛 Свежий лот: один город, старт из Владивостока!",
]

INTROS_VLADIVOSTOK_MULTI = [
    "🚛 Лот машин на автовоз: Владивосток → по России, сборный!",
    "🚛 Свежий лот на автовоз — из Владивостока по городам!",
    "🚛 Автовозам: сборный лот машин с Владивостока!",
    "🚛 Лот машин ждёт автовоз — Владивосток → несколько городов!",
    "🚛 Машины на отправку с Владивостока — нужен автовоз на маршрут!",
    "🚛 Лот сформирован, ищем автовоз: Владивосток → сборный рейс!",
    "🚛 Груз для автовоза — из Владивостока по всей стране!",
    "🚛 Есть машины на перевозку автовозом: Владивосток → Россия!",
    "🚛 Новый лот машин под автовоз — старт во Владивостоке, сборный!",
    "🚛 Автовозам на заметку: лот машин с Дальнего Востока!",
    "🚛 Лот машин с Владивостока — несколько городов выгрузки!",
    "🚛 Машины ждут автовоз: Владивосток → несколько городов!",
]

INTROS_EAST_ONE = [
    "🚛 Лот машин на автовоз: Москва → восточный город!",
    "🚛 Свежий лот на автовоз — одно направление на Восток!",
    "🚛 Автовозам: лот машин на восточное направление!",
    "🚛 Лот машин ждёт автовоз — Москва → Восток, одна точка!",
    "🚛 Машины на отправку на Восток — нужен автовоз!",
    "🚛 Лот сформирован, ищем автовоз: Москва → восточный город!",
    "🚛 Груз для автовоза — Москва → Урал или Сибирь!",
    "🚛 Есть машины на перевозку автовозом на Восток!",
    "🚛 Новый лот машин под автовоз — восточный маршрут!",
    "🚛 Лот машин на восточный город — загрузка на автовоз!",
    "🚛 Машины ждут автовоз: Москва → Восток, один город!",
    "🚛 Свежий лот: один восточный город, ищем автовоз!",
]

INTROS_EAST_MULTI = [
    "🚛 Лот машин на автовоз: Москва → Восток, сборный маршрут!",
    "🚛 Свежий лот на автовоз — Москва → Урал и Сибирь!",
    "🚛 Автовозам: сборный лот на восточное направление!",
    "🚛 Лот машин ждёт автовоз — Москва → Восток, несколько городов!",
    "🚛 Машины на отправку на Восток — нужен автовоз на маршрут!",
    "🚛 Лот сформирован, ищем автовоз: Москва → Восток, сборный!",
    "🚛 Груз для автовоза — Москва → Урал → Сибирь → ДВ!",
    "🚛 Есть машины на перевозку автовозом на Восток, сборный лот!",
    "🚛 Новый лот машин под автовоз — несколько городов Востока!",
    "🚛 Автовозам на заметку: лот машин с Москвы на Восток!",
    "🚛 Лот машин на восточные регионы — загрузка на автовоз!",
    "🚛 Машины ждут автовоз: Москва → города Востока!",
]

CARGO_SOUTH = [
    "🚗 Машины готовы к погрузке на автовоз.",
    "🚗 Лот сформирован, машины ждут автовоз.",
    "🚗 Автомобили для перевозки — ищем автовоз.",
    "🚗 Есть машины на перевозку автовозом.",
    "🚗 Машины ждут загрузки на автовоз.",
    "🚗 Лот машин ждёт автовоз — загрузка быстрая.",
    "🚗 Машины на площадке, нужен автовоз под погрузку.",
    "🚗 Машины ждут своей очереди на автовоз.",
    "🚗 Лот машин сформирован — ищем автовоз.",
    "🚗 Машины на площадке и готовы к погрузке, ищем автовоз.",
    "🚗 Машины к отправке готовы, нужен автовоз на рейс.",
    "🚗 Лот машин собран, автовоз требуется в ближайшие дни.",
]

PAYMENT_SOUTH = [
    "💰 Оплата по безналу (с НДС), сроки обсуждаем.",
    "💰 Оплата по безналу (с НДС), возможна предоплата.",
    "💰 Безналичный расчёт с НДС, гибкие условия.",
    "💰 Оплата по безналу (с НДС), работаем с юрлицами.",
    "💰 Оплата по безналу (с НДС), договор сразу.",
    "💰 Оплата по безналу (с НДС), закрывающие документы.",
    "💰 Оплата по безналу (с НДС), сроки подберём.",
]

LEGAL_SOUTH = [
    "📋 Работаем по договору, платим без задержек.",
    "📄 Заключаем договор, выплаты строго в срок, без просрочек.",
    "🤝 Официальное оформление, оплата вовремя, гарантируем надёжность.",
    "📑 Работаем официально, платежи без задержек, договор гарантирует защиту.",
    "📋 Все сделки по договору, выплаты чётко в срок, без задержек.",
    "📄 Юридически чистая работа — договор, оплата по факту, без просрочек.",
    "⚖️ Полное юридическое сопровождение — ваша безопасность в приоритете!",
    "📑 Закрывающие документы предоставляем в срок.",
    "🤝 Договор оформляем сразу, оплата в срок.",
    "📋 Прозрачные условия: договор, документы, оплата в срок.",
    "📄 Работаем официально, все документы предоставляем.",
    "⚖️ Работаем прозрачно: договор, счёт, закрывающие.",
]

CTA_SOUTH = [
    "✍️ Все детали и ставки — в личку. Пишите! 📩",
    "📩 Ждём ваших сообщений для обсуждения деталей и ставок!",
    "💬 Пишите в личные сообщения — обсудим все условия и ставки!",
    "📲 Свяжитесь с нами в личных сообщениях — пришлём ставки!",
    "✍️ Детали и стоимость — в личку. Ждём ваших вопросов!",
    "📩 Обсудим все нюансы и ставки в личных сообщениях!",
    "📱 Ждём ваших сообщений — готовы ответить на все вопросы!",
    "✍️ Ставки и даты по лоту — в личку, отвечаем быстро.",
    "💬 Пишите в личку — рассчитаем ставку под ваш рейс.",
    "📲 Есть свободный автовоз? Напишите — дадим детали.",
    "📩 Детали по лоту: города, даты, ставка — в личке.",
    "📱 Пишите в личку, если готовы взять рейс.",
]

CARGO_VLADIVOSTOK = [
    "🚗 Машины готовы к погрузке на автовоз.",
    "🚗 Лот сформирован, машины ждут автовоз.",
    "🚗 Автомобили для перевозки — ищем автовоз.",
    "🚗 Есть машины на перевозку автовозом.",
    "🚗 Машины ждут загрузки на автовоз.",
    "🚗 Лот машин ждёт автовоз — погрузка быстрая.",
    "🚗 Машины на площадке, автовоз нужен сегодня.",
    "🚗 Машины к отправке готовы, нужен автовоз на рейс.",
    "🚗 Лот собран, машины стоят под погрузку на автовоз.",
    "🚗 Ждём автовоз: машины на площадке, погрузка без ожидания.",
    "🚗 Машины ждут автовоз, погрузка займёт минимум времени.",
    "🚗 Лот машин ждёт автовоз — загрузка без простоев.",
]

PAYMENT_VLADIVOSTOK = [
    "💰 Оплата наличными (без НДС) — 50/50 при погрузке.",
    "💰 Оплата наличными (без НДС) — 50% при загрузке, 50% по факту.",
    "💰 Оплата по безналу (с НДС) — возможна предоплата.",
    "💰 Оплата по безналу (с НДС), работаем с юрлицами.",
    "💰 Условия оплаты обсуждаются — наличные или безнал с НДС.",
]

LEGAL_VLADIVOSTOK = [
    "📦 Работаем чисто, без задержек. Подробности по каждому лоту — в личные сообщения.",
    "📋 Работаем по договору, все сделки прозрачны.",
    "📄 Гарантируем чистоту сделок, работаем без задержек.",
    "🤝 Надёжные партнёры — оплата вовремя, без задержек.",
    "📑 Полная юридическая прозрачность — ваше спокойствие гарантируем!",
    "⚖️ Работаем по договору, документы предоставляем в срок.",
    "🧾 Оплата в срок, закрывающие документы — по договору.",
    "📋 Оформление официальное, реквизиты и договор сразу.",
    "📄 Все условия в договоре, оплата без просрочек.",
    "🤝 Договор и полный пакет документов — работаем чисто.",
    "⚖️ Юридическое сопровождение сделки — работаем официально.",
    "📑 Закрывающие документы в срок, работаем официально.",
]

CTA_VLADIVOSTOK = [
    "✍️ Пишите — обсудим ставки, даты и условия!",
    "📩 Обсудим детали в личных сообщениях — ждём ваших вопросов!",
    "💬 Свяжитесь с нами для уточнения ставок и дат!",
    "📲 Пишите — договоримся о всех условиях!",
    "📱 Ждём вашего сообщения — готовы к диалогу!",
    "✍️ Ставка и даты — в личке, отвечаем оперативно.",
    "💬 Напишите в личку — рассчитаем рейс под ваш автовоз.",
    "📩 Детали по лоту — в личных сообщениях.",
    "📲 Пишите, если готовы взять рейс из Владивостока.",
    "📱 Обсудим ставку в личке — пишите.",
    "✍️ Все подробности по лоту — в личке, ждём!",
    "💬 Есть вопросы по погрузке? Пишите в личку.",
]

CARGO_EAST = [
    "🚗 Машины готовы к погрузке на автовоз.",
    "🚗 Лот сформирован, машины ждут автовоз.",
    "🚗 Автомобили для перевозки — ищем автовоз.",
    "🚗 Есть машины на перевозку автовозом.",
    "🚗 Машины ждут загрузки на автовоз.",
    "🚗 Машины на площадке — нужен автовоз под погрузку.",
    "🚗 Лот машин ждёт автовоз, загрузка быстрая.",
    "🚗 Машины ждут своей очереди на автовоз.",
    "🚗 Машины на площадке, автовоз нужен сегодня.",
    "🚗 Лот готов: машины стоят под погрузку на автовоз.",
    "🚗 Машины к отправке готовы, нужен автовоз на рейс.",
    "🚗 Лот машин собран, автовоз требуется в ближайшие дни.",
]

PAYMENT_EAST = [
    "💰 Оплата по безналу (с НДС), сроки обсуждаем.",
    "💰 Оплата по безналу (с НДС), гибкие условия.",
    "💰 Безналичный расчёт с НДС — быстро и прозрачно.",
    "💰 Оплата по безналу (с НДС), условия под каждого.",
    "💰 Условия оплаты обсуждаются — наличные или безнал с НДС.",
]

LEGAL_EAST = [
    "📋 Работаем по договору, выплаты в срок.",
    "📄 Заключаем договор — гарантия надёжности.",
    "🤝 Официальное оформление, оплата вовремя.",
    "📑 Все сделки по договору, без задержек.",
    "⚖️ Полное юридическое сопровождение сделки!",
    "🧾 Закрывающие документы предоставляем в срок.",
    "📋 Прозрачные условия: договор, документы, оплата в срок.",
    "📄 Работаем официально, платежи без задержек.",
    "🤝 Договор оформляем сразу, оплата по факту выгрузки.",
    "📑 Оплата в срок по договору, без просрочек.",
    "⚖️ Работаем прозрачно: договор, счёт, закрывающие.",
    "📋 Юридически чистая сделка — договор и документы.",
]

CTA_EAST = [
    "✍️ Детали и ставки — в ЛС. Ждём ваших сообщений!",
    "📩 Обсудим условия и ставки в личных сообщениях.",
    "💬 Пишите — рассчитаем оптимальные условия!",
    "📲 Свяжитесь с нами в личных сообщениях — уточним детали и ставки!",
    "📱 Ждём вашего звонка или сообщения!",
    "✍️ Ставки и даты — в личку, отвечаем быстро.",
    "💬 Напишите в личку — рассчитаем ставку под ваш рейс.",
    "📩 Детали по лоту: города, даты, ставка — в личке.",
    "📲 Есть свободный автовоз на Восток? Пишите!",
    "📱 Пишите в личку, если готовы взять рейс.",
    "✍️ Обсудим ставку в личных сообщениях — пишите.",
    "💬 Все подробности по лоту — в личке, ждём!",
]

# ============================================
# ФУНКЦИЯ ГЕНЕРАЦИИ
# ============================================


def payment_phrase():
    """Фраза про оплату по выбранному в интерфейсе типу (с НДС / без НДС / наличные / обсуждается)."""
    phrases = PAYMENT_PHRASES.get(PAYMENT_KEY, PAYMENT_PHRASES[DEFAULT_PAYMENT_KEY])
    return _PHRASES.pick(f"payment_{PAYMENT_KEY}", phrases)


def get_payment_key():
    """Читает тип оплаты из выпадающего списка. До создания виджета — сохранённое значение."""
    try:
        label = payment_var.get()
    except NameError:
        return PAYMENT_KEY
    for key, shown in PAYMENT_LABELS_MAP.items():
        if shown == label:
            return key
    return PAYMENT_KEY


def update_payment(_event=None):
    """Обработчик выбора оплаты: сохраняет настройку и обновляет предпросмотр."""
    global PAYMENT_KEY
    PAYMENT_KEY = get_payment_key()
    logger.info(f"Выбран тип оплаты: {PAYMENT_KEY}")
    save_settings(get_selected_cities(), get_current_region(), PAYMENT_KEY)
    _update_mode_label()
    show_preview()


def _update_mode_label() -> None:
    """Показывает над предпросмотром, для какого региона и оплаты собран текст."""
    try:
        mode_label.configure(
            text=(
                f"Регион: {get_current_region()} · "
                f"оплата: {PAYMENT_LABELS_MAP.get(PAYMENT_KEY, '—')}"
            )
        )
    except NameError:
        # Виджета ещё нет (ранний вызов) — это нормально
        pass


def _region_phrases(region: str) -> Dict[str, List[str]]:
    """Наборы фраз для региона: свой список на каждый регион."""
    if region == "Владивосток":
        return {
            "cargo": CARGO_VLADIVOSTOK,
            "legal": LEGAL_VLADIVOSTOK,
            "cta": CTA_VLADIVOSTOK,
        }
    if region == "Восток":
        return {
            "cargo": CARGO_EAST,
            "legal": LEGAL_EAST,
            "cta": CTA_EAST,
        }
    return {
        "cargo": CARGO_SOUTH,
        "legal": LEGAL_SOUTH,
        "cta": CTA_SOUTH,
    }


def _pick_intro(region: str, city_count: int) -> str:
    """
    Открывающая фраза региона с учётом количества городов.

    Для одного города — «одно направление», для нескольких — «сборный
    маршрут». Это разные предложения для перевозчика, и раньше текст об этом
    умалчивал.
    """
    if region == "Владивосток":
        one, many, key = INTROS_VLADIVOSTOK_ONE, INTROS_VLADIVOSTOK_MULTI, "vl"
    elif region == "Восток":
        one, many, key = INTROS_EAST_ONE, INTROS_EAST_MULTI, "east"
    else:
        one, many, key = INTROS_SOUTH_ONE, INTROS_SOUTH_MULTI, "south"

    if city_count <= 1:
        return _PHRASES.pick(f"intro_{key}_one", one)
    return _PHRASES.pick(f"intro_{key}_multi", many)


def generate_creative_text():
    """Генерирует креативный текст коммерческого предложения"""
    region = get_current_region()
    route = format_route()
    logger.debug(f"Генерация текста для региона: {region}")

    route_cities = [c for c in get_selected_cities() if c != CREATIVE_MODE]
    phrases = _region_phrases(region)

    intro = _pick_intro(region, len(route_cities))
    cargo = _PHRASES.pick(f"cargo_{region}", phrases["cargo"])
    legal = _PHRASES.pick(f"legal_{region}", phrases["legal"])
    cta = _PHRASES.pick(f"cta_{region}", phrases["cta"])

    payment = payment_phrase()

    result = f"{intro}\n\n{route}\n\n{cargo}\n{payment}\n{legal}\n\n{cta}"
    logger.debug("Текст успешно сгенерирован")
    return result


# ============================================
# КНОПКИ
# ============================================


def _documents_dir() -> str:
    """
    Возвращает папку для сохранения Word-файлов.

    Обычно это рабочий стол. Но его может не быть: OneDrive перенёс папку,
    профиль ограничен, или система без рабочего стола — тогда сохранение
    раньше падало с FileNotFoundError.
    """
    home = os.path.expanduser("~")
    candidates = [
        os.path.join(home, "Desktop"),
        os.path.join(home, "OneDrive", "Desktop"),
        os.path.join(home, "Рабочий стол"),
        os.path.join(home, "Documents"),
        os.path.join(home, "Документы"),
        home,
    ]
    for path in candidates:
        try:
            if os.path.isdir(path):
                return path
        except OSError:
            continue
    return home


def _save_text_as_docx(text: str, filename: str) -> str:
    """Создаёт Word-файл с готовым текстом КП и возвращает путь к нему."""
    doc = Document()
    title = doc.add_heading("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_paragraph = doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y')}")
    date_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph()
    for line in text.split("\n"):
        if line.strip():
            doc.add_paragraph(line)

    filepath = os.path.join(_documents_dir(), filename)
    try:
        doc.save(filepath)
    except Exception:
        # Папка недоступна на запись — сохраняем рядом с программой
        filepath = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        logger.warning("Не удалось сохранить на рабочий стол — сохраняю в %s", filepath)
        doc.save(filepath)
    return filepath


def copy_to_clipboard():
    """Копирует сгенерированный текст в буфер обмена"""
    logger.info("Копирование текста в буфер обмена")
    try:
        text = generate_creative_text()
        pyperclip.copy(text)
        logger.debug("Текст скопирован в буфер обмена")
        messagebox.showinfo("Готово!", "✅ Текст скопирован в буфер обмена!")
    except Exception as e:
        logger.error(f"Ошибка копирования в буфер: {e}")
        messagebox.showerror("Ошибка", f"❌ {str(e)}")


def save_to_word():
    """Сохраняет сгенерированный текст в Word документ"""
    logger.info("Сохранение в Word")
    try:
        text = generate_creative_text()
        region = get_current_region()
        filename = f"КП_{region}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        filepath = _save_text_as_docx(text, filename)
        logger.info(f"Word-файл сохранён: {filepath}")
        messagebox.showinfo("Готово!", f"✅ Word-файл создан:\n\n{filepath}")
    except Exception as e:
        logger.error(f"Ошибка сохранения в Word: {e}")
        messagebox.showerror("Ошибка", f"❌ {str(e)}")


def save_multiple_variants():
    """Создаёт 3 разных варианта в Word"""
    logger.info("Создание 3 вариантов в Word")
    try:
        region = get_current_region()
        saved = []
        for i in range(3):
            text = generate_creative_text()
            filename = (
                f"КП_{region}_вариант_{i+1}_"
                f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            )
            saved.append(_save_text_as_docx(text, filename))
            logger.debug(f"Вариант {i+1} сохранён: {saved[-1]}")
        folder = os.path.dirname(saved[0]) if saved else ""
        messagebox.showinfo(
            "Готово!", f"✅ 3 файла созданы:\n\nПапка:\n{folder}"
        )
    except Exception as e:
        logger.error(f"Ошибка создания 3 вариантов: {e}")
        messagebox.showerror("Ошибка", f"❌ {str(e)}")


def show_preview():
    """Обновляет предпросмотр текста"""
    logger.debug("Обновление предпросмотра")
    try:
        text = generate_creative_text()
        text_preview.configure(state="normal")
        text_preview.delete("1.0", "end")
        text_preview.insert("1.0", text)
        text_preview.configure(state="disabled")
    except Exception as e:
        logger.error(f"Ошибка обновления предпросмотра: {e}")


def new_variant():
    """
    Кнопка «Показать новый вариант».

    Кроме перегенерации текста сбрасываем историю фраз: пользователь ждёт
    другой текст, а не повтор уже показанных строк.
    """
    _PHRASES.clear()
    show_preview()


def create_installer():
    """Создаёт установщик с помощью Inno Setup"""
    logger.info("Запуск создания установщика")
    try:
        import subprocess
        import shutil

        current_dir = os.path.dirname(os.path.abspath(__file__))
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        messagebox.showinfo("Установщик", "🔨 Сборка...\n\nПодождите 2-3 минуты.")
        subprocess.run(
            [
                "pyinstaller",
                "--onefile",
                "--console",
                "--name",
                "Генератор_КП",
                os.path.join(current_dir, "генератор_кп.py"),
            ],
            check=True,
        )
        exe_source = os.path.join(current_dir, "dist", "Генератор_КП.exe")
        setup_folder = os.path.join(desktop, "Setup")
        if not os.path.exists(setup_folder):
            os.makedirs(setup_folder)
        shutil.copy(exe_source, os.path.join(setup_folder, "Генератор_КП.exe"))
        iss_content = '''[Setup]
AppName=Генератор КП
AppVersion=1.0
DefaultDirName={pf}\\Генератор КП
DefaultGroupName=Генератор КП
OutputDir=.
OutputBaseFilename=Генератор_КП_Установщик
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Files]
Source: "Генератор_КП.exe"; DestDir: "{app}"

[Icons]
Name: "{group}\\Генератор КП"; Filename: "{app}\\Генератор_КП.exe"
Name: "{commondesktop}\\Генератор КП"; Filename: "{app}\\Генератор_КП.exe"

[Run]
Filename: "{app}\\Генератор_КП.exe"; Description: "Запустить Генератор КП"; Flags: postinstall nowait skipifsilent'''
        with open(os.path.join(setup_folder, "setup.iss"), "w", encoding="utf-8") as f:
            f.write(iss_content)
        os.startfile(setup_folder)
        logger.info("Папка Setup создана на рабочем столе")
        messagebox.showinfo(
            "Готово!",
            "✅ Папка Setup создана на рабочем столе!\n\n"
            "Откройте setup.iss в Inno Setup и нажмите F9",
        )
    except Exception as e:
        logger.error(f"Ошибка создания установщика: {e}")
        messagebox.showerror(
            "Ошибка", f"❌ {str(e)}\n\nУстановите PyInstaller:\npip install pyinstaller"
        )


def _open_log_folder() -> None:
    """Открывает папку с логами в проводнике."""
    try:
        folder = os.path.dirname(str(LOG_FILE)) or os.path.dirname(
            os.path.abspath(__file__)
        )
        if not os.path.isdir(folder):
            folder = os.path.dirname(os.path.abspath(__file__))
        os.startfile(folder)
    except Exception as exc:
        logger.warning("Не удалось открыть папку с логами: %s", exc)
        messagebox.showinfo("Логи", f"Файл лога лежит здесь:\n\n{LOG_FILE}")


def on_closing():
    """Обработчик закрытия окна - сохраняет настройки"""
    logger.info("Закрытие программы - сохранение настроек")
    try:
        # Сохраняем текущие настройки перед выходом
        selected = get_selected_cities()
        region = get_current_region()

        logger.info(f"Сохраняю перед выходом: регион={region}, городов={len(selected)}")

        if selected:
            success = save_settings(selected, region)
            if success:
                logger.info("Настройки успешно сохранены перед выходом")
            else:
                logger.error("Не удалось сохранить настройки перед выходом")
        else:
            # Если ничего не выбрано, сохраняем последние выбранные
            if SELECTED_CITIES:
                logger.info(f"Сохраняю последние выбранные города: {SELECTED_CITIES}")
                save_settings(SELECTED_CITIES, CURRENT_REGION)
            else:
                logger.warning("Нет городов для сохранения")
    except Exception as e:
        logger.error(f"Ошибка сохранения настроек при выходе: {e}")

    window.destroy()
    logger.info("Программа завершена")


# ============================================
# ГРАФИЧЕСКИЙ ИНТЕРФЕЙС (CustomTkinter)
# ============================================
#
# Раскладка та же, что в версии с GigaChat:
#   ┌──────────────────────────────────────────────────────────┐
#   │ шапка: заголовок + панель инструментов                   │
#   ├───────────────────┬──────────────────────────────────────┤
#   │ настройки (380px) │ предпросмотр (растягивается)         │
#   │ вкладки регионов  │ «показать новый вариант»             │
#   │ города, оплата    │ текст КП                             │
#   │ «применить»       │ копировать / Word / 3 варианта       │
#   ├───────────────────┴──────────────────────────────────────┤
#   │ статус-бар: режим работы + путь к логу                   │
#   └──────────────────────────────────────────────────────────┘

logger.info("Создание графического интерфейса (CustomTkinter)")


def _fit_window_to_screen(
    win: "ctk.CTk",
    want_width: int,
    want_height: int,
    min_width: int,
    min_height: int,
) -> None:
    """
    Подгоняет размер окна под экран и ставит его по центру.

    Зачем: на ноутбуке 1366x768 окно 1180x820 выше экрана — нижний край
    (кнопки и статус-бар) уходил под панель задач. Высота берётся с запасом
    88 % от экрана: остальное занимает панель задач.
    """
    try:
        screen_w = win.winfo_screenwidth()
        screen_h = win.winfo_screenheight()
        limit_w = int(screen_w * 0.95)
        limit_h = int(screen_h * 0.88)

        width = min(want_width, limit_w) if limit_w >= min_width else limit_w
        height = min(want_height, limit_h) if limit_h >= min_height else limit_h
        width = max(width, 900)
        height = max(height, 560)

        x = max((screen_w - width) // 2, 0)
        y = max((screen_h - height) // 3, 0)

        win.geometry(f"{width}x{height}+{x}+{y}")
        win.minsize(min(min_width, width), min(min_height, height))
        logger.info("Окно: %dx%d (экран %dx%d)", width, height, screen_w, screen_h)
    except Exception as exc:  # pragma: no cover — размер окна не критичен
        logger.debug("Не удалось подогнать окно под экран: %s", exc)
        win.geometry(f"{want_width}x{want_height}")


window = ctk.CTk()
window.title("Генератор КП — Автоперевозки")
window.configure(fg_color=BG_WINDOW)
_fit_window_to_screen(window, 1180, 820, 1000, 660)

# Обработчик закрытия окна
window.protocol("WM_DELETE_WINDOW", on_closing)

FONT_H1 = ctk.CTkFont(family="Segoe UI", size=22, weight="bold")
FONT_H2 = ctk.CTkFont(family="Segoe UI", size=15, weight="bold")
FONT_UI = ctk.CTkFont(family="Segoe UI", size=14)
FONT_SMALL_BOLD = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
FONT_SMALL = ctk.CTkFont(family="Segoe UI", size=12)
FONT_TINY = ctk.CTkFont(family="Segoe UI", size=11)

window.grid_columnconfigure(0, weight=1)
window.grid_rowconfigure(1, weight=1)  # растягивается тело окна

# ---------- Шапка ----------
# Заголовок и панель инструментов — в две строки: в одну они не помещаются,
# заголовок при этом обрезается по краям.
header_frame = ctk.CTkFrame(window, corner_radius=0, fg_color=BG_HEADER)
header_frame.grid(row=0, column=0, sticky="ew")
header_frame.grid_columnconfigure(0, weight=1)

title_box = ctk.CTkFrame(header_frame, fg_color="transparent")
title_box.grid(row=0, column=0, sticky="w", padx=18, pady=(12, 0))

ctk.CTkLabel(
    title_box,
    text="🚛 ГЕНЕРАТОР КОММЕРЧЕСКИХ ПРЕДЛОЖЕНИЙ",
    font=FONT_H1,
    text_color=TEXT_MAIN,
).pack(anchor="w")

ctk.CTkLabel(
    title_box,
    text="Шаблонная версия: текст собирается из готовых фраз, интернет не нужен",
    font=FONT_TINY,
    text_color=TEXT_DIM,
).pack(anchor="w")

toolbar = ctk.CTkFrame(header_frame, fg_color="transparent")
toolbar.grid(row=1, column=0, sticky="e", padx=16, pady=(4, 10))


def _tool_button(text: str, command, color: str = COLOR_TOOL, hover: str = COLOR_TOOL_HOVER):
    """Кнопка панели инструментов одного размера — шапка не «плывёт»."""
    button = ctk.CTkButton(
        toolbar,
        text=text,
        font=FONT_SMALL,
        height=32,
        fg_color=color,
        hover_color=hover,
        command=command,
    )
    button.pack(side="left", padx=3)
    return button


_tool_button("🛠 Установщик", create_installer, COLOR_WARN, COLOR_WARN_HOVER)
_tool_button("📂 Логи", _open_log_folder)
_tool_button("✖ Выход", on_closing, COLOR_DANGER, COLOR_DANGER_HOVER)

# ---------- Тело окна: две колонки ----------
body = ctk.CTkFrame(window, fg_color="transparent")
body.grid(row=1, column=0, sticky="nsew", padx=12, pady=(10, 6))
body.grid_rowconfigure(0, weight=1)
# Ширину левой колонки держит minsize: grid_propagate(False) у
# CustomTkinter-фрейма перестаёт работать, как только внутри появляются виджеты.
body.grid_columnconfigure(0, weight=0, minsize=LEFT_COLUMN_WIDTH)
body.grid_columnconfigure(1, weight=1)

# ---------- Левая колонка: настройки ----------
left_column = ctk.CTkFrame(body, corner_radius=12, fg_color=BG_CARD)
left_column.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

ctk.CTkLabel(
    left_column,
    text="⚙️ НАСТРОЙКИ ЛОТА",
    font=FONT_H2,
    text_color=TEXT_MAIN,
).pack(anchor="w", padx=16, pady=(14, 4))

notebook = ctk.CTkTabview(
    left_column,
    command=switch_region,
    width=LEFT_COLUMN_WIDTH - 40,
    # Высота заведомо небольшая: список городов прокручивается, а окно на
    # маленьком экране должно помещаться целиком (растёт по pack).
    height=150,
    fg_color=BG_WINDOW,
    segmented_button_selected_color=COLOR_PRIMARY,
    segmented_button_selected_hover_color=COLOR_PRIMARY_HOVER,
)
notebook.pack(fill="both", expand=True, padx=12, pady=(0, 8))

tab_south = notebook.add(TAB_TITLES["Юга"])
tab_vladivostok = notebook.add(TAB_TITLES["Владивосток"])
tab_east = notebook.add(TAB_TITLES["Восток"])


def _make_city_checkbox(parent, city: str, italic: bool = False) -> None:
    """Один флажок города. Значения выставляет apply_region_selection()."""
    var = tk.IntVar(value=0)
    city_vars[city] = var
    ctk.CTkCheckBox(
        parent,
        text=city,
        variable=var,
        onvalue=1,
        offvalue=0,
        font=FONT_SMALL,
        text_color="#ff8a65" if italic else TEXT_MAIN,
        fg_color=COLOR_PRIMARY,
        hover_color=COLOR_PRIMARY_HOVER,
        command=lambda name=city: _on_city_toggle(name),
    ).pack(anchor="w", padx=6, pady=3)


# ===== ЮГА =====
south_list = ctk.CTkScrollableFrame(tab_south, fg_color="transparent")
south_list.pack(fill="both", expand=True)
ctk.CTkLabel(
    south_list,
    text="📍 Выберите города для маршрута:",
    font=FONT_SMALL_BOLD,
    text_color=TEXT_MAIN,
).pack(anchor="w", padx=6, pady=(2, 4))
for _city in ALL_CITIES_SOUTH:
    _make_city_checkbox(south_list, _city)

# ===== ВЛАДИВОСТОК =====
vladivostok_list = ctk.CTkScrollableFrame(tab_vladivostok, fg_color="transparent")
vladivostok_list.pack(fill="both", expand=True)
ctk.CTkLabel(
    vladivostok_list,
    text="📍 Выберите направления:",
    font=FONT_SMALL_BOLD,
    text_color=TEXT_MAIN,
).pack(anchor="w", padx=6, pady=(2, 4))
for _city in ALL_CITIES_VLADIVOSTOK:
    _make_city_checkbox(vladivostok_list, _city)

# ===== ВОСТОК =====
east_list = ctk.CTkScrollableFrame(tab_east, fg_color="transparent")
east_list.pack(fill="both", expand=True)
ctk.CTkLabel(
    east_list,
    text="📍 Выберите города (в порядке следования):",
    font=FONT_SMALL_BOLD,
    text_color=TEXT_MAIN,
).pack(anchor="w", padx=6, pady=(2, 4))

_make_city_checkbox(east_list, CREATIVE_MODE, italic=True)

ctk.CTkLabel(
    east_list,
    text="«Без города» — креативный режим: города в текст не подставляются.",
    font=FONT_TINY,
    text_color=TEXT_DIM,
    wraplength=LEFT_COLUMN_WIDTH - 90,
    justify="left",
).pack(anchor="w", padx=6, pady=(0, 6))

for _city in ALL_CITIES_EAST:
    _make_city_checkbox(east_list, _city)

# ===== ТИП ОПЛАТЫ =====
payment_frame = ctk.CTkFrame(left_column, fg_color="transparent")
payment_frame.pack(fill="x", padx=14, pady=(2, 6))

ctk.CTkLabel(
    payment_frame,
    text="💰 Оплата:",
    font=FONT_SMALL_BOLD,
    text_color=TEXT_MAIN,
).pack(side="left", padx=(0, 8))

payment_var = tk.StringVar(
    value=PAYMENT_LABELS_MAP.get(PAYMENT_KEY, PAYMENT_LABELS_MAP[DEFAULT_PAYMENT_KEY])
)
payment_combo = ctk.CTkOptionMenu(
    payment_frame,
    values=PAYMENT_LABELS,
    variable=payment_var,
    command=update_payment,
    width=170,
    height=30,
    font=FONT_SMALL,
    fg_color=COLOR_TOOL,
    button_color=COLOR_TOOL_HOVER,
    button_hover_color=COLOR_TOOL,
    dropdown_font=FONT_SMALL,
)
payment_combo.pack(side="left")

# В шаблонной версии нет ИИ, поэтому нет и слайдера креативности: этот
# параметр относится только к GigaChat. Вместо него — пояснение, чтобы
# левая колонка не выглядела «недоделанной».
ctk.CTkLabel(
    left_column,
    text=(
        "🧩 Версия без ИИ: текст собирается из готовых фраз.\n"
        "Каждый вариант — новая комбинация формулировок."
    ),
    font=FONT_TINY,
    text_color=TEXT_DIM,
    wraplength=LEFT_COLUMN_WIDTH - 60,
    justify="left",
).pack(anchor="w", padx=16, pady=(2, 6))

ctk.CTkButton(
    left_column,
    text="✅ ПРИМЕНИТЬ ВЫБОР ГОРОДОВ",
    font=FONT_SMALL_BOLD,
    height=40,
    fg_color=COLOR_PRIMARY,
    hover_color=COLOR_PRIMARY_HOVER,
    command=update_cities,
).pack(fill="x", padx=14, pady=(8, 14))

# ---------- Правая колонка: предпросмотр ----------
right_column = ctk.CTkFrame(body, corner_radius=12, fg_color=BG_CARD)
right_column.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
right_column.grid_rowconfigure(1, weight=1)
right_column.grid_columnconfigure(0, weight=1)

preview_top = ctk.CTkFrame(right_column, fg_color="transparent")
preview_top.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 4))

ctk.CTkButton(
    preview_top,
    text="🔄 ПОКАЗАТЬ НОВЫЙ ВАРИАНТ",
    font=FONT_SMALL_BOLD,
    height=42,
    width=260,
    fg_color=COLOR_VIOLET,
    hover_color=COLOR_VIOLET_HOVER,
    command=new_variant,
).pack(side="left")

mode_label = ctk.CTkLabel(
    preview_top,
    text=f"Регион: {CURRENT_REGION} · оплата: {PAYMENT_LABELS_MAP[PAYMENT_KEY]}",
    font=FONT_SMALL,
    text_color=TEXT_DIM,
)
mode_label.pack(side="left", padx=12)

text_preview = ctk.CTkTextbox(
    right_column,
    wrap="word",
    font=FONT_UI,
    height=140,
    fg_color="#0f1216",
    text_color=TEXT_MAIN,
    corner_radius=10,
)
text_preview.grid(row=1, column=0, sticky="nsew", padx=16, pady=(8, 4))
text_preview.configure(state="disabled")

ctk.CTkLabel(
    right_column,
    text="📌 Для Востока: отметьте «Без города» — маршрут будет описан направлением, без перечисления городов",
    font=FONT_TINY,
    text_color=TEXT_DIM,
).grid(row=2, column=0, sticky="w", padx=16, pady=(0, 4))

button_frame = ctk.CTkFrame(right_column, fg_color="transparent")
button_frame.grid(row=3, column=0, sticky="ew", padx=16, pady=(4, 16))
button_frame.grid_columnconfigure(0, weight=1)
button_frame.grid_columnconfigure(1, weight=1)
button_frame.grid_columnconfigure(2, weight=1)

ctk.CTkButton(
    button_frame,
    text="📋 КОПИРОВАТЬ В БУФЕР",
    font=FONT_SMALL_BOLD,
    height=46,
    fg_color=COLOR_PRIMARY,
    hover_color=COLOR_PRIMARY_HOVER,
    command=copy_to_clipboard,
).grid(row=0, column=0, sticky="ew", padx=(0, 6))

ctk.CTkButton(
    button_frame,
    text="📄 СОХРАНИТЬ В WORD",
    font=FONT_SMALL_BOLD,
    height=46,
    fg_color=COLOR_ACCENT,
    hover_color=COLOR_ACCENT_HOVER,
    command=save_to_word,
).grid(row=0, column=1, sticky="ew", padx=6)

ctk.CTkButton(
    button_frame,
    text="📚 3 РАЗНЫХ ВАРИАНТА",
    font=FONT_SMALL_BOLD,
    height=46,
    fg_color=COLOR_WARN,
    hover_color=COLOR_WARN_HOVER,
    command=save_multiple_variants,
).grid(row=0, column=2, sticky="ew", padx=(6, 0))

# ---------- Статус-бар ----------
status_bar = ctk.CTkFrame(window, corner_radius=0, fg_color=BG_STATUS)
status_bar.grid(row=2, column=0, sticky="ew")
status_bar.grid_columnconfigure(0, weight=1)

ctk.CTkLabel(
    status_bar,
    text="🧩 Шаблонный режим: ИИ не используется, ключ и интернет не нужны",
    font=FONT_SMALL,
    text_color="#66bb6a",
).grid(row=0, column=0, sticky="w", padx=14, pady=6)

log_box = ctk.CTkFrame(status_bar, fg_color="transparent")
log_box.grid(row=0, column=1, sticky="e", padx=14, pady=6)

ctk.CTkLabel(
    log_box,
    text=f"📄 Лог: {LOG_FILE}",
    font=FONT_TINY,
    text_color=TEXT_DIM,
).pack(side="left", padx=(0, 8))

ctk.CTkButton(
    log_box,
    text="Открыть папку",
    font=FONT_TINY,
    width=110,
    height=24,
    fg_color=COLOR_TOOL,
    hover_color=COLOR_TOOL_HOVER,
    command=_open_log_folder,
).pack(side="left")

# Показываем сохранённый регион и его выбор городов, затем первый предпросмотр
notebook.set(TAB_TITLES[CURRENT_REGION])
apply_region_selection(CURRENT_REGION)
_active_region = CURRENT_REGION
_gui_ready["value"] = True

show_preview()
logger.info("Графический интерфейс загружен")

if __name__ == "__main__":
    window.mainloop()
