#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Генератор коммерческих предложений с ИИ (GigaChat)
Версия: 2.0 — интерфейс на CustomTkinter

Особенности:
  * Генерация через облачный GigaChat (src/llm_provider.py)
  * Предпросмотр строится в отдельном потоке — окно не «зависает»
  * Если GigaChat недоступен/вернул None — автоматический fallback на шаблоны
  * Маршрут, оплата и НДС подставляются из настроек программы, а не моделью
  * Выбор городов запоминается ОТДЕЛЬНО для каждого региона (cities_by_region):
    раньше при переключении вкладки выбор предыдущего региона затирался

Безопасность: ключ читается только из .env и НИКОГДА не пишется в лог.

Интерфейс: только CustomTkinter (тёмная тема). Чистый tkinter оставлен там,
где он и должен быть: messagebox, filedialog и переменные (IntVar/StringVar) —
они работают поверх CustomTkinter без конфликтов.
"""

import json
import os
import random
import re
import sys
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import customtkinter as ctk
import pyperclip
import tkinter as tk
from tkinter import messagebox

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Импортируем общий модуль: пути проекта и логгер
from logger_config import setup_logger, SETTINGS_PATH, PROJECT_ROOT, LOG_FILE as LOG_FILE_HINT

# Единый провайдер LLM (GigaChat)
from llm_provider import (
    DEFAULT_GIGACHAT_TIMEOUT,
    DEFAULT_PAYMENT_KEY,
    DEFAULT_SYSTEM_PROMPT,
    PAYMENT_LABELS,
    PAYMENT_TYPES,
    detect_ssl_interception,
    generate_gigachat,
    get_ca_bundle_status,
    get_env_write_path,
    get_ssl_help,
    get_status_info,
    has_ssl_error,
    is_gigachat_configured,
    payment_key_from_label,
    payment_line,
    sanitize_key,
    save_gigachat_key,
    start_interception_scan,
    warmup_gigachat,
)

# ============================================
# ОФОРМЛЕНИЕ (CustomTkinter)
# ============================================
# Тема и цветовая схема задаются ДО создания окна: CustomTkinter читает их
# при создании первого виджета.

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

# Палитра окна. Одна гамма на весь интерфейс: иначе кнопки, карточки и
# статус-бар «разъезжаются» по оттенкам и выглядят собранными из разных программ.
BG_WINDOW = "#12151a"      # фон окна
BG_CARD = "#1b1f27"        # карточки (колонки настроек и предпросмотра)
BG_HEADER = "#171b22"      # шапка
BG_STATUS = "#0e1116"      # статус-бар
TEXT_MAIN = "#e9ecf1"      # основной текст
TEXT_DIM = "#98a2b3"       # пояснения и подписи

COLOR_PRIMARY = "#2e7d32"       # главное действие (применить, копировать)
COLOR_PRIMARY_HOVER = "#256428"
COLOR_ACCENT = "#1565c0"        # Word, служебные действия
COLOR_ACCENT_HOVER = "#11519c"
COLOR_VIOLET = "#7b1fa2"        # «показать новый вариант»
COLOR_VIOLET_HOVER = "#641a84"
COLOR_WARN = "#ef6c00"          # варианты, предупреждения
COLOR_WARN_HOVER = "#c85b00"
COLOR_DANGER = "#c62828"        # выход, SSL
COLOR_DANGER_HOVER = "#a42020"
COLOR_TOOL = "#37474f"          # служебные кнопки шапки
COLOR_TOOL_HOVER = "#2c3a41"
COLOR_AI_ON = "#2e7d32"
COLOR_AI_OFF = "#c62828"

# text_color для надписи о состоянии GigaChat
COLOR_TEXT_OK = "#66bb6a"
COLOR_TEXT_ERR = "#ef5350"
COLOR_TEXT_WARN = "#ffb74d"

# Ширина левой колонки с настройками: фиксированная, чтобы предпросмотр
# «не прыгал» при переключении вкладок.
LEFT_COLUMN_WIDTH = 380

# Настраиваем логгер
logger = setup_logger(__name__)

# ============================================
# ГЛОБАЛЬНЫЙ ПЕРЕХВАТ ИСКЛЮЧЕНИЙ
# ============================================
# Раньше необработанное исключение в фоновом потоке убивало программу молча:
# окно исчезало, в логе оставалась одна строка «Программа завершена».
# Теперь любая такая ошибка попадает в лог с полным трейсбеком.


def _excepthook(exc_type, exc_value, exc_tb) -> None:
    """Пишет необработанные исключения в лог и в stderr."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return

    try:
        logger.critical(
            "НЕОБРАБОТАННОЕ ИСКЛЮЧЕНИЕ", exc_info=(exc_type, exc_value, exc_tb)
        )
    except Exception:
        pass

    print("НЕОБРАБОТАННОЕ ИСКЛЮЧЕНИЕ:", file=sys.stderr)
    traceback.print_exception(exc_type, exc_value, exc_tb)


def _thread_excepthook(args) -> None:
    """То же самое для фоновых потоков (threading.excepthook)."""
    _excepthook(args.exc_type, args.exc_value, args.exc_traceback)


sys.excepthook = _excepthook
try:
    threading.excepthook = _thread_excepthook
except Exception:  # pragma: no cover
    pass

logger.info("=" * 70)
logger.info("ЗАПУСК ГЕНЕРАТОРА КП С GIGACHAT (интерфейс CustomTkinter)")
logger.info("=" * 70)
logger.info(
    "Режим запуска: %s | Python %s",
    "собранный EXE" if getattr(sys, "frozen", False) else "исходники",
    sys.version.split()[0],
)
logger.info("Корень проекта: %s", PROJECT_ROOT)

# ============================================
# КОНФИГУРАЦИЯ ГЕНЕРАЦИИ (параметры под скорость)
# ============================================

AI_TEMPERATURE = 0.85      # креативность
AI_MAX_TOKENS = 150        # короче ответ — быстрее приходит (было 280)
# Таймаут НЕ задаём здесь: единственный источник истины — llm_provider,
# который читает GIGACHAT_TIMEOUT из .env (иначе берёт 60 с). Значение ниже
# только для справки в логе/интерфейсе, чтобы не было двух правд.
AI_TIMEOUT = int(DEFAULT_GIGACHAT_TIMEOUT)

# Короткий системный промпт: меньше токенов на вход — быстрее ответ.
# Специфика проекта: машины в лоте — это ГРУЗ для автовоза, а не наш транспорт,
# поэтому роль и запреты продублированы здесь явно (провайдер общий с другими
# задачами и его DEFAULT_SYSTEM_PROMPT может измениться).
GIGACHAT_SYSTEM_PROMPT = (
    DEFAULT_SYSTEM_PROMPT
    + " Ты пишешь объявление для автовозов: «лот машин» — это груз, который ждёт"
    " автовоз. Отвечай 4–5 строками: заголовок, маршрут, одна фраза про ожидание"
    " автовоза, оплата и короткий призыв писать в личку."
)

# ============================================
# ТИП ОПЛАТЫ
# ============================================
# Сами константы (PAYMENT_TYPES, PAYMENT_ORDER, DEFAULT_PAYMENT_KEY, payment_line)
# живут в llm_provider — единый источник истины для интерфейса и текста.

# Значения, которые считаем «ключ не задан» (для отчёта диагностики).
# Держим копию списка из llm_provider: GUI не должен зависеть от приватных
# констант провайдера, а расхождение здесь не критично — это только текст.
_PLACEHOLDER_KEY_VALUES = {
    "",
    "your_key_here",
    "your_key",
    "changeme",
    "none",
    "null",
    "вставьте_ключ",
    "тут_ключ",
}

# Виджеты окна диагностики. Хранится ссылка на текстовое поле, чтобы
# фоновый поток мог вернуть в него готовый отчёт через window.after().
_diag_widgets: Dict[str, object] = {}

# ============================================
# НАСТРОЙКИ
# ============================================

SETTINGS_FILE = SETTINGS_PATH

# Регионы (они же — вкладки выбора городов). Порядок важен: он же порядок вкладок.
REGIONS = ("Юга", "Владивосток", "Восток")
DEFAULT_REGION = "Юга"

# Статус GigaChat определяем один раз при старте (ключ читается из .env)
GIGACHAT_READY = is_gigachat_configured()
STATUS_INFO = get_status_info()
GIGACHAT_MODEL = STATUS_INFO.get("gigachat_model", "GigaChat")

logger.info(
    "GigaChat: настроен=%s | модель=%s | .env=%s (существует=%s)",
    GIGACHAT_READY,
    GIGACHAT_MODEL,
    STATUS_INFO.get("env_path"),
    STATUS_INFO.get("env_exists"),
)


def load_settings():
    """Загружает настройки из файла (сырой словарь или None)."""
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _normalize_region(value) -> str:
    """
    Приводит название региона из настроек к одному из REGIONS.

    В старых/битых settings.json там могло оказаться что угодно — программа
    не должна из-за этого падать или показывать пустую вкладку.
    """
    text = str(value or "").strip()
    for region in REGIONS:
        if text.lower() == region.lower():
            return region
    return DEFAULT_REGION


def _read_cities_by_region(saved: Optional[dict]) -> Dict[str, List[str]]:
    """
    Достаёт выбор городов по регионам, ПЕРЕНОСЯ старый формат настроек.

    Новый формат: {"cities_by_region": {"Юга": [...], "Владивосток": [...], ...}}
    Старый формат: {"cities": [...], "region": "Юга"} — один список на всё.

    Миграция обязательна: у пользователей уже лежит старый settings.json,
    и при обновлении программы их выбор не должен пропасть.
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


def save_settings(
    selected_cities,
    region,
    use_ai=True,
    creativity=0.8,
    payment_key: str = DEFAULT_PAYMENT_KEY,
    cities_map: Optional[Dict[str, List[str]]] = None,
):
    """
    Сохраняет настройки в файл. Ключи GigaChat здесь НЕ хранятся.

    Выбор городов пишется в двух видах:
      * cities_by_region — основной формат: свой список для каждого региона;
      * cities           — выбор ТЕКУЩЕГО региона. Оставлен для совместимости
        с уже разосланными сборками: старый EXE читает только это поле и без
        него потерял бы выбор городов.

    Args:
        cities_map: снимок выбора по регионам. None — берём текущее состояние
                    из памяти (cities_by_region).
    """
    try:
        if cities_map is None:
            cities_by_region[region] = list(selected_cities)
            snapshot = dict(cities_by_region)
        else:
            snapshot = dict(cities_map)
            snapshot[region] = list(selected_cities)

        payload = {
            "cities_by_region": snapshot,
            "cities": list(selected_cities),
            "region": region,
            "use_ai": use_ai,
            "model": GIGACHAT_MODEL,
            "creativity": creativity,
            "payment": payment_key,
        }

        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        logger.info(
            "Настройки сохранены: регион=%s, городов=%d, оплата=%s",
            region,
            len(selected_cities),
            payment_key,
        )
        return True
    except Exception as e:
        logger.error("Ошибка сохранения настроек: %s", e)
        return False


# Загружаем сохранённые настройки
saved = load_settings()

# Выбор городов по регионам (миграция старого формата внутри _read_cities_by_region)
cities_by_region: Dict[str, List[str]] = _read_cities_by_region(saved)

if saved:
    CURRENT_REGION = _normalize_region(saved.get("region"))
    USE_AI = bool(saved.get("use_ai", True))
    try:
        CREATIVITY = float(saved.get("creativity", 0.8))
    except (TypeError, ValueError):
        CREATIVITY = 0.8
    # Слайдер ограничен диапазоном 0.2…1.0 — держим значение в этих рамках,
    # иначе ползунок встанет не туда, где написано число.
    CREATIVITY = min(max(CREATIVITY, 0.2), 1.0)
    PAYMENT_KEY = saved.get("payment", DEFAULT_PAYMENT_KEY)
    if PAYMENT_KEY not in PAYMENT_TYPES:
        # Старый или битый settings.json — не роняем программу, берём умолчание
        PAYMENT_KEY = DEFAULT_PAYMENT_KEY
else:
    CURRENT_REGION = DEFAULT_REGION
    USE_AI = True
    CREATIVITY = 0.8
    PAYMENT_KEY = DEFAULT_PAYMENT_KEY

logger.info(
    "Загружена креативность: %s, тип оплаты: %s, регион: %s",
    CREATIVITY,
    PAYMENT_KEY,
    CURRENT_REGION,
)

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

# Города по умолчанию для региона, который пользователь ещё не открывал.
# Для Востока креативный режим сюда НЕ входит: это режим, а не город —
# раньше при переключении вкладки он отмечался вместе со всеми городами,
# и в тексте получалось «Без города» плюс список городов.
DEFAULT_CITIES_BY_REGION: Dict[str, List[str]] = {
    "Юга": list(ALL_CITIES_SOUTH),
    "Владивосток": list(ALL_CITIES_VLADIVOSTOK),
    "Восток": list(ALL_CITIES_EAST),
}

# Названия вкладок. Регион определяется по названию, а не по номеру вкладки:
# номера сбиваются, если вкладку когда-нибудь добавят или переставят.
TAB_TITLES: Dict[str, str] = {
    "Юга": "📍 Юга",
    "Владивосток": "📍 Владивосток",
    "Восток": "📍 Восток",
}
TAB_TITLES_BY_NAME: Dict[str, str] = {title: region for region, title in TAB_TITLES.items()}

# Регион, города которого показываются сейчас (для сохранения выбора при
# переключении вкладок).
_active_region = CURRENT_REGION

# Если для текущего региона выбора нет вовсе — показываем города региона
# целиком: так же вела себя прежняя версия при переключении вкладок.
if CURRENT_REGION not in cities_by_region:
    cities_by_region[CURRENT_REGION] = list(DEFAULT_CITIES_BY_REGION[CURRENT_REGION])

SELECTED_CITIES = list(cities_by_region[CURRENT_REGION])

# ============================================
# ВАРИАНТЫ СТИЛЕЙ ДЛЯ FALLBACK-ШАБЛОНОВ
# ============================================
#
# Специфика проекта: машины в лоте — это ГРУЗ для автовоза. Поэтому в каждой
# фразе рядом с упоминанием машин стоит контекст «на автовоз»: иначе читатель
# понимает текст как «наши машины/транспорт готовы ехать».
#
# Первая строка объявления зависит от того, один город в лоте или несколько:
# «одно направление» и «сборный маршрут» — это разные предложения для
# перевозчика, и раньше текст об этом умалчивал.

OPENERS_ONE = [
    "🚛 Есть лот машин на автовоз — направление ниже!",
    "🚛 Свежий лот на автовоз: одно направление, один рейс!",
    "🚛 Лот машин ждёт автовоз — выгрузка в одном городе!",
    "🚛 Автовозам: свежий лот на одно направление!",
    "🚛 Лот сформирован, ищем автовоз на один город!",
    "🚛 Машины на отправку — нужен автовоз на точку ниже!",
    "🚛 Груз для автовоза: одно направление, без перевалок!",
    "🚛 Новый лот машин под автовоз — один город на выгрузку!",
    "🚛 Лот машин на автовоз — забирайте одним рейсом!",
    "🚛 Есть машины на перевозку автовозом — маршрут ниже!",
    "🚛 Лот машин на одно направление — нужен автовоз!",
    "🚛 Свежий лот: один город на выгрузку, ищем автовоз!",
]

OPENERS_MULTI = [
    "🚛 Есть лот машин на автовоз — сборный маршрут!",
    "🚛 Свежий лот на автовоз: несколько городов в одном рейсе!",
    "🚛 Лот машин ждёт автовоз — маршрут из нескольких точек!",
    "🚛 Автовозам: сборный лот по нескольким городам!",
    "🚛 Лот сформирован, ищем автовоз на сборный маршрут!",
    "🚛 Машины на отправку — нужен автовоз по маршруту ниже!",
    "🚛 Груз для автовоза: сборный маршрут, несколько городов!",
    "🚛 Новый лот машин под автовоз — забрать из нескольких точек!",
    "🚛 Лот машин на автовоз — маршрут с несколькими выгрузками!",
    "🚛 Сборный лот машин — нужен автовоз на весь маршрут!",
    "🚛 Лот машин: несколько городов, один автовоз!",
    "🚛 Есть машины на перевозку автовозом — маршрут ниже!",
]

# Общий список — на случай, если понадобится выбор без учёта количества городов
OPENERS = OPENERS_ONE + OPENERS_MULTI

STATUSES = [
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
    "🚗 Ждём автовоз — машины на площадке, погрузка без ожидания.",
    "🚗 Машины к отправке готовы, нужен автовоз на рейс.",
    "🚗 Лот машин собран, автовоз требуется в ближайшие дни.",
    "🚗 Машины ждут автовоз, погрузка займёт минимум времени.",
    "🚗 Есть машины на автовоз — погрузка по готовности.",
    "🚗 Машины на площадке и готовы к погрузке, ищем автовоз.",
    "🚗 Лот машин ждёт автовоз — загрузка быстрая, без простоев.",
    "🚗 Машины стоят под погрузку, автовоз нужен срочно.",
    "🚗 Автовоз требуется: машины готовы, погрузка день в день.",
    "🚗 Машины ждут автовоз — рейс можно планировать сейчас.",
    "🚗 Лот машин на площадке, ждём автовоз под загрузку.",
    "🚗 Машины готовы к отправке автовозом — нужен перевозчик.",
]

LEGALS = [
    "📋 Работаем по договору, оплата без задержек.",
    "📄 Заключаем договор, выплаты строго в срок.",
    "🤝 Официальное оформление, гарантируем надёжность.",
    "📑 Работаем официально, с закрывающими документами.",
    "📋 Все сделки по договору, выплаты чётко в срок.",
    "📄 Юридически чистая сделка — договор и документы.",
    "⚖️ Работаем по договору, все документы предоставляем.",
    "🧾 Оплата с полным пакетом закрывающих документов.",
    "🤝 Договор оформляем сразу, оплата в срок.",
    "📑 Официальный договор, оплата по факту выгрузки.",
    "📋 Прозрачные условия: договор, документы, оплата в срок.",
    "⚖️ Юридическое сопровождение сделки — работаем официально.",
    "📄 Заявка, договор, документы — всё официально.",
    "🧾 Закрывающие документы предоставляем в срок.",
    "📑 Оплата в срок по договору, без задержек и сюрпризов.",
    "🤝 Договор и полный пакет документов — работаем чисто.",
    "📋 Оформление официальное, реквизиты и договор сразу.",
    "⚖️ Работаем прозрачно: договор, счёт, закрывающие.",
    "📄 Все условия в договоре, оплата без просрочек.",
    "📑 Закрывающие документы в срок, работаем официально.",
]

CTA = [
    "✍️ Все детали и ставки — в личку. Пишите!",
    "✍️ Ставки и подробности — в личные сообщения.",
    "✍️ Детали и ставки — в личку, ждём сообщений!",
    "✍️ Пишите в личку — обсудим ставки и даты.",
    "✍️ Подробности и ставки — только в личку.",
    "✍️ Всё по лоту — в личку. Пишите!",
    "✍️ Ставки, даты и детали — в личку, отвечаем быстро.",
    "✍️ Напишите в личку — рассчитаем ставку под ваш рейс.",
    "✍️ Пишите в личку: обсудим ставку и дату загрузки.",
    "✍️ Есть свободный автовоз? Пишите в личку — дадим детали.",
    "✍️ Все условия — в личке, отвечаем оперативно.",
    "✍️ Ставка и детали по лоту — в личных сообщениях.",
    "✍️ Пишите — обсудим рейс, ставку и сроки.",
    "✍️ Детали в личку: города, даты, ставка.",
    "✍️ Есть вопросы по лоту? Пишите в личку.",
    "✍️ Связь в личке — ставка и даты там же.",
    "✍️ Пишите в личку, если готовы взять рейс.",
    "✍️ Обсудим ставку в личных сообщениях — пишите.",
    "✍️ Все подробности по лоту — в личке, ждём!",
    "✍️ Пишите в личку: договоримся по ставке и датам.",
]

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

    История хранится в памяти процесса и на диск не пишется: после перезапуска
    повтор фразы безобиден, а лишний файл состояния — лишняя точка отказа.
    """

    def __init__(self, depth: int = 3) -> None:
        self._depth = max(1, int(depth))
        self._history: Dict[str, List[str]] = {}

    def pick(self, key: str, variants: Sequence[str]) -> str:
        """
        Возвращает вариант, которого не было среди последних выданных.

        Args:
            key:      имя списка («statuses», «cta»…) — история у каждого своя;
            variants: список фраз.

        Returns:
            Выбранную фразу. Пустой список — пустую строку (не исключение:
            вызывающий код подставляет шаблон и не должен падать).
        """
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

        # Держим историю короче самого списка, иначе выбор «не из недавних»
        # станет невозможным и мы будем каждый раз чистить историю.
        limit = min(self._depth, max(1, len(pool) - 1))
        del recent[:-limit]
        return choice

    def remember(self, key: str, value: str) -> None:
        """Запоминает фразу, полученную не через pick (например, от ИИ)."""
        if not value:
            return
        recent = self._history.setdefault(key, [])
        recent.append(value)
        del recent[:-self._depth]

    def clear(self) -> None:
        """Забывает всё — например, по кнопке «Кэш»."""
        self._history.clear()


# Единая история фраз для всего модуля
_PHRASES = RecentHistory()


def _pick_opener(city_count: int) -> str:
    """Открывающая строка: для одного города и для сборного маршрута — разные."""
    if city_count <= 1:
        return _PHRASES.pick("openers_one", OPENERS_ONE)
    return _PHRASES.pick("openers_multi", OPENERS_MULTI)


# ============================================
# ЗАЩИТА ОТ ГАЛЛЮЦИНАЦИЙ ИИ
# ============================================


class AIGuard:
    """Класс для защиты от галлюцинаций ИИ"""

    @staticmethod
    def validate_ai_response(text: str) -> Tuple[bool, str]:
        """Валидация ответа ИИ"""
        if not text or len(text.strip()) < 20:
            return False, "Текст слишком короткий"

        # «автовоз» обязателен: без него текст читается как предложение нашего
        # транспорта, а не лота машин, который ждёт перевозчика-автовоза.
        required_keywords = ["автовоз", "машин", "ло", "груз", "авто", "транспорт"]
        if not any(kw in text.lower() for kw in required_keywords):
            return False, "Нет ключевых слов по теме"

        if "автовоз" not in text.lower():
            return False, "В тексте нет контекста автовоза"

        wrong_phrases = ["ваши грузы", "ваш груз", "вашего груза", "погрузим ваши"]
        if any(phrase in text.lower() for phrase in wrong_phrases):
            return False, "Неправильная формулировка"

        if len(text) > 2000:
            return False, "Текст слишком длинный"

        return True, "OK"

    @staticmethod
    def fix_ai_response(text: str) -> str:
        """Исправляет типичные проблемы в ответе ИИ"""
        text = re.sub(r"\n{3,}", "\n\n", text)
        text = re.sub(r" {2,}", " ", text)

        # Убираем служебные пометки разметки, которые иногда добавляет модель
        text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
        text = re.sub(r"^#+\s*", "", text, flags=re.MULTILINE)

        # Удаляем подписи и обращения
        text = re.sub(r"С уважением,?\s*\[.*?\]", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\(название компании\)", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\[Название компании\]", "", text, flags=re.IGNORECASE)
        text = re.sub(r"С уважением,?\s*$", "", text, flags=re.IGNORECASE)

        return text.strip()


# ============================================
# УМНЫЙ КЭШ
# ============================================


class AICache:
    """Кэш с учётом времени жизни (30 секунд — чтобы текст оставался свежим)"""

    _cache: Dict[str, Tuple[str, float]] = {}

    @classmethod
    def get(cls, key: str) -> Optional[str]:
        if key in cls._cache:
            value, timestamp = cls._cache[key]
            if time.time() - timestamp < 30:
                return value
            del cls._cache[key]
        return None

    @classmethod
    def set(cls, key: str, value: str):
        cls._cache[key] = (value, time.time())

    @classmethod
    def clear(cls):
        cls._cache.clear()
        logger.info("Кэш очищен")


# ============================================
# ПРОМПТ ДЛЯ GIGACHAT (КОРОТКИЙ — РАДИ СКОРОСТИ)
# ============================================


# Заголовки по регионам для ИИ-текста. Маршрут и заголовок программа собирает
# сама: модель сочиняет только тело объявления, поэтому города и служебные
# формулировки всегда совпадают с тем, что отмечено на вкладке.
REGION_HEADERS = {
    "Юга": "🚛 Лот машин на автовоз: Москва → Юга!",
    "Владивосток": "🚛 Лот машин на автовоз: Владивосток → по России!",
    "Восток": "🚛 Лот машин на автовоз: Москва → Восток!",
}
DEFAULT_REGION_HEADER = "🚛 Лот машин на автовоз!"


def build_lot_header(region: str) -> str:
    """Заголовок объявления для региона."""
    return REGION_HEADERS.get(region, DEFAULT_REGION_HEADER)


def build_route_block(cities: List[str]) -> str:
    """
    Собирает блок маршрута для текста КП.

    Маршрут строит КОД, а не модель: города с вкладки известны точно, а ИИ
    может их перепутать, дописать лишние или потерять часть. Заодно формат
    «Погрузка: X + Y» всегда одинаковый.
    """
    route_cities = [c for c in cities if c != CREATIVE_MODE]
    if not route_cities:
        return ""

    cities_line = " + ".join(route_cities)
    origin = "Владивосток" if all("Владивосток" in c for c in route_cities) else "Москва"
    return f"📍 Погрузка: {origin}\n📍 Выгрузка: {cities_line}"


def _build_prompt(region: str, cities: List[str], pay_key: str = DEFAULT_PAYMENT_KEY) -> str:
    """
    Короткий промпт: меньше входных токенов — быстрее ответ.

    Держим его компактным осознанно: длинные инструкции заметно увеличивают
    время генерации, а качество для такого короткого текста не растёт.

    Специфика автовозов передана явно: машины в лоте — это груз, который ждёт
    автовоз. Без этого модель писала «авто готовы к отправке», и текст читался
    как предложение наших грузовиков, а не лота для перевозчика.

    Маршрут подставляется готовым блоком: модель его не сочиняет, а только
    обрамляет — так города в тексте гарантированно совпадают с вкладкой.

    В конце дан один пример удачного ответа (few-shot): на нём модель заметно
    стабильнее держит формат «три строки с эмодзи». Оплата в примере взята
    настоящая — из настроек, иначе модель копировала бы чужую формулировку.
    """
    region_map = {
        "Юга": "юг России (Ростов, Краснодар, Новороссийск)",
        "Владивосток": "из Владивостока по России",
        "Восток": "Урал, Сибирь, Дальний Восток",
    }
    region_desc = region_map.get(region, region)
    route_block = build_route_block(cities)
    pay = payment_line(pay_key)

    if route_block:
        lot_block = (
            f"Лот машин (груз) для автовоза, точка А → точка Б: Москва → {region_desc}.\n"
            f"Города уже указаны в блоке маршрута ниже, придумывать свои нельзя.\n"
            f"Оплата: {pay}\n"
        )
    else:
        # Креативный режим «Без города»: конкретных городов нет — маршрут
        # описываем направлением, выдумывать города модели запрещено.
        lot_block = (
            f"Лот машин (груз) для автовоза: Москва → {region_desc}.\n"
            f"Оплата: {pay}\n"
        )

    # Пример удачного ответа. Копировать дословно не нужно — важно, что строк
    # ровно три, у каждой свой эмодзи и нет ни городов, ни цены.
    example_block = (
        "Пример удачного ответа (держи формат и тон, дословно не копируй):\n"
        "🚗 Машины на площадке, автовоз нужен сегодня — погрузка быстрая.\n"
        f"💰 {pay}\n"
        "✍️ Ставка и даты загрузки — в личку, отвечаем сразу.\n"
    )

    return (
        f"Составь строки для объявления перевозчикам-автовозам.\n"
        f"{lot_block}"
        f"{example_block}"
        f"Верни РОВНО три строки, каждая со своим эмодзи:\n"
        f"🚗 — короткое живое предложение (до 12 слов) о том, что машины ждут автовоз; "
        f"города в этой строке не называй;\n"
        f"💰 — текст оплаты из условия выше, слово в слово;\n"
        f"✍️ — короткий призыв писать в личку.\n"
        f"Без заголовка, без маршрута, без markdown, без пояснений: "
        f"маршрут и заголовок добавит программа. "
        f"НЕ упоминай: города, цену, количество машин, названия и марки машин."
    )


def generate_with_gigachat(
    region: str, cities: List[str], pay_key: str = DEFAULT_PAYMENT_KEY
) -> Optional[str]:
    """
    Генерация текста через GigaChat.

    Returns:
        Валидный текст или None (тогда вызывающий код использует шаблоны).
    """
    if not GIGACHAT_READY:
        return None

    prompt = _build_prompt(region, cities, pay_key)

    start = time.time()
    text = generate_gigachat(
        prompt=prompt,
        system_prompt=GIGACHAT_SYSTEM_PROMPT,
        temperature=AI_TEMPERATURE,
        max_tokens=AI_MAX_TOKENS,
    )
    elapsed = time.time() - start

    if not text:
        logger.warning("GigaChat не вернул текст (%.2f сек) — будет fallback", elapsed)
        return None

    is_valid, reason = AIGuard.validate_ai_response(text)
    if not is_valid:
        logger.warning("Ответ GigaChat не прошёл валидацию: %s", reason)
        return None

    logger.info("✅ GigaChat сгенерировал текст за %.2f сек (%d символов)", elapsed, len(text))
    return AIGuard.fix_ai_response(text)


# ============================================
# СБОРКА ТЕЛА ОБЪЯВЛЕНИЯ (данные подставляет код, а не ИИ)
# ============================================

# Строки, которые модель иногда добавляет сама, хотя их место занимает код.
# Если такую строку оставить, в тексте появится второй маршрут или «оплата»
# без НДС — то есть именно то, от чего мы уходим.
_ROUTE_LINE_MARKERS = (
    "погрузка",
    "выгрузка",
    "маршрут",
    "точка а",
    "точка б",
    "загрузка:",
    "→",
    "->",
    "—>",
)
_PAY_LINE_MARKERS = (
    "оплата",
    "оплату",
    "оплате",
    "безнал",
    "наличн",
    "ндс",
    "предоплат",
)


def _mentions_route_city(line: str, cities: Optional[List[str]]) -> bool:
    """
    Есть ли в строке хотя бы один город из маршрута.

    Сравниваем по первому слову названия: «Владивосток → Москва» и
    «Владивосток» должны считаться одним городом.
    """
    if not cities:
        return False

    low = line.lower()
    for city in cities:
        for part in city.split("→"):
            word = part.strip().lower()
            if not word:
                continue
            # «Ростов-на-Дону» ищем целиком, «Владивосток → Москва» — по частям
            if word in low:
                return True
            head = word.split()[0] if word.split() else ""
            if len(head) > 4 and head in low:
                return True
    return False


def _iter_body_lines(text: str) -> List[str]:
    """
    Разбивает ответ модели на строки, попутно снимая markdown и декор.

    Модель любит дописывать украшения вроде «🌿» или «ждет» без ё — для
    объявления это мусор, поэтому чистим здесь, а не в шаблонах.
    """
    cleaned = AIGuard.fix_ai_response(text)
    lines = []
    for raw in cleaned.splitlines():
        line = raw.strip()
        if not line:
            continue
        line = re.sub(r"^[-•*—\d.)\s]+", "", line).strip()
        # Пробелы перед знаками препинания и «ждет» без ё — частая мелочь модели
        line = re.sub(r"\s+([,.!?…])", r"\1", line)
        line = line.replace("ждет", "ждёт").replace("Ждет", "Ждёт")
        if line:
            lines.append(line)
    return lines


def _strip_decorative_emoji(line: str, keep: str) -> str:
    """
    Убирает из строки эмодзи, кроме «своего» (например, 🚗 для статуса).

    ИИ добавляет 🌿, ✨, 🚚 и прочее — в коротком объявлении это выглядит
    случайным набором картинок. Кириллица и пунктуация не затрагиваются.
    """
    allowed = set(keep)
    return "".join(
        ch
        for ch in line
        if ch in allowed
        or not (0x1F000 <= ord(ch) <= 0x1FAFF or 0x2600 <= ord(ch) <= 0x27BF)
    ).strip()


def build_body_from_ai(
    text: str, pay_key: str, cities: Optional[List[str]] = None
) -> Tuple[str, str, str]:
    """
    Превращает ответ модели в три строки тела: 🚗 статус, 💰 оплата, ✍️ призыв.

    Маршрут и оплата — данные, а не креатив. Модель их периодически путает:
    вписывает свои города или теряет строку про оплату вместе с НДС. Поэтому
    берём от неё только фразу-статус, а оплату и призыв ставим сами —
    из исходных данных, слово в слово.

    Запасные фразы берутся через _PHRASES: подряд одна и та же строка не
    повторится, иначе два сгенерированных текста выглядели бы одинаково.
    """
    status_line = ""
    cta_line = ""

    for line in _iter_body_lines(text):
        low = line.lower()

        # Строку с маршрутом или оплатой моделью игнорируем: маршрут добавит
        # код, оплата тоже подставляется из настроек.
        if any(marker in low for marker in _ROUTE_LINE_MARKERS):
            continue
        if any(marker in low for marker in _PAY_LINE_MARKERS):
            continue

        if not status_line and ("🚗" in line or "машин" in low or "авто" in low):
            status_line = line
            continue
        if not cta_line and ("✍" in line or "пиш" in low or "лич" in low):
            cta_line = line

    # Статус — единственное место с творчеством. Если модель не дала внятной
    # фразы, берём шаблонную: текст всё равно останется про автовоз.
    if status_line:
        status_line = _strip_decorative_emoji(status_line, "🚗")
    if len(status_line) < 15 or "автовоз" not in status_line.lower():
        status_line = _PHRASES.pick("statuses", STATUSES)
    elif _mentions_route_city(status_line, cities):
        # «…автовоз! Москва—Владивосток, Россия.» — города уже есть в блоке
        # маршрута, в статусе они только создают путаницу.
        status_line = _PHRASES.pick("statuses", STATUSES)

    if cta_line:
        cta_line = _strip_decorative_emoji(cta_line, "✍️")
    if len(cta_line) < 5:
        cta_line = _PHRASES.pick("cta", CTA)

    return status_line, payment_line(pay_key), cta_line


# ============================================
# ФУНКЦИИ ГЕНЕРАЦИИ
# ============================================


def get_selected_cities():
    """
    Отмеченные города.

    Отметки есть только у текущего региона: при переключении вкладки галочки
    остальных регионов снимаются (их выбор лежит в cities_by_region).
    """
    try:
        return [city for city, var in city_vars.items() if var.get() == 1]
    except NameError:
        # Виджеты ещё не созданы (ранний вызов) — это нормально
        return []


def get_current_region():
    """
    Определяет текущий активный регион по вкладке.

    Регион ищем по НАЗВАНИЮ вкладки, а не по её номеру: номер сбивается,
    стоит поменять порядок вкладок или добавить новую.
    """
    try:
        tab_name = notebook.get()
    except Exception:
        # Виджета ещё нет или окно уже закрыто — отдаём последний известный регион
        return CURRENT_REGION
    return TAB_TITLES_BY_NAME.get(tab_name, CURRENT_REGION)


def get_payment_key() -> str:
    """
    Возвращает выбранный тип оплаты.

    Значение читается из выпадающего списка, а PAYMENT_KEY — запасной вариант:
    до создания виджета (или если Tk уже не отвечает) функция не должна падать.
    """
    try:
        label = payment_var.get()
    except NameError:
        return PAYMENT_KEY
    return payment_key_from_label(label)


def update_payment(_event=None):
    """Обработчик выбора оплаты: сохраняем настройку и обновляем предпросмотр."""
    global PAYMENT_KEY
    PAYMENT_KEY = get_payment_key()
    logger.info("Выбран тип оплаты: %s", PAYMENT_KEY)
    AICache.clear()  # старый текст собран с прежней формулировкой оплаты
    save_settings(
        get_selected_cities(), get_current_region(), USE_AI, CREATIVITY, PAYMENT_KEY
    )
    refresh_preview_async()


def format_cities_list(cities: List[str]) -> str:
    """Форматирует список городов для вставки в текст"""
    if not cities:
        return ""

    cities = [c for c in cities if c != CREATIVE_MODE]
    if not cities:
        return ""

    if len(cities) == 1:
        return cities[0]

    styles = [
        lambda: ", ".join(cities),
        lambda: ", ".join(cities[:-1]) + " и " + cities[-1],
        lambda: "\n".join([f"• {city}" for city in cities]),
        lambda: "\n".join([f"{i+1}. {city}" for i, city in enumerate(cities)]),
    ]
    return random.choice(styles)()


def _guess_city_count(cities_text: str) -> int:
    """
    Прикидывает количество городов по готовой строке маршрута.

    Нужна только для обратной совместимости: если build_template_text вызвали
    без явного количества городов, выбираем открывающую фразу «на глазок».
    """
    if not cities_text:
        return 0
    if "\n" in cities_text:
        return 2
    if " и " in cities_text or ", " in cities_text:
        return 2
    return 1


def build_template_text(
    cities_text: str,
    pay_key: str = DEFAULT_PAYMENT_KEY,
    city_count: Optional[int] = None,
) -> str:
    """
    Fallback: собирает текст из случайных шаблонных фраз.

    Структура повторяет то, что мы требуем от ИИ: заголовок с автовозом,
    маршрут, фраза про ожидание автовоза, оплата и призыв в личку.

    Args:
        cities_text: уже отформатированный список городов (может быть пустым);
        pay_key:     тип оплаты;
        city_count:  сколько городов в лоте. Если None — оцениваем по тексту.
    """
    count = _guess_city_count(cities_text) if city_count is None else int(city_count)
    opener = _pick_opener(count)
    status = _PHRASES.pick("statuses", STATUSES)
    legal = _PHRASES.pick("legals", LEGALS)
    cta = _PHRASES.pick("cta", CTA)
    pay = payment_line(pay_key)

    if cities_text:
        route_block = f"📍 Погрузка: Москва\n📍 Выгрузка: {cities_text}"
    else:
        route_block = "📍 Маршрут: Москва → по направлению"

    blocks = [opener, route_block, status, f"💰 {pay} {legal}", cta]
    return "\n\n".join(blocks)


def generate_text_sync(
    region: str,
    cities: List[str],
    use_cache: bool = True,
    pay_key: str = DEFAULT_PAYMENT_KEY,
) -> str:
    """
    Синхронно генерирует текст: сначала GigaChat, при неудаче — шаблоны.

    ВНИМАНИЕ: блокирующая функция (может занять 2–4 сек).
    Из GUI вызывать только через generate_text_async().
    """
    route_cities = [c for c in cities if c != CREATIVE_MODE]
    cities_text = format_cities_list(route_cities) if route_cities else ""

    if USE_AI and GIGACHAT_READY:
        city_str = ",".join(sorted(route_cities))
        # Уникальный ключ: кэш не должен «залипать» на одном варианте.
        # Тип оплаты входит в ключ, иначе после смены оплаты вернулся бы
        # старый текст с прежней формулировкой.
        cache_key = (
            f"gigachat_{region}_{city_str}_{pay_key}_{int(CREATIVITY * 10)}_"
            f"{random.randint(1, 1000)}"
        )

        if use_cache:
            cached = AICache.get(cache_key)
            if cached:
                logger.info("✅ Использован кэшированный текст GigaChat")
                return cached

        logger.info("Генерация GigaChat: регион=%s, городов=%d", region, len(route_cities))
        ai_text = generate_with_gigachat(region, route_cities, pay_key)

        if ai_text:
            # Заголовок, маршрут и оплату подставляет код, а не модель: у ИИ
            # берём только фразу-статус. Иначе в тексте появлялись чужие города
            # или пропадала строка оплаты вместе с упоминанием НДС.
            status_line, pay_line_text, cta_line = build_body_from_ai(
                ai_text, pay_key, route_cities
            )
            route_block = build_route_block(route_cities)
            if route_block:
                result = (
                    f"{build_lot_header(region)}\n\n"
                    f"{route_block}\n\n"
                    f"{status_line}\n"
                    f"💰 {pay_line_text}\n"
                    f"{cta_line}"
                )
            else:
                # Креативный режим: конкретных городов нет, маршрут не выводим
                result = (
                    f"{build_lot_header(region)}\n\n"
                    f"{status_line}\n"
                    f"💰 {pay_line_text}\n"
                    f"{cta_line}"
                )

            if use_cache:
                AICache.set(cache_key, result)
            # Парная запись к «🔄 Использован шаблонный текст» ниже.
            # Без неё по логу нельзя было понять, что текст пришёл от ИИ:
            # успешный путь не логировался вообще, и казалось, что программа
            # всегда работает на шаблонах.
            logger.info("✅ Использован ИИ-текст (тело от GigaChat, данные — из настроек)")
            return result

    elif USE_AI and not GIGACHAT_READY:
        logger.warning("ИИ включён, но GigaChat не настроен — использую шаблоны")

    logger.info("🔄 Использован шаблонный текст (fallback)")
    return build_template_text(cities_text, pay_key, len(route_cities))


# ============================================
# АСИНХРОННАЯ ГЕНЕРАЦИЯ ДЛЯ GUI (threading)
# ============================================

_generation_lock = threading.Lock()
_generation_in_progress = False


def _set_progress_visible(visible: bool) -> None:
    """
    Показывает/прячет анимированный прогресс-бар.

    Полоса «ползёт» (режим indeterminate) всё время, пока идёт фоновая
    генерация: по ней видно, что программа работает, а не зависла.
    """
    try:
        if visible:
            if not progress_bar.winfo_manager():
                progress_bar.grid(row=1, column=0, sticky="ew", padx=16, pady=(2, 0))
            progress_bar.start()
        else:
            progress_bar.stop()
            if progress_bar.winfo_manager():
                progress_bar.grid_remove()
    except NameError:
        # Прогресс-бар ещё не создан (ранний вызов при старте) — это нормально
        pass
    except Exception as exc:  # pragma: no cover — прогресс не критичен
        logger.debug("Не удалось переключить прогресс-бар: %s", exc)


def _set_busy(busy: bool, message: str = "") -> None:
    """
    Переключает состояние «идёт генерация»: блокирует кнопки и меняет статус.

    ВАЖНО: эту функцию нельзя вызывать, удерживая _generation_lock —
    она сама его захватывает, а Lock не реентрантный (получится вечный deadlock).
    """
    global _generation_in_progress

    with _generation_lock:
        _generation_in_progress = busy

    state = "disabled" if busy else "normal"
    try:
        refresh_button.configure(state=state)
        copy_button.configure(state=state)
        word_button.configure(state=state)
        variants_button.configure(state=state)
    except NameError:
        # Кнопки ещё не созданы (ранний вызов при старте) — это нормально
        pass

    _set_progress_visible(busy)

    if message:
        try:
            gen_status_label.configure(text=message)
        except NameError:
            pass


def generate_text_async(on_done, use_cache: bool = True) -> bool:
    """
    Запускает генерацию в отдельном потоке, чтобы окно не «зависало».

    on_done(text, error) вызывается в главном потоке Tk через window.after(0, ...),
    потому что менять виджеты из фонового потока в Tkinter нельзя.

    Returns:
        True  — генерация поставлена в очередь;
        False — уже идёт другая генерация, вызов проигнорирован.
    """
    global _generation_in_progress

    # Слот резервируется атомарно, ещё до старта потока: иначе два быстрых
    # нажатия успевают проскочить проверку и запустить два запроса.
    with _generation_lock:
        if _generation_in_progress:
            logger.info("Генерация уже идёт — повторный запрос пропущен")
            return False
        _generation_in_progress = True

    region = get_current_region()
    cities = get_selected_cities()
    # Тип оплаты снимаем в главном потоке: читать виджеты из рабочего потока нельзя
    pay_key = get_payment_key()

    _set_busy(True, "⏳ Генерация... (GigaChat)")

    def worker():
        text = ""
        error = None

        # Лок держим ТОЛЬКО вокруг генерации: клиент GigaChat не рассчитан на
        # параллельные запросы, но планировать callback под локом нельзя —
        # finish() вызовет _set_busy(), которому нужен тот же лок (deadlock).
        with _generation_lock:
            try:
                text = generate_text_sync(region, cities, use_cache=use_cache, pay_key=pay_key)
            except Exception as e:
                logger.exception("Непредвиденная ошибка генерации: %s", e)
                error = str(e)
                try:
                    route_cities = [c for c in cities if c != CREATIVE_MODE]
                    text = build_template_text(
                        format_cities_list(route_cities),
                        pay_key,
                        len(route_cities),
                    )
                except Exception:
                    text = ""

        def finish():
            _set_busy(False, "")
            # Плашку SSL обновляем в главном потоке: фон менять виджеты не может
            update_ssl_banner()
            on_done(text, error)

        _dispatch_to_ui(finish)

    threading.Thread(target=worker, name="gigachat-generate", daemon=True).start()
    return True


def _release_busy_flag() -> None:
    """
    Снимает признак «идёт генерация» без обращения к виджетам.

    Аварийный выход: если главный поток Tk недоступен (окно уже закрыто или
    цикл событий не запущен), window.after() из фонового потока падает, и без
    этого флага-предохранителя он остался бы выставленным навсегда — кнопки
    стали бы «мёртвыми» до перезапуска программы.
    """
    global _generation_in_progress
    try:
        with _generation_lock:
            _generation_in_progress = False
    except Exception:
        pass


def _dispatch_to_ui(func) -> None:
    """
    Выполняет func в главном потоке Tk.

    Вызывать виджеты из рабочего потока нельзя, поэтому используем
    window.after(0, ...). Если окно уже закрыто или главный цикл Tk
    недоступен — снимаем флаг занятости, чтобы кнопки не остались мёртвыми.
    """
    try:
        window.after(0, func)
    except Exception as e:
        logger.warning("Не удалось передать результат в главный поток Tk: %s", e)
        _release_busy_flag()


def refresh_preview_async() -> None:
    """Обновляет предпросмотр асинхронно — окно остаётся отзывчивым."""

    def on_done(text, error):
        if error:
            logger.error("Ошибка генерации предпросмотра: %s", error)
        try:
            text_preview.configure(state="normal")
            text_preview.delete("1.0", "end")
            text_preview.insert("1.0", text or "")
            text_preview.configure(state="disabled")
        except Exception as e:
            logger.error("Ошибка обновления предпросмотра: %s", e)

    generate_text_async(on_done)


# ============================================
# ДЕЙСТВИЯ КНОПОК
# ============================================


def copy_to_clipboard():
    def on_done(text, error):
        if error:
            messagebox.showerror("Ошибка", f"❌ {error}")
            return
        try:
            pyperclip.copy(text)
            messagebox.showinfo("Готово!", "✅ Текст скопирован в буфер обмена!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"❌ {e}")

    generate_text_async(on_done)


def _documents_dir() -> str:
    """
    Возвращает папку для сохранения Word-файлов.

    Обычно это рабочий стол. Но его может не быть: OneDrive перенёс папку,
    профиль ограничен, или система без рабочего стола. Раньше в этом случае
    сохранение падало с FileNotFoundError — теперь есть надёжный запасной
    вариант (папка «Документы», затем домашняя папка).
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


def _save_docx(text: str, filename: str) -> str:
    """Создаёт Word-файл с текстом КП в доступной папке пользователя."""
    doc = Document()
    title = doc.add_heading("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    date_paragraph = doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y')}")
    date_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph()

    for line in text.split("\n"):
        if line.strip():
            doc.add_paragraph(line)

    target_dir = _documents_dir()
    filepath = os.path.join(target_dir, filename)
    try:
        doc.save(filepath)
    except Exception:
        # Папка оказалась недоступна на запись — сохраняем рядом с программой
        filepath = os.path.join(PROJECT_ROOT, filename)
        logger.warning("Не удалось сохранить в %s — сохраняю в %s", target_dir, filepath)
        doc.save(filepath)

    return filepath


def save_to_word():
    def on_done(text, error):
        if error:
            messagebox.showerror("Ошибка", f"❌ {error}")
            return
        try:
            region = get_current_region()
            filename = f"КП_{region}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            path = _save_docx(text, filename)
            messagebox.showinfo("Готово!", f"✅ Word-файл создан:\n\n{path}")
        except Exception as e:
            logger.error("Ошибка сохранения Word: %s", e)
            messagebox.showerror("Ошибка", f"❌ {e}")

    generate_text_async(on_done)


def save_multiple_variants():
    """Создаёт 3 Word-файла. Генерация идёт в фоне, с прогрессом в статусе."""

    def worker():
        region = get_current_region()
        saved_files = []
        try:
            for i in range(3):
                window.after(
                    0, lambda n=i: gen_status_label.configure(text=f"⏳ Вариант {n+1}/3...")
                )
                # use_cache=False — каждый вариант должен быть уникальным
                text = generate_text_sync(region, get_selected_cities(), use_cache=False)
                filename = (
                    f"КП_{region}_вариант_{i+1}_"
                    f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                )
                saved_files.append(_save_docx(text, filename))
        except Exception as e:
            logger.error("Ошибка создания вариантов: %s", e)
            window.after(
                0,
                lambda: (
                    _set_busy(False, ""),
                    messagebox.showerror("Ошибка", f"❌ {e}"),
                ),
            )
            return

        def finish():
            _set_busy(False, "")
            # Показываем реальную папку: рабочий стол есть не на каждой машине
            folder = os.path.dirname(saved_files[0]) if saved_files else ""
            messagebox.showinfo(
                "Готово!",
                f"✅ Создано файлов: {len(saved_files)}\n\nПапка:\n{folder}",
            )

        window.after(0, finish)

    if _generation_in_progress:
        messagebox.showinfo("Подождите", "⏳ Генерация уже выполняется.")
        return

    _set_busy(True, "⏳ Создание 3 вариантов...")
    threading.Thread(target=worker, name="gigachat-variants", daemon=True).start()


def update_cities():
    selected = get_selected_cities()
    region = get_current_region()

    if selected:
        global SELECTED_CITIES
        SELECTED_CITIES = selected
        # Запоминаем выбор именно этого региона — при переходе на другой он
        # останется на месте
        cities_by_region[region] = list(selected)
        save_settings(selected, region, USE_AI, CREATIVITY, PAYMENT_KEY)

        if len(selected) == 1:
            if selected[0] == CREATIVE_MODE:
                msg = "✅ Выбран креативный режим 'Без города'"
            else:
                msg = f"✅ Выбран город:\n\n{selected[0]}"
        else:
            msg = f"✅ Выбрано: {len(selected)}\n\n{', '.join(selected)}"

        refresh_preview_async()
        messagebox.showinfo("Готово!", msg)
    else:
        messagebox.showwarning("Внимание!", "Выберите хотя бы один город!")


def apply_region_selection(region: str) -> None:
    """
    Показывает сохранённый выбор городов для региона.

    Для региона, который ещё не открывали, берём города целиком — так же
    вела себя прежняя версия при переключении вкладки. Города других регионов
    снимаются: get_selected_cities() должен возвращать только текущий регион.
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
    Смена вкладки региона.

    Главное отличие от прежней версии: выбор городов больше НЕ теряется.
    Перед переключением он сохраняется в cities_by_region[_active_region],
    а для нового региона восстанавливается из сохранённого.
    """
    global CURRENT_REGION, SELECTED_CITIES, _active_region

    # Во время построения окна вкладка выставляется программно — реагировать
    # на это не нужно: виджетов ещё нет.
    if not _gui_ready["value"]:
        return

    region = get_current_region()
    if region == _active_region:
        return

    # 1. Сохраняем выбор покидаемого региона
    cities_by_region[_active_region] = get_selected_cities()
    logger.info(
        "Регион %s: сохранён выбор городов (%d шт.)",
        _active_region,
        len(cities_by_region[_active_region]),
    )

    # 2. Показываем выбор нового региона
    _active_region = region
    CURRENT_REGION = region
    apply_region_selection(region)

    selected = get_selected_cities()
    SELECTED_CITIES = selected
    save_settings(selected, region, USE_AI, CREATIVITY, PAYMENT_KEY)

    refresh_preview_async()


def _on_city_toggle(city: str) -> None:
    """
    Реакция на установку галочки города.

    «Без города» — это режим, а не город: вместе со списком городов он
    противоречит сам себе (раньше так и получалось — в тексте появлялось
    «Без города» и рядом перечисление городов). Поэтому отметка одного
    снимает другое.
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


def _update_ai_switch_text() -> None:
    """Подпись переключателя ИИ: сразу видно, включён он или нет."""
    try:
        ai_switch.configure(text=f"🤖 ИИ: {'ВКЛ' if USE_AI else 'ВЫКЛ'}")
    except NameError:
        pass


def toggle_ai():
    global USE_AI
    USE_AI = bool(use_ai_var.get())
    logger.info("ИИ переключён: %s", "ВКЛ" if USE_AI else "ВЫКЛ")
    _update_ai_switch_text()
    save_settings(
        get_selected_cities(), get_current_region(), USE_AI, CREATIVITY, PAYMENT_KEY
    )
    refresh_preview_async()


def clear_cache():
    AICache.clear()
    # Заодно забываем историю шаблонных фраз: пользователь ждёт «новые тексты»,
    # а не повтор уже показанных удачных строк.
    _PHRASES.clear()
    messagebox.showinfo("Готово!", "✅ Кэш ИИ очищен!")
    refresh_preview_async()


def update_creativity(val):
    global CREATIVITY
    try:
        CREATIVITY = float(val)
    except (TypeError, ValueError):
        return
    try:
        creativity_label.configure(text=f"Креативность: {int(CREATIVITY * 100)}%")
    except NameError:
        pass
    save_settings(
        get_selected_cities(), get_current_region(), USE_AI, CREATIVITY, PAYMENT_KEY
    )
    AICache.clear()


def warmup_async():
    """Прогревает авторизацию GigaChat в фоне — первый запрос будет быстрее."""
    if not GIGACHAT_READY:
        messagebox.showwarning(
            "GigaChat не настроен",
            "❌ Ключ не задан.\n\n"
            "Запустите меню scripts\\Генератор КП.bat и выберите\n"
            "пункт [4] «Ввести API-ключ GigaChat».",
        )
        return

    def worker():
        ok = warmup_gigachat()
        window.after(
            0,
            lambda: messagebox.showinfo(
                "Прогрев",
                "✅ GigaChat авторизован, генерация будет быстрее!"
                if ok
                else "⚠️ Не удалось прогреть GigaChat — смотрите логи.",
            ),
        )

    threading.Thread(target=worker, name="gigachat-warmup", daemon=True).start()


def show_gigachat_help():
    """Краткая справка по ключу и SSL — прямо из интерфейса."""
    messagebox.showinfo(
        "Справка GigaChat",
        "🔑 Ключ: developers.sber.ru → Личный кабинет → Настройки → Authorization Key\n\n"
        f"Ключ хранится ТОЛЬКО в файле .env:\n{get_env_write_path()}\n"
        "и никогда не попадает в Git.\n\n"
        "🌐 Если в логах ошибка SSL (certificate verify failed) — нажмите\n"
        "кнопку «⚠️ SSL» на главном окне: там пошаговая инструкция.",
    )


def _open_log_folder() -> None:
    """Открывает папку с логами в проводнике — чтобы быстро найти файл лога."""
    try:
        log_path = str(LOG_FILE_HINT)
        folder = os.path.dirname(log_path) or str(PROJECT_ROOT)
        if not os.path.isdir(folder):
            folder = str(PROJECT_ROOT)
        if not hasattr(os, "startfile"):
            raise RuntimeError("открытие папки поддерживается только в Windows")
        os.startfile(folder)
    except Exception as exc:
        logger.warning("Не удалось открыть папку с логами: %s", exc)
        messagebox.showinfo("Логи", f"Файл лога лежит здесь:\n\n{LOG_FILE_HINT}")


# ============================================
# ДИАЛОГИ (CustomTkinter)
# ============================================


def _center_dialog(dialog: "ctk.CTkToplevel", width: int, height: int) -> None:
    """
    Ставит диалог по центру главного окна.

    Раньше диалоги открывались в углу экрана — на большом мониторе их можно
    было просто не заметить.
    """
    try:
        dialog.update_idletasks()
        x = window.winfo_rootx() + max((window.winfo_width() - width) // 2, 0)
        y = window.winfo_rooty() + max((window.winfo_height() - height) // 3, 0)
        dialog.geometry(f"{width}x{height}+{max(x, 0)}+{max(y, 0)}")
    except Exception as exc:  # pragma: no cover — центрирование не критично
        logger.debug("Не удалось отцентрировать диалог: %s", exc)


def _make_dialog(title: str, width: int, height: int) -> "ctk.CTkToplevel":
    """
    Создаёт модальное окно в стиле программы.

    CTkToplevel на Windows иногда открывается ПОД главным окном — поэтому
    после отрисовки поднимаем его явно.
    """
    dialog = ctk.CTkToplevel(window)
    dialog.title(title)
    dialog.geometry(f"{width}x{height}")
    dialog.configure(fg_color=BG_WINDOW)
    dialog.transient(window)
    _center_dialog(dialog, width, height)
    dialog.after(60, lambda: _lift_dialog(dialog))
    return dialog


def _lift_dialog(dialog: "ctk.CTkToplevel") -> None:
    """Поднимает диалог поверх главного окна (окно могло быть уже закрыто)."""
    try:
        if dialog.winfo_exists():
            dialog.lift()
            dialog.focus_force()
    except Exception:
        pass


def show_ssl_help():
    """
    Инструкция «Как исправить» для SSL-ошибки.

    Текст берётся из llm_provider.get_ssl_help(), чтобы лог и интерфейс
    не расходились. Окно с прокруткой: инструкция длинная.
    """
    dialog = _make_dialog("⚠️ SSL: как исправить", 900, 600)
    dialog.minsize(720, 480)

    ctk.CTkLabel(
        dialog,
        text="⚠️ Ошибка проверки SSL-сертификата",
        font=FONT_H2,
        text_color=COLOR_TEXT_ERR,
    ).pack(pady=(16, 6), padx=16)

    text = ctk.CTkTextbox(dialog, wrap="word", font=FONT_MONO, fg_color="#0f1216")
    text.pack(padx=16, pady=6, fill="both", expand=True)
    text.insert("1.0", get_ssl_help())
    text.configure(state="disabled")

    def open_download():
        """Открывает страницу с сертификатом в браузере по умолчанию."""
        url = "https://gu-st.ru/content/Other/doc/russian_trusted_root_ca.cer"
        try:
            import webbrowser

            webbrowser.open(url)
        except Exception as e:
            messagebox.showinfo("Ссылка", f"Откройте вручную:\n\n{url}")
            logger.warning("Не удалось открыть браузер: %s", e)

    buttons = ctk.CTkFrame(dialog, fg_color="transparent")
    buttons.pack(pady=(6, 16))

    ctk.CTkButton(
        buttons,
        text="⬇️ Открыть страницу сертификата",
        font=FONT_SMALL,
        fg_color=COLOR_ACCENT,
        hover_color=COLOR_ACCENT_HOVER,
        command=open_download,
    ).pack(side="left", padx=6)

    ctk.CTkButton(
        buttons,
        text="Закрыть",
        font=FONT_SMALL,
        fg_color=COLOR_TOOL,
        hover_color=COLOR_TOOL_HOVER,
        command=dialog.destroy,
    ).pack(side="left", padx=6)

    text.see("1.0")
    dialog.bind("<Escape>", lambda event: dialog.destroy())


def show_key_warning():
    """
    Понятное объяснение, что делать без ключа.

    Программа при этом полностью работоспособна: текст КП строится по
    шаблонам, кнопки копирования и Word работают.
    """
    messagebox.showwarning(
        "GigaChat не настроен",
        "⚠️ Ключ GigaChat не задан.\n\n"
        "Сейчас текст коммерческого предложения собирается по шаблонам —\n"
        "программа работает, но без ИИ.\n\n"
        "Чтобы включить ИИ:\n"
        "1. Нажмите кнопку 🔑 на главном окне\n"
        "2. Вставьте Authorization Key из личного кабинета developers.sber.ru\n"
        "3. Нажмите «Сохранить»\n\n"
        f"Ключ сохранится в файл:\n{get_env_write_path()}",
    )


def _diag_env_credential_state() -> str:
    """
    Определяет состояние ключа GigaChat В ФАЙЛЕ .env (не в окружении).

    Нужно, чтобы отличить три разные ситуации: ключа нет вовсе, ключ есть,
    ключ есть но выглядит затёртым/подстановочным. Возвращает строку для
    отчёта диагностики. Сам ключ не показывается — только длина.
    """
    env_path = get_env_write_path()

    try:
        content = Path(env_path).read_text(encoding="utf-8-sig")
    except Exception as exc:
        return f"⚠️ не удалось прочитать .env: {exc}"

    value = None
    for line in content.splitlines():
        line = line.strip()
        if line.startswith("GIGACHAT_CREDENTIALS="):
            value = line.split("=", 1)[1].strip()

    if value is None:
        return "❌ строки GIGACHAT_CREDENTIALS в .env нет"
    if not value:
        return "❌ GIGACHAT_CREDENTIALS пустой"
    if value.lower() in _PLACEHOLDER_KEY_VALUES:
        return f"⚠️ подстановочное значение ({value!r}) — вставьте настоящий ключ"
    return f"✅ есть ({len(value)} символов, значение скрыто)"


def _diag_certificate_section() -> List[str]:
    """
    Раздел отчёта про сертификат Минцифры: найден/не найден и что делать.

    Проверка только файловая — без сетевых запросов, поэтому окно
    диагностики открывается мгновенно.
    """
    lines: List[str] = []

    try:
        expected = get_ca_bundle_status()
    except Exception as exc:
        lines.append(f"⚠️ Не удалось проверить сертификат: {exc}")
        lines.append("")
        return lines

    if expected.get("explicit"):
        if expected.get("explicit_valid"):
            lines.append(
                f"✅ SSL-сертификат (из .env): {expected['explicit']}"
            )
        else:
            lines.append("⚠️ SSL-сертификат Минцифры: путь в .env НЕ существует")
            lines.append(f"   GIGACHAT_CA_BUNDLE_FILE={expected['explicit']}")
    elif expected.get("found"):
        lines.append(f"✅ SSL-сертификат Минцифры: найден")
        lines.append(f"   Путь: {expected['found']}")
        lines.append(
            "   📌 Если GigaChat всё равно падает с SSL-ошибкой — антивирус"
        )
        lines.append(
            "      подменяет сертификаты. Добавьте программу в исключения"
        )
        lines.append("      SSL-инспекции (Kaspersky: Настройки → Сеть → …).")
    else:
        lines.append("⚠️ SSL-сертификат Минцифры: НЕ НАЙДЕН")
        lines.append(f"   Ожидаемое место: {expected['expected']}")
        lines.append(f"   Скачать: {expected['url']}")
        lines.append(
            "   📌 Без сертификата GigaChat будет падать с SSL-ошибкой."
        )
        lines.append(
            "      Программа при этом работает: текст КП берётся из шаблонов."
        )

    # Что именно ушло в проверку TLS: сертификат Минцифры сам по себе мало
    # помогает, если антивирус подменяет цепочку — поэтому программа
    # собирает один файл из всех нужных корней и показывает его состав.
    if expected.get("problem"):
        lines.append(f"⚠️ Проблема с файлом сертификата: {expected['problem']}")

    merged = expected.get("merged_bundle")
    if merged:
        count = expected.get("merged_count") or 0
        lines.append(f"✅ Набор для проверки TLS собран ({count} корней):")
        lines.append(f"   {merged}")
    elif expected.get("found"):
        lines.append(
            "   ℹ️ Набор соберётся при первом обращении к GigaChat "
            "(см. запись в логе)."
        )

    interception = expected.get("interception") or []
    if interception:
        lines.append("")
        lines.append("🔎 Обнаружена SSL-инспекция (проверка трафика на лету):")
        for subject in interception:
            short = subject.split(",")[0].replace("CN=", "").strip()
            lines.append(f"   - {short}")
        lines.append(
            "   Корень антивируса добавлен в набор сертификатов, поэтому"
        )
        lines.append(
            "   GigaChat может работать и без отключения проверки."
        )
        lines.append(
            "   Если SSL-ошибка останется — исключите программу из"
        )
        lines.append(
            "   SSL-инспекции антивируса (Kaspersky: Настройки → Сеть →"
        )
        lines.append("   Проверка защищённых соединений → Исключения).")
    elif expected.get("store_scan") == "empty":
        lines.append("")
        lines.append(
            "ℹ️ Список корневых сертификатов Windows прочитать не удалось —"
        )
        lines.append("   проверить SSL-инспекцию автоматически нельзя.")

    lines.append("")
    lines.append("🔎 Где искали сертификат:")
    for candidate in expected.get("search_paths", []):
        lines.append(f"   - {candidate}")

    lines.append(
        f"   проверка SSL включена: {'да' if expected.get('verify_ssl') else 'НЕТ'}"
    )
    return lines


def _build_diagnostics_report() -> str:
    """
    Собирает текст отчёта «📋 Диагностика».

    Отдельная функция (а не тело обработчика кнопки) нужна, чтобы отчёт
    можно было собирать в фоновом потоке: если писать его прямо в кнопке,
    интерфейс замирает на время проверки.
    """
    import platform

    lines: List[str] = ["🔍 ДИАГНОСТИКА", "=" * 46, ""]

    # ---------- .env ----------
    env_path = get_env_write_path()
    env_size = "—"
    try:
        if Path(env_path).exists():
            env_size = f"{Path(env_path).stat().st_size} байт"
            lines.append(f"✅ .env найден: {env_path} ({env_size})")
        else:
            lines.append(f"❌ .env НЕ найден: {env_path}")
    except Exception as exc:
        lines.append(f"⚠️ .env: не удалось прочитать {env_path} ({exc})")

    lines.append(f"   ключ GigaChat: {_diag_env_credential_state()}")
    lines.append(f"   модель: {os.environ.get('GIGACHAT_MODEL', 'GigaChat')}")
    lines.append(
        f"   scope: {os.environ.get('GIGACHAT_SCOPE', 'GIGACHAT_API_PERS')}"
    )
    lines.append(
        f"   таймаут запроса: "
        f"{os.environ.get('GIGACHAT_TIMEOUT', int(DEFAULT_GIGACHAT_TIMEOUT))} сек"
    )
    # Адрес API показываем явно: именно неверный/недоступный адрес даёт
    # ошибку [WinError 10060], и по отчёту это видно сразу.
    lines.append(
        f"   адрес API (base_url): "
        f"{os.environ.get('GIGACHAT_BASE_URL') or 'по умолчанию (адрес Сбера)'}"
    )
    lines.append(f"   режим запуска: {'собранный EXE' if STATUS_INFO.get('frozen') else 'исходники'}")
    lines.append("")

    # ---------- SSL ----------
    lines += _diag_certificate_section()
    lines.append("")

    # ---------- Логи ----------
    try:
        log_module = sys.modules.get("logger_config")
        log_path = getattr(log_module, "LOG_FILE", None) or str(LOG_FILE_HINT)
        warn = getattr(log_module, "_LOG_WARNING", None)
        lines.append(f"✅ Логгер: пишет в {log_path}")
        if warn:
            # _LOG_WARNING = «<причина> — логи перенесены в <папка> (…)».
            # Путь уже показан строкой выше, поэтому берём только причину.
            reason = str(warn).split(" — ", 1)[0].strip() or str(warn)
            lines.append(f"   ⚠️ {reason}")
    except Exception as exc:
        lines.append(f"⚠️ Логгер: не удалось определить путь ({exc})")
    lines.append("")

    # ---------- Настройки ----------
    try:
        lines.append(f"✅ Настройки: {SETTINGS_FILE}")
        lines.append(f"   регион: {CURRENT_REGION}, городов выбрано: {len(cities_by_region.get(CURRENT_REGION, []))}")
        for region in REGIONS:
            lines.append(
                f"   {region}: {len(cities_by_region.get(region, []))} шт. в памяти"
            )
    except Exception as exc:
        lines.append(f"⚠️ Настройки: не удалось прочитать ({exc})")
    lines.append("")

    # ---------- Мьютекс ----------
    if _mutex_handle is None and not _already_running:
        lines.append("✅ Мьютекс: свободен (это единственная копия программы)")
    elif _already_running:
        lines.append(
            "⚠️ Мьютекс: занят другой копией программы — эта работает вторым экземпляром"
        )
    else:
        lines.append("✅ Мьютекс: получен")
    lines.append("")

    # ---------- Окружение ----------
    lines.append(f"✅ Python: {platform.python_version()} ({sys.executable})")
    lines.append(f"✅ tkinter: OK | CustomTkinter: {ctk.__version__}")

    for module_name, human in (
        ("docx", "python-docx"),
        ("pyperclip", "pyperclip"),
        ("dotenv", "python-dotenv"),
        ("gigachat", "gigachat"),
        ("requests", "requests"),
        ("customtkinter", "customtkinter"),
    ):
        try:
            __import__(module_name)
            lines.append(f"✅ {human}: OK")
        except Exception as exc:
            lines.append(f"❌ {human}: НЕ установлен ({exc})")
            lines.append(f"   Решение: pip install -r requirements.txt")

    lines.append("")

    # ---------- SSL-ошибка в этой сессии ----------
    if has_ssl_error():
        lines.append("⚠️ В этой сессии уже была SSL-ошибка — см. раздел SSL выше")
        lines.append("")

    # ---------- Итог ----------
    ca_ok = False
    try:
        status = get_ca_bundle_status()
        ca_ok = bool(status.get("found")) or bool(status.get("explicit_valid"))
    except Exception:
        ca_ok = False

    if not GIGACHAT_READY:
        lines.append("📊 Итог: программа запустится, но ИИ выключен —")
        lines.append("         ключ GigaChat не задан. Нажмите «🔑 Ключ».")
    elif ca_ok:
        lines.append("📊 Итог: всё на месте — GigaChat должен отвечать.")
    else:
        lines.append("📊 Итог: программа запустится, но GigaChat не будет отвечать")
        lines.append("         до установки сертификата Минцифры (fallback на шаблоны).")

    return "\n".join(lines)


def _diag_show_progress(dialog: "ctk.CTkToplevel", text: "ctk.CTkTextbox") -> None:
    """Заглушка «идёт проверка», чтобы окно не выглядело зависшим."""
    text.configure(state="normal")
    text.delete("1.0", "end")
    text.insert("1.0", "🔍 Идёт диагностика, подождите…")
    text.configure(state="disabled")


def _show_diagnostics_result(dialog: "ctk.CTkToplevel", report: str) -> None:
    """Показывает готовый отчёт в уже открытом окне."""
    try:
        text = _diag_widgets.get("text")
        if text is None or not text.winfo_exists():
            return
        text.configure(state="normal")
        text.delete("1.0", "end")
        text.insert("1.0", report)
        text.configure(state="disabled")
        text.see("1.0")
    except Exception as exc:
        logger.warning("Не удалось показать диагностику: %s", exc)


def show_diagnostics():
    """
    Кнопка «📋 Диагностика»: показывает, что программа реально видит.

    Окно с прокруткой и возможностью скопировать текст: отчёт длинный
    (ключ, .env, сертификат Минцифры, логи, зависимости, итог), а
    messagebox такое не вмещает и не даёт выделить.

    Сбор отчёта идёт в фоновом потоке, а результат возвращается в главный
    поток через window.after(): обращаться к виджетам из другого потока
    нельзя.
    """
    dialog = _make_dialog("📋 Диагностика", 1000, 700)
    dialog.minsize(820, 520)

    ctk.CTkLabel(
        dialog,
        text="📋 Диагностика программы",
        font=FONT_H2,
        text_color=TEXT_MAIN,
    ).pack(pady=(16, 6), padx=16)

    text = ctk.CTkTextbox(dialog, wrap="word", font=FONT_MONO, fg_color="#0f1216")
    text.pack(padx=16, pady=6, fill="both", expand=True)

    _diag_widgets["text"] = text
    _diag_show_progress(dialog, text)

    buttons = ctk.CTkFrame(dialog, fg_color="transparent")
    buttons.pack(pady=(6, 16))

    def copy_report():
        try:
            content = text.get("1.0", "end").strip()
            pyperclip.copy(content)
            messagebox.showinfo("Диагностика", "Отчёт скопирован в буфер обмена.", parent=dialog)
        except Exception as exc:
            messagebox.showwarning(
                "Диагностика", f"Не удалось скопировать: {exc}", parent=dialog
            )

    def refresh():
        _diag_show_progress(dialog, text)
        _run_diagnostics_async(dialog)

    ctk.CTkButton(
        buttons,
        text="🔄 Обновить",
        font=FONT_SMALL,
        fg_color=COLOR_ACCENT,
        hover_color=COLOR_ACCENT_HOVER,
        command=refresh,
    ).pack(side="left", padx=5)

    ctk.CTkButton(
        buttons,
        text="📄 Скопировать отчёт",
        font=FONT_SMALL,
        fg_color=COLOR_TOOL,
        hover_color=COLOR_TOOL_HOVER,
        command=copy_report,
    ).pack(side="left", padx=5)

    ctk.CTkButton(
        buttons,
        text="🌐 Как исправить SSL",
        font=FONT_SMALL,
        fg_color=COLOR_DANGER,
        hover_color=COLOR_DANGER_HOVER,
        command=show_ssl_help,
    ).pack(side="left", padx=5)

    ctk.CTkButton(
        buttons,
        text="Закрыть",
        font=FONT_SMALL,
        fg_color=COLOR_TOOL,
        hover_color=COLOR_TOOL_HOVER,
        command=dialog.destroy,
    ).pack(side="left", padx=5)

    dialog.bind("<Escape>", lambda event: dialog.destroy())

    _run_diagnostics_async(dialog)
    text.see("1.0")


def _run_diagnostics_async(dialog: "ctk.CTkToplevel") -> None:
    """
    Собирает отчёт в фоновом потоке и отдаёт его в главный поток.

    Исключения не пробрасываются наружу: диагностика не должна ронять
    программу — при сбое в окне будет понятная ошибка.
    """
    def worker():
        try:
            report = _build_diagnostics_report()
        except Exception as exc:
            logger.error("Диагностика упала: %s", exc, exc_info=True)
            report = (
                "❌ Не удалось собрать сведения\n\n"
                f"{type(exc).__name__}: {exc}\n\n"
                f"Подробности в логе:\n{LOG_FILE_HINT}"
            )

        try:
            window.after(0, lambda: _show_diagnostics_result(dialog, report))
        except Exception:
            # Окно уже закрыто — это нормально, просто нечего показывать
            pass

    threading.Thread(target=worker, name="diagnostics", daemon=True).start()


# ============================================
# ВВОД КЛЮЧА ПРЯМО ИЗ ИНТЕРФЕЙСА
# ============================================


def enter_api_key():
    """
    Диалог ввода API-ключа GigaChat.

    Ключ сохраняется в .env (в .gitignore) и подхватывается без перезапуска.
    Сам ключ никогда не логируется и не показывается на экране.
    """
    global GIGACHAT_READY, GIGACHAT_MODEL

    dialog = _make_dialog("Ввод API-ключа GigaChat", 640, 430)
    dialog.resizable(False, False)
    dialog.grab_set()

    ctk.CTkLabel(
        dialog,
        text="🔑 API-ключ GigaChat",
        font=FONT_H2,
        text_color=TEXT_MAIN,
    ).pack(pady=(18, 8), padx=20)

    ctk.CTkLabel(
        dialog,
        text=(
            "Получите Authorization Key (base64):\n"
            "developers.sber.ru → Личный кабинет → Настройки\n\n"
            "Ключ сохранится в .env и НЕ попадёт в Git."
        ),
        font=FONT_SMALL,
        text_color=TEXT_DIM,
        justify="center",
    ).pack(padx=20, pady=(0, 8))

    # show="*" — ключ не отображается на экране
    key_entry = ctk.CTkEntry(dialog, width=460, font=FONT_SMALL, show="*")
    key_entry.pack(padx=20, pady=(6, 4))
    key_entry.focus_set()

    # Счётчик символов: благодаря ему видно, что вставка сработала. Раньше
    # при пустом поле (например, Ctrl+V не сработал) кнопка «Сохранить»
    # отвечала «Ключ не введён» — и было непонятно, почему.
    key_hint = ctk.CTkLabel(
        dialog,
        text="Длина ключа: 0 символов",
        font=FONT_SMALL,
        text_color=TEXT_DIM,
    )
    key_hint.pack()

    def _update_key_hint(event=None) -> None:
        """Показывает, сколько символов реально попало в поле."""
        try:
            length = len(key_entry.get())
        except Exception:
            return
        key_hint.configure(
            text=f"Длина ключа: {length} символов"
            + ("" if length else "  ← поле пустое"),
            text_color=TEXT_DIM if length else COLOR_TEXT_ERR,
        )

    def _paste_key(event=None) -> str:
        """
        Вставляет ключ из буфера обмена, очищая лишние символы.

        Зачем свой обработчик: при копировании из браузера в ключ часто
        попадают пробелы, кавычки и невидимые символы. Внешне ключ верный,
        а сервер отвечает «неверный ключ» — поэтому чистим сразу при вставке.
        """
        try:
            text = window.clipboard_get()
        except Exception:
            # Буфер пуст или занят другим приложением — отдаём событие Tkinter
            return ""
        key_entry.delete(0, "end")
        key_entry.insert(0, sanitize_key(text))
        _update_key_hint()
        return "break"

    key_entry.bind("<KeyRelease>", _update_key_hint)
    key_entry.bind("<Control-v>", _paste_key)
    key_entry.bind("<Control-V>", _paste_key)
    key_entry.bind("<<Paste>>", _paste_key)

    ctk.CTkLabel(
        dialog,
        text="(символы скрыты — это нормально, вставка через Ctrl+V)",
        font=FONT_SMALL,
        text_color=TEXT_DIM,
    ).pack(pady=(2, 10))

    def save():
        global GIGACHAT_READY, GIGACHAT_MODEL

        # sanitize_key убирает пробелы, кавычки и невидимые символы, которые
        # попадают в ключ при копировании. Раньше такой ключ сохранялся
        # «как есть», и GigaChat отвечал «неверный ключ» при верном значении.
        key = sanitize_key(key_entry.get())

        if not key:
            key_hint.configure(
                text="Длина ключа: 0 символов  ← поле пустое",
                text_color=COLOR_TEXT_ERR,
            )
            messagebox.showwarning(
                "Внимание",
                "Ключ не введён.\n\n"
                "Вставьте Authorization Key в поле (Ctrl+V) и нажмите "
                "«Сохранить» ещё раз.",
                parent=dialog,
            )
            key_entry.focus_set()
            return

        try:
            env_path = save_gigachat_key(key)
        except Exception as e:
            logger.error("Не удалось сохранить ключ в .env: %s", e)
            messagebox.showerror(
                "Ошибка", f"❌ Не удалось записать .env:\n{e}", parent=dialog
            )
            return

        # save_gigachat_key уже обновил os.environ и сбросил кэш клиента —
        # перечитываем флаг готовности и обновляем интерфейс.
        GIGACHAT_READY = is_gigachat_configured()
        GIGACHAT_MODEL = get_status_info().get("gigachat_model", "GigaChat")
        AICache.clear()

        dialog.destroy()

        # Обновляем индикатор в статус-баре
        try:
            status_label.configure(
                text=f"GigaChat: ✅ GigaChat готов ({GIGACHAT_MODEL})",
                text_color=COLOR_TEXT_OK,
            )
        except NameError:
            pass

        # Плашка «❌ Ключ не задан» больше не нужна — ключ только что сохранён
        try:
            update_key_banner()
        except Exception as exc:
            logger.debug("Не удалось обновить плашку ключа: %s", exc)

        logger.info(
            "Ключ GigaChat сохранён в %s и активирован (значение не логируется)",
            env_path,
        )

        if messagebox.askyesno(
            "Готово!",
            f"✅ Ключ сохранён в:\n{env_path}\n\n"
            "🔒 Убедитесь, что .env в .gitignore (он там по умолчанию).\n\n"
            "Выполнить проверку подключения сейчас?",
        ):
            warmup_async()
        else:
            refresh_preview_async()

    buttons = ctk.CTkFrame(dialog, fg_color="transparent")
    buttons.pack(pady=(6, 18), padx=20, fill="x")

    ctk.CTkButton(
        buttons,
        text="Отмена",
        font=FONT_SMALL,
        width=140,
        fg_color=COLOR_TOOL,
        hover_color=COLOR_TOOL_HOVER,
        command=dialog.destroy,
    ).pack(side="left")

    ctk.CTkButton(
        buttons,
        text="💾 Сохранить",
        font=FONT_SMALL_BOLD,
        width=160,
        fg_color=COLOR_PRIMARY,
        hover_color=COLOR_PRIMARY_HOVER,
        command=save,
    ).pack(side="right")

    dialog.bind("<Return>", lambda event: save())
    dialog.bind("<Escape>", lambda event: dialog.destroy())


# ============================================
# ЗАКРЫТИЕ ОКНА И ГЛАВНЫЙ ЦИКЛ
# ============================================
# Определяем эти функции ДО построения интерфейса: обработчик закрытия
# подключается к окну и к кнопке «ВЫХОД» в момент их создания.
# Явный обработчик нужен, чтобы закрытие всегда проходило через destroy()
# и в лог попадала понятная причина, а не молчаливый выход из mainloop().
# Раньше в логе оставалась одна строка «Программа завершена» без объяснения —
# именно это и выглядело как «программа закрылась сама».

_window_closed = {"value": False}

# Заполняются при старте главного цикла (см. _run_gui)
_mutex_handle = None
_already_running = False

# False до конца построения окна: обработчик смены вкладки не должен
# срабатывать на программную установку вкладки при старте.
_gui_ready = {"value": False}


def on_closing() -> None:
    """Корректно закрывает окно и пишет причину в лог."""
    if _window_closed["value"]:
        return
    _window_closed["value"] = True
    logger.info("Закрытие окна по команде пользователя (крестик/ВЫХОД)")
    try:
        window.destroy()
    except Exception as exc:
        logger.warning("Ошибка при закрытии окна: %s", exc)


def on_exit_button() -> None:
    """Кнопка «ВЫХОД»: тот же путь, что и крестик."""
    on_closing()


def _acquire_single_instance_lock():
    """
    Пытается захватить именованный мьютекс Windows.

    Две одновременно запущенные копии писали в один лог-файл и мешали друг
    другу, а вторая копия могла упасть ещё до появления окна.

    ВАЖНО: используем use_last_error=True и ctypes.get_last_error().
    Обычный kernel32.GetLastError() здесь ненадёжен — Python успевает
    сбросить код ошибки до того, как мы его прочитаем.

    Returns:
        (handle, already_running) — handle мьютекса (или None) и признак
        того, что программа уже запущена.
    """
    if os.name != "nt":
        return None, False

    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.argtypes = [
            wintypes.LPVOID,
            wintypes.BOOL,
            wintypes.LPCWSTR,
        ]
        kernel32.CreateMutexW.restype = wintypes.HANDLE

        handle = kernel32.CreateMutexW(None, True, "Global\\GeneratorKP_GigaChat")
        last_error = ctypes.get_last_error()

        if not handle:
            logger.debug("CreateMutexW не удался (код %s)", last_error)
            return None, False

        # ERROR_ALREADY_EXISTS = 183: мьютекс уже создан другой копией
        if last_error == 183:
            return handle, True

        return handle, False
    except Exception as exc:
        logger.debug("Проверка единственного экземпляра недоступна: %s", exc)
        return None, False


# ============================================
# ГРАФИЧЕСКИЙ ИНТЕРФЕЙС (CustomTkinter)
# ============================================
#
# Раскладка:
#   ┌──────────────────────────────────────────────────────────┐
#   │ шапка: заголовок + панель инструментов                   │
#   │ плашка «ключ не задан» / плашка SSL (по необходимости)   │
#   ├───────────────────┬──────────────────────────────────────┤
#   │ настройки (380px) │ предпросмотр (растягивается)         │
#   │ вкладки регионов  │ «показать новый вариант» + прогресс  │
#   │ города, оплата,   │ текст КП                             │
#   │ креативность      │ копировать / Word / 3 варианта       │
#   ├───────────────────┴──────────────────────────────────────┤
#   │ статус-бар: GigaChat + путь к логу                       │
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
    (кнопки и статус-бар) уходил под панель задач, и дотянуться до него было
    нельзя. Плюс Windows открывает новое окно со сдвигом вниз, из-за чего та же
    проблема возникала и на 1920x1080.

    Высота берётся с запасом 88 % от экрана: остальное занимает панель задач.
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
        # Делим на три, а не на два: окно встаёт чуть выше центра и не
        # упирается в панель задач снизу.
        y = max((screen_h - height) // 3, 0)

        win.geometry(f"{width}x{height}+{x}+{y}")
        win.minsize(min(min_width, width), min(min_height, height))
        logger.info("Окно: %dx%d (экран %dx%d)", width, height, screen_w, screen_h)
    except Exception as exc:  # pragma: no cover — размер окна не критичен
        logger.debug("Не удалось подогнать окно под экран: %s", exc)
        win.geometry(f"{want_width}x{want_height}")


window = ctk.CTk()
window.title("Генератор КП — Автоперевозки (GigaChat)")
window.configure(fg_color=BG_WINDOW)
_fit_window_to_screen(window, 1180, 820, 1000, 660)

# Шрифты. Segoe UI — системный шрифт Windows: есть кириллица и подстановка
# эмодзи, чего не хватает встроенному в CustomTkinter Roboto.
FONT_H1 = ctk.CTkFont(family="Segoe UI", size=22, weight="bold")
FONT_H2 = ctk.CTkFont(family="Segoe UI", size=15, weight="bold")
FONT_UI = ctk.CTkFont(family="Segoe UI", size=14)
FONT_SMALL_BOLD = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
FONT_SMALL = ctk.CTkFont(family="Segoe UI", size=12)
FONT_TINY = ctk.CTkFont(family="Segoe UI", size=11)
FONT_MONO = ctk.CTkFont(family="Consolas", size=12)

window.grid_columnconfigure(0, weight=1)
window.grid_rowconfigure(3, weight=1)  # растягивается только тело окна

# ---------- Шапка ----------
# Заголовок и панель инструментов стоят в ДВЕ строки: в одну они не помещаются,
# и заголовок обрезался по краям (проверено — не хватало ~150 px).
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
    text="Лоты машин для автовозов · текст собирается из данных лота и настроек",
    font=FONT_TINY,
    text_color=TEXT_DIM,
).pack(anchor="w")

# Панель инструментов: ИИ вкл/выкл, прогрев, ключ, кэш, диагностика, помощь, выход
toolbar = ctk.CTkFrame(header_frame, fg_color="transparent")
toolbar.grid(row=1, column=0, sticky="e", padx=16, pady=(4, 10))

use_ai_var = tk.IntVar(value=1 if USE_AI else 0)
ai_switch = ctk.CTkSwitch(
    toolbar,
    text=f"🤖 ИИ: {'ВКЛ' if USE_AI else 'ВЫКЛ'}",
    font=FONT_SMALL_BOLD,
    variable=use_ai_var,
    onvalue=1,
    offvalue=0,
    progress_color=COLOR_AI_ON,
    button_color="#e9ecf1",
    button_hover_color="#cfd6e0",
    command=toggle_ai,
)
ai_switch.pack(side="left", padx=(0, 10))

ctk.CTkLabel(
    toolbar,
    text=f"Модель: {GIGACHAT_MODEL}",
    font=FONT_TINY,
    text_color=TEXT_DIM,
).pack(side="left", padx=(0, 10))


def _tool_button(text: str, command, color: str = COLOR_TOOL, hover: str = COLOR_TOOL_HOVER, width: int = 0):
    """Кнопка панели инструментов одного размера — шапка не «плывёт»."""
    button = ctk.CTkButton(
        toolbar,
        text=text,
        font=FONT_SMALL,
        width=width,
        height=32,
        fg_color=color,
        hover_color=hover,
        command=command,
    )
    button.pack(side="left", padx=3)
    return button


_tool_button("⚡ Прогрев", warmup_async, COLOR_ACCENT, COLOR_ACCENT_HOVER)
_tool_button("🔑 Ключ", enter_api_key, COLOR_ACCENT, COLOR_ACCENT_HOVER)
_tool_button("🗑️ Кэш", clear_cache, COLOR_WARN, COLOR_WARN_HOVER)
_tool_button("📋 Диагностика", show_diagnostics)
_tool_button("❓ Помощь", show_gigachat_help)
_tool_button("✖ Выход", on_exit_button, COLOR_DANGER, COLOR_DANGER_HOVER)

# ---------- Плашки предупреждений ----------
# Плашка создаётся всегда, а видимостью управляет update_key_banner():
# после ввода ключа кнопкой «🔑 Ключ» она должна исчезать без перезапуска.
key_warning_label = ctk.CTkLabel(
    window,
    text="⚠️ Ключ GigaChat не задан. Нажмите 🔑 «Ключ» в шапке, чтобы вставить ключ",
    font=FONT_SMALL_BOLD,
    text_color="#7f1d1d",
    fg_color="#ffe0e0",
    corner_radius=8,
    height=30,
)


def update_key_banner() -> None:
    """
    Показывает или скрывает плашку «Ключ GigaChat не задан».

    Нужна после сохранения ключа из интерфейса: раньше плашка оставалась
    на экране до перезапуска и выглядела как «ключ не сохранился», хотя
    .env уже был обновлён и ключ работал.
    """
    if GIGACHAT_READY:
        if key_warning_label.winfo_manager():
            key_warning_label.grid_forget()
    else:
        if not key_warning_label.winfo_manager():
            key_warning_label.grid(row=1, column=0, sticky="ew", padx=16, pady=(8, 0))


# Плашка про SSL. Показывается в двух случаях:
#   1. СРАЗУ при старте — сертификата Минцифры нет, а ключ задан. Тогда
#      GigaChat заведомо не ответит, и пользователь узнаёт об этом до
#      первого нажатия «Сгенерировать» (раньше — только по факту ошибки).
#   2. После падения запроса по сертификату (has_ssl_error()).
# Плашка немодальная: она ничего не блокирует и не мешает работать.
ssl_banner = ctk.CTkFrame(window, fg_color="#ffcdd2", corner_radius=8)
ssl_banner.grid_columnconfigure(0, weight=1)

ssl_label = ctk.CTkLabel(
    ssl_banner,
    text="⚠️ SSL: не удалось проверить сертификат GigaChat",
    font=FONT_SMALL_BOLD,
    text_color="#7f1d1d",
    justify="left",
)
ssl_label.grid(row=0, column=0, sticky="w", padx=(12, 8), pady=6)

ctk.CTkButton(
    ssl_banner,
    text="Как исправить",
    font=FONT_TINY,
    width=120,
    height=26,
    fg_color=COLOR_DANGER,
    hover_color=COLOR_DANGER_HOVER,
    command=show_ssl_help,
).grid(row=0, column=1, sticky="e", padx=(4, 10), pady=6)


def _certificate_present() -> bool:
    """
    Есть ли корневой сертификат Минцифры (файл, без сетевых проверок).

    Нужна и для плашки при старте, и для отчёта диагностики.
    """
    try:
        status = get_ca_bundle_status()
    except Exception as exc:
        logger.debug("Не удалось проверить наличие сертификата: %s", exc)
        return False
    return bool(status.get("found")) or bool(status.get("explicit_valid"))


def update_ssl_banner(cert_ok: Optional[bool] = None) -> None:
    """
    Обновляет состояние плашки SSL.

    Показывает предупреждение, если сертификат Минцифры отсутствует (при
    заданном ключе) ИЛИ если запрос уже упал по SSL. Скрывает, когда всё
    в порядке.

    Args:
        cert_ok: результат проверки сертификата. None — проверить здесь
                 (при старте проверка уже сделана и передаётся готовой,
                 чтобы не писать одно предупреждение в лог дважды).

    Проверяем winfo_manager(), а НЕ winfo_ismapped(): у окна, которое ещё
    не отрисовано (или свёрнуто), ismapped() возвращает 0 даже когда плашка
    размещена — из-за этого она не убиралась после успешного запроса.
    Плашку мы только ставим в grid, поэтому менеджер 'grid' == «показана».

    Вызывается из главного потока: менять виджеты из фонового потока
    в Tkinter нельзя.
    """
    try:
        if cert_ok is None:
            cert_ok = _certificate_present() if GIGACHAT_READY else True

        need_banner = has_ssl_error() or (GIGACHAT_READY and not cert_ok)

        if need_banner:
            if has_ssl_error():
                ssl_label.configure(
                    text=(
                        "⚠️ SSL: не удалось проверить сертификат GigaChat.\n"
                        "GigaChat не отвечает — текст КП берётся из шаблонов."
                    )
                )
            else:
                ssl_label.configure(
                    text=(
                        "⚠️ SSL-сертификат Минцифры не найден. "
                        "GigaChat может не отвечать. [Как исправить]"
                    )
                )

        shown = bool(ssl_banner.winfo_manager())
        if need_banner:
            if not shown:
                ssl_banner.grid(row=2, column=0, sticky="ew", padx=16, pady=(8, 0))
        else:
            if shown:
                ssl_banner.grid_forget()
    except Exception as exc:
        logger.debug("Не удалось обновить SSL-плашку: %s", exc)


# ---------- Тело окна: две колонки ----------
body = ctk.CTkFrame(window, fg_color="transparent")
body.grid(row=3, column=0, sticky="nsew", padx=12, pady=(10, 6))
body.grid_rowconfigure(0, weight=1)
# Левая колонка фиксированной ширины, правая растягивается.
# ВАЖНО: ширину держит именно minsize колонки — grid_propagate(False) у
# CustomTkinter-фрейма перестаёт работать, как только внутри появляются виджеты
# (проверено: фрейм «съёживался» до 364 px по размеру вкладок).
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
    # Высота заведомо небольшая: список городов всё равно прокручивается, а
    # окно на маленьком экране должно помещаться целиком (растёт по pack).
    height=150,
    fg_color=BG_WINDOW,
    segmented_button_selected_color=COLOR_PRIMARY,
    segmented_button_selected_hover_color=COLOR_PRIMARY_HOVER,
)
notebook.pack(fill="both", expand=True, padx=12, pady=(0, 8))

tab_south = notebook.add(TAB_TITLES["Юга"])
tab_vladivostok = notebook.add(TAB_TITLES["Владивосток"])
tab_east = notebook.add(TAB_TITLES["Восток"])

city_vars: Dict[str, tk.IntVar] = {}
_city_checks: Dict[str, "ctk.CTkCheckBox"] = {}


def _make_city_checkbox(parent, city: str, italic: bool = False) -> None:
    """Один флажок города. Значения выставляет apply_region_selection()."""
    var = tk.IntVar(value=0)
    city_vars[city] = var
    checkbox = ctk.CTkCheckBox(
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
    )
    checkbox.pack(anchor="w", padx=6, pady=3)
    _city_checks[city] = checkbox


# ===== ЮГА =====
south_list = ctk.CTkScrollableFrame(tab_south, fg_color="transparent")
south_list.pack(fill="both", expand=True)
ctk.CTkLabel(
    south_list,
    text="📍 Выберите города:",
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
    text="📍 Выберите города:",
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
# Отдельный выбор, потому что из макета лота оплата нигде не спрашивалась:
# программа всегда писала «безнал». Формулировка попадает и в промпт ИИ,
# и в шаблонный текст, и в сохранённые настройки.
payment_frame = ctk.CTkFrame(left_column, fg_color="transparent")
payment_frame.pack(fill="x", padx=14, pady=(2, 6))

ctk.CTkLabel(
    payment_frame,
    text="💰 Оплата:",
    font=FONT_SMALL_BOLD,
    text_color=TEXT_MAIN,
).pack(side="left", padx=(0, 8))

payment_var = tk.StringVar(value=PAYMENT_TYPES[PAYMENT_KEY][0])
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

# ===== КРЕАТИВНОСТЬ =====
creativity_frame = ctk.CTkFrame(left_column, fg_color="transparent")
creativity_frame.pack(fill="x", padx=14, pady=(0, 4))

creativity_label = ctk.CTkLabel(
    creativity_frame,
    text=f"Креативность: {int(CREATIVITY * 100)}%",
    font=FONT_SMALL_BOLD,
    text_color=TEXT_MAIN,
)
creativity_label.pack(anchor="w")

creativity_slider = ctk.CTkSlider(
    creativity_frame,
    from_=0.2,
    to=1.0,
    number_of_steps=16,
    command=update_creativity,
    progress_color=COLOR_PRIMARY,
    button_color=COLOR_PRIMARY,
    button_hover_color=COLOR_PRIMARY_HOVER,
)
creativity_slider.set(CREATIVITY)
creativity_slider.pack(fill="x", pady=(4, 2))

ctk.CTkLabel(
    creativity_frame,
    text="0.2 — шаблонно, 1.0 — креативно",
    font=FONT_TINY,
    text_color=TEXT_DIM,
).pack(anchor="w")

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
right_column.grid_rowconfigure(2, weight=1)
right_column.grid_columnconfigure(0, weight=1)

preview_top = ctk.CTkFrame(right_column, fg_color="transparent")
preview_top.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 4))

refresh_button = ctk.CTkButton(
    preview_top,
    text="🔄 ПОКАЗАТЬ НОВЫЙ ВАРИАНТ",
    font=FONT_SMALL_BOLD,
    height=42,
    width=260,
    fg_color=COLOR_VIOLET,
    hover_color=COLOR_VIOLET_HOVER,
    command=refresh_preview_async,
)
refresh_button.pack(side="left")

gen_status_label = ctk.CTkLabel(
    preview_top,
    text="",
    font=FONT_SMALL_BOLD,
    text_color=COLOR_TEXT_WARN,
)
gen_status_label.pack(side="left", padx=12)

# Прогресс-бар «ползёт», пока идёт фоновая генерация. Изначально скрыт:
# постоянная анимация в окне только отвлекает.
progress_bar = ctk.CTkProgressBar(right_column, mode="indeterminate", height=6)
progress_bar.set(0)

# Текстовое поле с предпросмотром — растягивается вместе с окном.
# Высота минимальная: на маленьком экране окно должно помещаться целиком,
# на большом поле растягивается само (grid, row weight).
text_preview = ctk.CTkTextbox(
    right_column,
    wrap="word",
    font=FONT_UI,
    height=140,
    fg_color="#0f1216",
    text_color=TEXT_MAIN,
    corner_radius=10,
)
text_preview.grid(row=2, column=0, sticky="nsew", padx=16, pady=(8, 4))
text_preview.configure(state="disabled")

ctk.CTkLabel(
    right_column,
    text="💡 Для Востока: отметьте «Без города» — маршрут будет описан направлением, без перечисления городов",
    font=FONT_TINY,
    text_color=TEXT_DIM,
).grid(row=3, column=0, sticky="w", padx=16, pady=(0, 4))

button_frame = ctk.CTkFrame(right_column, fg_color="transparent")
button_frame.grid(row=4, column=0, sticky="ew", padx=16, pady=(4, 16))
button_frame.grid_columnconfigure(0, weight=1)
button_frame.grid_columnconfigure(1, weight=1)
button_frame.grid_columnconfigure(2, weight=1)

copy_button = ctk.CTkButton(
    button_frame,
    text="📋 КОПИРОВАТЬ",
    font=FONT_SMALL_BOLD,
    height=46,
    fg_color=COLOR_PRIMARY,
    hover_color=COLOR_PRIMARY_HOVER,
    command=copy_to_clipboard,
)
copy_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))

word_button = ctk.CTkButton(
    button_frame,
    text="📄 СОХРАНИТЬ WORD",
    font=FONT_SMALL_BOLD,
    height=46,
    fg_color=COLOR_ACCENT,
    hover_color=COLOR_ACCENT_HOVER,
    command=save_to_word,
)
word_button.grid(row=0, column=1, sticky="ew", padx=6)

variants_button = ctk.CTkButton(
    button_frame,
    text="📚 3 ВАРИАНТА",
    font=FONT_SMALL_BOLD,
    height=46,
    fg_color=COLOR_WARN,
    hover_color=COLOR_WARN_HOVER,
    command=save_multiple_variants,
)
variants_button.grid(row=0, column=2, sticky="ew", padx=(6, 0))

# ---------- Статус-бар ----------
status_bar = ctk.CTkFrame(window, corner_radius=0, fg_color=BG_STATUS)
status_bar.grid(row=4, column=0, sticky="ew")
status_bar.grid_columnconfigure(0, weight=1)

# Индикатор статуса GigaChat
if GIGACHAT_READY:
    status_text = f"✅ GigaChat готов ({GIGACHAT_MODEL})"
    status_color = COLOR_TEXT_OK
else:
    status_text = "❌ Ключ не задан — используется шаблонный текст"
    status_color = COLOR_TEXT_ERR

status_label = ctk.CTkLabel(
    status_bar,
    text=f"GigaChat: {status_text}",
    font=FONT_SMALL,
    text_color=status_color,
)
status_label.grid(row=0, column=0, sticky="w", padx=14, pady=6)

log_box = ctk.CTkFrame(status_bar, fg_color="transparent")
log_box.grid(row=0, column=1, sticky="e", padx=14, pady=6)

ctk.CTkLabel(
    log_box,
    text=f"📄 Лог: {LOG_FILE_HINT}",
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

# Показать сохранённый регион и его выбор городов
notebook.set(TAB_TITLES[CURRENT_REGION])
apply_region_selection(CURRENT_REGION)
_active_region = CURRENT_REGION
_gui_ready["value"] = True

logger.info("Интерфейс загружен")

# Плашку SSL выставляем сразу после построения окна: если сертификата нет,
# пользователь видит предупреждение ещё до первого запроса к GigaChat.
# Повторная (уже без записи в лог) проверка — в _run_gui перед mainloop().
update_key_banner()
update_ssl_banner()


def _startup_ssl_check() -> None:
    """
    Проверка сертификата при старте: показать проблему ДО первого запроса.

    Вызывается из _run_gui перед mainloop(). Сертификат проверяется один
    раз, чтобы в лог не попадало одно и то же предупреждение дважды.
    Ошибки проверки не должны мешать запуску окна.
    """
    try:
        # Поиск SSL-инспекции идёт в фоне: он запускает PowerShell и на
        # «холодной» машине занимает секунды. Ждать его при старте нельзя —
        # окно должно открываться сразу, а результат появится в диагностике.
        start_interception_scan()
    except Exception as exc:  # pragma: no cover
        logger.debug("Фоновый поиск SSL-инспекции не запустился: %s", exc)

    try:
        cert_ok = _certificate_present()

        if GIGACHAT_READY and not cert_ok:
            interception = detect_ssl_interception()
            if interception:
                # Корень антивируса программа добавляет в набор сама, поэтому
                # без сертификата Минцифры GigaChat всё равно, скорее всего,
                # не ответит — предупреждаем и называем виновника.
                logger.warning(
                    "SSL: корневой сертификат Минцифры не найден, а в системе "
                    "включена SSL-инспекция (%s). GigaChat, скорее всего, не "
                    "ответит — программа будет работать на шаблонах. "
                    "Сертификат: %s",
                    ", ".join(s.split(",")[0] for s in interception),
                    get_ca_bundle_status().get("url"),
                )
            else:
                logger.warning(
                    "SSL: корневой сертификат Минцифры не найден — GigaChat, "
                    "скорее всего, не ответит. Программа будет работать на "
                    "шаблонах. Скачать сертификат: %s",
                    get_ca_bundle_status().get("url"),
                )

        update_ssl_banner(cert_ok)
    except Exception as exc:  # pragma: no cover
        logger.debug("Стартовая проверка сертификата не удалась: %s", exc)


def _run_gui() -> int:
    """
    Запускает главный цикл приложения.

    ВАЖНО: главный цикл вынесен в функцию и вызывается только при прямом
    запуске файла (и только под __name__ == "__main__"). Раньше mainloop()
    вызывался на уровне модуля, поэтому ЛЮБОЙ импорт этого файла
    (диагностика, тест, внешний запускатель) захватывал окно и висел —
    отладить проблему запуска было нельзя.

    Returns:
        Код возврата процесса.
    """
    # Обработчик крестика: закрытие всегда идёт через destroy() и попадает в лог
    window.protocol("WM_DELETE_WINDOW", on_closing)

    # Плашка SSL проверяется ещё раз перед показом окна: так пользователь
    # видит проблему с сертификатом сразу, а не после нажатия «Сгенерировать»
    _startup_ssl_check()

    # Первый предпросмотр и прогрев авторизации — в фоне, окно открывается сразу
    window.after(100, refresh_preview_async)

    if GIGACHAT_READY:
        threading.Thread(
            target=warmup_gigachat, name="gigachat-startup-warmup", daemon=True
        ).start()
    else:
        logger.warning(
            "Ключ GigaChat не задан — используется шаблонный текст. "
            "Введите ключ кнопкой «🔑 Ключ» или в файле .env"
        )

    global _mutex_handle, _already_running
    _mutex_handle, _already_running = _acquire_single_instance_lock()
    if _already_running:
        # Не мешаем работать, но честно говорим в логе, что это вторая копия
        logger.warning(
            "Обнаружена уже запущенная копия программы — работаю вторым экземпляром"
        )

    try:
        window.mainloop()
    except SystemExit as exc:
        logger.info("Выход по SystemExit: %s", exc)
    except BaseException as exc:
        # Падение в главном цикле больше не выглядит как «окно мигнуло и исчезло»
        logger.critical("КРИТИЧЕСКАЯ ОШИБКА в главном цикле", exc_info=True)
        try:
            messagebox.showerror(
                "Критическая ошибка",
                "❌ Программа столкнулась с ошибкой и будет закрыта.\n\n"
                f"{type(exc).__name__}: {exc}\n\n"
                f"Подробности записаны в лог:\n{LOG_FILE_HINT}",
            )
        except Exception:
            traceback.print_exception(type(exc), exc, exc.__traceback__)
        return 1
    else:
        # mainloop вернулся нормально — это штатное закрытие окна
        if not _window_closed["value"]:
            logger.warning(
                "Главный цикл завершился БЕЗ команды закрытия окна — "
                "проверьте, не закрыла ли программу другая программа "
                "(антивирус, диспетчер задач)"
            )

    logger.info("Программа завершена")
    return 0


if __name__ == "__main__":
    raise SystemExit(_run_gui())
