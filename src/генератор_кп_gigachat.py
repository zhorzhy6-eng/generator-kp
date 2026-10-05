#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Генератор коммерческих предложений с ИИ (GigaChat)
Версия: 1.0

Отличия от версии с Ollama:
  * Генерация через облачный GigaChat (src/llm_provider.py)
  * Предпросмотр строится в отдельном потоке — окно не «зависает»
  * Если GigaChat недоступен/вернул None — автоматический fallback на шаблоны

Безопасность: ключ читается только из .env и НИКОГДА не пишется в лог.
"""

import pyperclip
import tkinter as tk
from tkinter import messagebox, Checkbutton, IntVar, ttk
from datetime import datetime
import os
import random
import json
import re
import sys
import threading
import time
import traceback
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Импортируем общий модуль: пути проекта и логгер
from logger_config import setup_logger, SETTINGS_PATH, PROJECT_ROOT, LOG_FILE as LOG_FILE_HINT

# Единый провайдер LLM (GigaChat + Ollama)
from llm_provider import (
    DEFAULT_GIGACHAT_TIMEOUT,
    DEFAULT_SYSTEM_PROMPT,
    detect_ssl_interception,
    generate_gigachat,
    get_ca_bundle_status,
    get_env_write_path,
    get_ssl_help,
    get_status_info,
    has_ssl_error,
    is_gigachat_configured,
    sanitize_key,
    save_gigachat_key,
    start_interception_scan,
    warmup_gigachat,
)

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
logger.info("ЗАПУСК ГЕНЕРАТОРА КП С GIGACHAT")
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

# Короткий системный промпт: меньше токенов на вход — быстрее ответ
GIGACHAT_SYSTEM_PROMPT = (
    DEFAULT_SYSTEM_PROMPT + " Отвечай строго 3 предложениями и коротким призывом к действию."
)

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
    """Загружает настройки из файла"""
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def save_settings(selected_cities, region, use_ai=True, creativity=0.8):
    """Сохраняет настройки в файл. Ключи GigaChat здесь НЕ хранятся."""
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "cities": selected_cities,
                    "region": region,
                    "use_ai": use_ai,
                    "model": GIGACHAT_MODEL,
                    "creativity": creativity,
                },
                f,
                ensure_ascii=False,
                indent=2,
            )
        logger.info(
            "Настройки сохранены: регион=%s, городов=%d", region, len(selected_cities)
        )
    except Exception as e:
        logger.error("Ошибка сохранения настроек: %s", e)


# Загружаем сохранённые настройки
saved = load_settings()
if saved:
    SELECTED_CITIES = saved.get("cities", [])
    CURRENT_REGION = saved.get("region", "Юга")
    USE_AI = saved.get("use_ai", True)
    CREATIVITY = saved.get("creativity", 0.8)
else:
    SELECTED_CITIES = []
    CURRENT_REGION = "Юга"
    USE_AI = True
    CREATIVITY = 0.8

logger.info("Загружена креативность: %s", CREATIVITY)

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

# ============================================
# ВАРИАНТЫ СТИЛЕЙ ДЛЯ FALLBACK-ШАБЛОНОВ
# ============================================

OPENERS = [
    "🔥 Лоты сформированы!",
    "🚛 Свежие машины в наличии!",
    "💥 Отличные предложения!",
    "⭐ Авто готовы к отправке!",
    "📦 Машины на месте!",
    "🚀 Горячие лоты!",
    "⚡ Быстрая загрузка!",
    "🎯 Точечные предложения!",
    "🏆 Топ-лоты!",
    "✨ Свежее поступление!",
]

STATUSES = [
    "машины в наличии, загрузка быстрая.",
    "авто готовы к отправке в любой момент.",
    "лоты полностью укомплектованы.",
    "машины на площадке, ждут загрузки.",
    "транспорт готов к выезду.",
    "все авто в наличии, документы в порядке.",
    "машины уже на месте, готовы к погрузке.",
    "лоты сформированы, отправка сегодня.",
]

PAYMENTS = [
    "Работаем по безналу, сроки обсуждаем.",
    "Оплата безналичная, возможна предоплата.",
    "Гибкие условия оплаты, безналичный расчёт.",
    "Предоплата или полная оплата — на ваш выбор.",
    "Безналичный расчёт, работаем с юрлицами.",
    "Оплата по безналу, договор сразу.",
]

LEGALS = [
    "Заключаем договор, гарантируем надёжность.",
    "Работаем по договору, все выплаты в срок.",
    "Официальное оформление, без задержек.",
    "Договор сразу, платим чётко по графику.",
    "Юридически чистая сделка, договор гарантирует.",
    "Официально, по договору, без задержек.",
]

CTA = [
    "Ждём ваших ставок!",
    "Пишите в ЛС!",
    "Обсудим детали!",
    "Звоните, договариваемся!",
    "Все вопросы в личку!",
    "Ждём ваших сообщений!",
    "Пишите уже сегодня!",
]

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

        required_keywords = ["машин", "ло", "груз", "авто", "транспорт"]
        if not any(kw in text.lower() for kw in required_keywords):
            return False, "Нет ключевых слов по теме"

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


def _build_prompt(region: str, cities: List[str]) -> str:
    """
    Короткий промпт: меньше входных токенов — быстрее ответ.

    Держим его компактным осознанно: длинные инструкции заметно увеличивают
    время генерации, а качество для такого короткого текста не растёт.
    """
    region_map = {
        "Юга": "юг России (Ростов, Краснодар, Новороссийск)",
        "Владивосток": "из Владивостока по России",
        "Восток": "Урал, Сибирь, Дальний Восток",
    }
    region_desc = region_map.get(region, region)
    cities_str = ", ".join(cities) if cities else "все города"

    return (
        f"Напиши коммерческое предложение по автоперевозкам. "
        f"Маршрут: Москва → {region_desc}. "
        f"Города: {cities_str}. "
        f"Упомяни: машины в наличии, оплата безналом, работа по договору. "
        f"3 предложения. В конце призыв к действию."
    )


def generate_with_gigachat(region: str, cities: List[str]) -> Optional[str]:
    """
    Генерация текста через GigaChat.

    Returns:
        Валидный текст или None (тогда вызывающий код использует шаблоны).
    """
    if not GIGACHAT_READY:
        return None

    prompt = _build_prompt(region, cities)

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
# ФУНКЦИИ ГЕНЕРАЦИИ
# ============================================


def get_selected_cities():
    return [city for city, var in city_vars.items() if var.get() == 1]


def get_current_region():
    tab = notebook.index(notebook.select())
    if tab == 0:
        return "Юга"
    elif tab == 1:
        return "Владивосток"
    else:
        return "Восток"


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


def build_template_text(cities_text: str) -> str:
    """Fallback: собирает текст из случайных шаблонных фраз."""
    opener = random.choice(OPENERS)
    status = random.choice(STATUSES)
    payment = random.choice(PAYMENTS)
    legal = random.choice(LEGALS)
    cta = random.choice(CTA)

    variants = [
        f"{opener} {status} {payment} {legal} {cta}",
        f"{opener} {payment} {status} {legal} {cta}",
        f"{opener} {status} {legal} {payment} {cta}",
    ]

    result = random.choice(variants)

    if cities_text:
        result = f"{result}\n\n{cities_text}"

    return result


def generate_text_sync(region: str, cities: List[str], use_cache: bool = True) -> str:
    """
    Синхронно генерирует текст: сначала GigaChat, при неудаче — шаблоны.

    ВНИМАНИЕ: блокирующая функция (может занять 2–4 сек).
    Из GUI вызывать только через generate_text_async().
    """
    route_cities = [c for c in cities if c != CREATIVE_MODE]
    cities_text = format_cities_list(route_cities) if route_cities else ""

    if USE_AI and GIGACHAT_READY:
        city_str = ",".join(sorted(route_cities))
        # Уникальный ключ: кэш не должен «залипать» на одном варианте
        cache_key = (
            f"gigachat_{region}_{city_str}_{int(CREATIVITY * 10)}_"
            f"{random.randint(1, 1000)}"
        )

        if use_cache:
            cached = AICache.get(cache_key)
            if cached:
                logger.info("✅ Использован кэшированный текст GigaChat")
                return cached

        logger.info("Генерация GigaChat: регион=%s, городов=%d", region, len(route_cities))
        ai_text = generate_with_gigachat(region, route_cities)

        if ai_text:
            if use_cache:
                AICache.set(cache_key, ai_text)
            # Парная запись к «🔄 Использован шаблонный текст» ниже.
            # Без неё по логу нельзя было понять, что текст пришёл от ИИ:
            # успешный путь не логировался вообще, и казалось, что программа
            # всегда работает на шаблонах.
            logger.info("✅ Использован ИИ-текст")
            if cities_text:
                return f"{ai_text}\n\n{cities_text}"
            return ai_text

    elif USE_AI and not GIGACHAT_READY:
        logger.warning("ИИ включён, но GigaChat не настроен — использую шаблоны")

    logger.info("🔄 Использован шаблонный текст (fallback)")
    return build_template_text(cities_text)


# ============================================
# АСИНХРОННАЯ ГЕНЕРАЦИЯ ДЛЯ GUI (threading)
# ============================================

_generation_lock = threading.Lock()
_generation_in_progress = False


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
        refresh_button.config(state=state)
        copy_button.config(state=state)
        word_button.config(state=state)
        variants_button.config(state=state)
    except NameError:
        # Кнопки ещё не созданы (ранний вызов при старте) — это нормально
        pass

    if message:
        try:
            gen_status_label.config(text=message)
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

    _set_busy(True, "⏳ Генерация... (GigaChat)")

    def worker():
        text = ""
        error = None

        # Лок держим ТОЛЬКО вокруг генерации: клиент GigaChat не рассчитан на
        # параллельные запросы, но планировать callback под локом нельзя —
        # finish() вызовет _set_busy(), которому нужен тот же лок (deadlock).
        with _generation_lock:
            try:
                text = generate_text_sync(region, cities, use_cache=use_cache)
            except Exception as e:
                logger.exception("Непредвиденная ошибка генерации: %s", e)
                error = str(e)
                try:
                    text = build_template_text(
                        format_cities_list([c for c in cities if c != CREATIVE_MODE])
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


def _dispatch_to_ui(func) -> None:
    """
    Выполняет func в главном потоке Tk.

    Вызывать виджеты из рабочего потока нельзя, поэтому используем
    window.after(0, ...). Если окно уже закрыто или главный цикл Tk
    недоступен — снимаем флаг занятости, чтобы кнопки не остались мёртвыми.
    """

    def fallback():
        try:
            _set_busy(False, "")
        except Exception:
            pass

    try:
        window.after(0, func)
    except Exception as e:
        logger.debug("Не удалось отправить callback в главный поток Tk: %s", e)
        fallback()


def refresh_preview_async() -> None:
    """Обновляет предпросмотр асинхронно — окно остаётся отзывчивым."""

    def on_done(text, error):
        if error:
            logger.error("Ошибка генерации предпросмотра: %s", error)
        try:
            text_preview.config(state="normal")
            text_preview.delete("1.0", tk.END)
            text_preview.insert("1.0", text or "")
            text_preview.config(state="disabled")
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
                    0, lambda n=i: gen_status_label.config(text=f"⏳ Вариант {n+1}/3...")
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
        save_settings(selected, region, USE_AI, CREATIVITY)

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


def switch_region(event):
    region = get_current_region()
    global CURRENT_REGION
    CURRENT_REGION = region

    if region == "Владивосток":
        cities = ALL_CITIES_VLADIVOSTOK
    elif region == "Восток":
        cities = ALL_CITIES_EAST + [CREATIVE_MODE]
    else:
        cities = ALL_CITIES_SOUTH

    for city, var in city_vars.items():
        var.set(1 if city in cities else 0)

    selected = get_selected_cities()
    if selected:
        global SELECTED_CITIES
        SELECTED_CITIES = selected
        save_settings(selected, region, USE_AI, CREATIVITY)

    refresh_preview_async()


def toggle_ai():
    global USE_AI
    USE_AI = not USE_AI
    ai_button.config(
        text=f"{'✅' if USE_AI else '❌'} ИИ: {'ВКЛ' if USE_AI else 'ВЫКЛ'}",
        bg="#4CAF50" if USE_AI else "#f44336",
    )
    save_settings(get_selected_cities(), get_current_region(), USE_AI, CREATIVITY)
    refresh_preview_async()


def clear_cache():
    AICache.clear()
    messagebox.showinfo("Готово!", "✅ Кэш ИИ очищен!")
    refresh_preview_async()


def update_creativity(val):
    global CREATIVITY
    CREATIVITY = float(val)
    creativity_label.config(text=f"Креативность: {int(CREATIVITY * 100)}%")
    save_settings(get_selected_cities(), get_current_region(), USE_AI, CREATIVITY)
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


def show_ssl_help():
    """
    Инструкция «Как исправить» для SSL-ошибки.

    Текст берётся из llm_provider.get_ssl_help(), чтобы лог и интерфейс
    не расходились. Окно с прокруткой: инструкция длинная.
    """
    dialog = tk.Toplevel(window)
    dialog.title("⚠️ SSL: как исправить")
    dialog.configure(bg=BG_COLOR)
    dialog.transient(window)

    tk.Label(
        dialog,
        text="⚠️ Ошибка проверки SSL-сертификата",
        font=("Arial", 12, "bold"),
        fg="#c62828",
        bg=BG_COLOR,
    ).pack(pady=(15, 5))

    text = tk.Text(
        dialog, width=84, height=20, wrap="word", font=("Consolas", 9), bg="white"
    )
    text.pack(padx=15, pady=5, fill="both", expand=True)
    text.insert("1.0", get_ssl_help())
    text.config(state="disabled")

    def open_download():
        """Открывает страницу с сертификатом в браузере по умолчанию."""
        url = "https://gu-st.ru/content/Other/doc/russian_trusted_root_ca.cer"
        try:
            import webbrowser

            webbrowser.open(url)
        except Exception as e:
            messagebox.showinfo("Ссылка", f"Откройте вручную:\n\n{url}")
            logger.warning("Не удалось открыть браузер: %s", e)

    buttons = tk.Frame(dialog, bg=BG_COLOR)
    buttons.pack(pady=10)

    tk.Button(
        buttons,
        text="⬇️ Открыть страницу сертификата",
        font=("Arial", 9, "bold"),
        bg="#2196F3",
        fg="white",
        padx=12,
        pady=5,
        cursor="hand2",
        command=open_download,
    ).pack(side="left", padx=5)

    tk.Button(
        buttons,
        text="Закрыть",
        font=("Arial", 9),
        bg="#9E9E9E",
        fg="white",
        padx=12,
        pady=5,
        cursor="hand2",
        command=dialog.destroy,
    ).pack(side="left", padx=5)

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
    lines.append("✅ tkinter: OK")

    for module_name, human in (
        ("docx", "python-docx"),
        ("pyperclip", "pyperclip"),
        ("dotenv", "python-dotenv"),
        ("gigachat", "gigachat"),
        ("requests", "requests"),
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

    # ---------- Локальный Ollama ----------
    # Проверяем его только здесь: это сетевой запрос с таймаутом, и при
    # незапущенном Ollama он ждёт несколько секунд. Диагностика и без того
    # идёт в фоновом потоке, а вот при построении окна такая проверка
    # задерживала бы запуск программы.
    try:
        ollama = get_status_info(include_ollama=True)
        if ollama.get("ollama_available"):
            models = ", ".join(ollama.get("ollama_models") or []) or "—"
            lines.append(f"✅ Ollama (локальный ИИ): доступен, модели: {models}")
        else:
            lines.append(
                "ℹ️ Ollama (локальный ИИ): не запущен — для версии с GigaChat это нормально"
            )
        lines.append("")
    except Exception as exc:
        lines.append(f"ℹ️ Ollama: проверить не удалось ({exc})")
        lines.append("")

    if not GIGACHAT_READY:
        lines.append("📊 Итог: программа запустится, но ИИ выключен —")
        lines.append("         ключ GigaChat не задан. Нажмите «🔑 Ключ».")
    elif ca_ok:
        lines.append("📊 Итог: всё на месте — GigaChat должен отвечать.")
    else:
        lines.append("📊 Итог: программа запустится, но GigaChat не будет отвечать")
        lines.append("         до установки сертификата Минцифры (fallback на шаблоны).")

    return "\n".join(lines)


def _diag_show_progress(dialog: "tk.Toplevel", text: "tk.Text") -> None:
    """Заглушка «идёт проверка», чтобы окно не выглядело зависшим."""
    text.config(state="normal")
    text.delete("1.0", "end")
    text.insert("1.0", "🔍 Идёт диагностика, подождите…")
    text.config(state="disabled")


def _show_diagnostics_result(dialog: "tk.Toplevel", report: str) -> None:
    """Показывает готовый отчёт в уже открытом окне."""
    try:
        text = _diag_widgets.get("text")
        if text is None or not text.winfo_exists():
            return
        text.config(state="normal")
        text.delete("1.0", "end")
        text.insert("1.0", report)
        text.config(state="disabled")
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
    dialog = tk.Toplevel(window)
    dialog.title("📋 Диагностика")
    dialog.configure(bg=BG_COLOR)
    dialog.transient(window)

    tk.Label(
        dialog,
        text="📋 Диагностика программы",
        font=("Arial", 12, "bold"),
        bg=BG_COLOR,
    ).pack(pady=(12, 4))

    frame = tk.Frame(dialog, bg=BG_COLOR)
    frame.pack(padx=12, pady=5, fill="both", expand=True)

    text = tk.Text(
        frame, width=92, height=30, wrap="word", font=("Consolas", 9), bg="white"
    )
    scrollbar = tk.Scrollbar(frame, command=text.yview)
    text.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    text.pack(side="left", fill="both", expand=True)

    _diag_widgets["text"] = text
    _diag_show_progress(dialog, text)

    buttons = tk.Frame(dialog, bg=BG_COLOR)
    buttons.pack(pady=10)

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

    tk.Button(
        buttons,
        text="🔄 Обновить",
        font=("Arial", 9),
        bg="#2196F3",
        fg="white",
        padx=12,
        pady=4,
        cursor="hand2",
        command=refresh,
    ).pack(side="left", padx=4)

    tk.Button(
        buttons,
        text="📄 Скопировать отчёт",
        font=("Arial", 9),
        bg="#455A64",
        fg="white",
        padx=12,
        pady=4,
        cursor="hand2",
        command=copy_report,
    ).pack(side="left", padx=4)

    tk.Button(
        buttons,
        text="🌐 Как исправить SSL",
        font=("Arial", 9),
        bg="#c62828",
        fg="white",
        padx=12,
        pady=4,
        cursor="hand2",
        command=show_ssl_help,
    ).pack(side="left", padx=4)

    tk.Button(
        buttons,
        text="Закрыть",
        font=("Arial", 9),
        bg="#9E9E9E",
        fg="white",
        padx=12,
        pady=4,
        cursor="hand2",
        command=dialog.destroy,
    ).pack(side="left", padx=4)

    dialog.bind("<Escape>", lambda event: dialog.destroy())

    _run_diagnostics_async(dialog)
    text.see("1.0")


def _run_diagnostics_async(dialog: "tk.Toplevel") -> None:
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

    dialog = tk.Toplevel(window)
    dialog.title("Ввод API-ключа GigaChat")
    dialog.configure(bg=BG_COLOR)
    dialog.resizable(False, False)
    dialog.transient(window)
    dialog.grab_set()

    tk.Label(
        dialog,
        text="🔑 API-ключ GigaChat",
        font=("Arial", 12, "bold"),
        bg=BG_COLOR,
    ).pack(pady=(15, 5))

    tk.Label(
        dialog,
        text=(
            "Получите Authorization Key (base64):\n"
            "developers.sber.ru → Личный кабинет → Настройки\n\n"
            "Ключ сохранится в .env и НЕ попадёт в Git."
        ),
        font=("Arial", 9),
        bg=BG_COLOR,
        justify="center",
    ).pack(padx=20, pady=5)

    # show="*" — ключ не отображается на экране
    key_entry = tk.Entry(dialog, width=55, font=("Arial", 9), show="*")
    key_entry.pack(padx=20, pady=(10, 2))
    key_entry.focus_set()

    # Счётчик символов: благодаря ему видно, что вставка сработала. Раньше
    # при пустом поле (например, Ctrl+V не сработал) кнопка «Сохранить»
    # отвечала «Ключ не введён» — и было непонятно, почему.
    key_hint = tk.Label(
        dialog,
        text="Длина ключа: 0 символов",
        font=("Arial", 8),
        fg="#888",
        bg=BG_COLOR,
    )
    key_hint.pack()

    def _update_key_hint(event=None) -> None:
        """Показывает, сколько символов реально попало в поле."""
        try:
            length = len(key_entry.get())
        except Exception:
            return
        key_hint.config(
            text=f"Длина ключа: {length} символов"
            + ("" if length else "  ← поле пустое"),
            fg="#888" if length else "#c62828",
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

    tk.Label(
        dialog,
        text="(символы скрыты — это нормально, вставка через Ctrl+V)",
        font=("Arial", 8),
        fg="#888",
        bg=BG_COLOR,
    ).pack()

    def save():
        global GIGACHAT_READY, GIGACHAT_MODEL

        # sanitize_key убирает пробелы, кавычки и невидимые символы, которые
        # попадают в ключ при копировании. Раньше такой ключ сохранялся
        # «как есть», и GigaChat отвечал «неверный ключ» при верном значении.
        key = sanitize_key(key_entry.get())

        if not key:
            key_hint.config(
                text="Длина ключа: 0 символов  ← поле пустое", fg="#c62828"
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

        # Обновляем индикатор в шапке
        try:
            status_label.config(
                text=f"GigaChat: ✅ GigaChat готов ({GIGACHAT_MODEL})",
                fg="#4CAF50",
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

    cancel_button = tk.Button(
        dialog,
        text="Отмена",
        font=("Arial", 9),
        bg="#9E9E9E",
        fg="white",
        padx=15,
        pady=5,
        cursor="hand2",
        command=dialog.destroy,
    )
    cancel_button.pack(side="left", padx=20, pady=15)

    save_button = tk.Button(
        dialog,
        text="💾 Сохранить",
        font=("Arial", 9, "bold"),
        bg="#4CAF50",
        fg="white",
        padx=15,
        pady=5,
        cursor="hand2",
        command=save,
    )
    save_button.pack(side="right", padx=20, pady=15)

    dialog.bind("<Return>", lambda event: save())
    dialog.bind("<Escape>", lambda event: dialog.destroy())

    # Центрируем диалог относительно главного окна
    dialog.update_idletasks()
    x = window.winfo_rootx() + (window.winfo_width() - dialog.winfo_width()) // 2
    y = window.winfo_rooty() + (window.winfo_height() - dialog.winfo_height()) // 3
    dialog.geometry(f"+{max(x, 0)}+{max(y, 0)}")


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
# ГРАФИЧЕСКИЙ ИНТЕРФЕЙС
# ============================================

window = tk.Tk()
window.title("Генератор КП - Автоперевозки (GigaChat)")
window.geometry("880x980")
window.resizable(False, False)
BG_COLOR = "#f0f4f8"
window.configure(bg=BG_COLOR)

header_frame = tk.Frame(window, bg=BG_COLOR)
header_frame.pack(fill="x", pady=10)

tk.Label(
    header_frame,
    text="🚛 ГЕНЕРАТОР КОММЕРЧЕСКИХ ПРЕДЛОЖЕНИЙ",
    font=("Arial", 16, "bold"),
    bg=BG_COLOR,
).pack()

# Индикатор статуса GigaChat
if GIGACHAT_READY:
    status_text = f"✅ GigaChat готов ({GIGACHAT_MODEL})"
    status_color = "#4CAF50"
else:
    status_text = "❌ Ключ не задан — используется шаблонный текст"
    status_color = "#f44336"

status_label = tk.Label(
    header_frame,
    text=f"GigaChat: {status_text}",
    font=("Arial", 9),
    fg=status_color,
    bg=BG_COLOR,
)
status_label.pack()

# Если ключа нет — сразу говорим, что делать (а не просто красная строка).
# Плашка создаётся всегда, а видимостью управляет update_key_banner():
# после ввода ключа кнопкой «🔑 Ключ» она должна исчезать без перезапуска.
key_warning_label = tk.Label(
    header_frame,
    text="⚠️ Ключ GigaChat не задан. Нажмите 🔑 для ввода",
    font=("Arial", 9, "bold"),
    fg="#b71c1c",
    bg="#ffe0e0",
    padx=8,
    pady=3,
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
            key_warning_label.pack_forget()
    else:
        if not key_warning_label.winfo_manager():
            key_warning_label.pack(pady=3)


update_key_banner()

# Плашка про SSL. Показывается в двух случаях:
#   1. СРАЗУ при старте — сертификата Минцифры нет, а ключ задан. Тогда
#      GigaChat заведомо не ответит, и пользователь узнаёт об этом до
#      первого нажатия «Сгенерировать» (раньше — только по факту ошибки).
#   2. После падения запроса по сертификату (has_ssl_error()).
# Плашка немодальная: она ничего не блокирует и не мешает работать.
ssl_banner = tk.Frame(window, bg="#ffcdd2")
ssl_label = tk.Label(
    ssl_banner,
    text="⚠️ SSL: не удалось проверить сертификат GigaChat",
    font=("Arial", 9, "bold"),
    fg="#b71c1c",
    bg="#ffcdd2",
    justify="left",
)
ssl_label.pack(side="left", padx=8, pady=3)
tk.Button(
    ssl_banner,
    text="Как исправить",
    font=("Arial", 8, "bold"),
    bg="#c62828",
    fg="white",
    padx=8,
    pady=1,
    cursor="hand2",
    command=show_ssl_help,
).pack(side="left", padx=4, pady=3)


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
    упакована — из-за этого она не убиралась после успешного запроса.
    Плашку мы только пакуем (pack), поэтому менеджер 'pack' == «показана».

    Вызывается из главного потока: менять виджеты из фонового потока
    в Tkinter нельзя.
    """
    try:
        if cert_ok is None:
            cert_ok = _certificate_present() if GIGACHAT_READY else True

        need_banner = has_ssl_error() or (GIGACHAT_READY and not cert_ok)

        if need_banner:
            if has_ssl_error():
                ssl_label.config(
                    text=(
                        "⚠️ SSL: не удалось проверить сертификат GigaChat.\n"
                        "GigaChat не отвечает — текст КП берётся из шаблонов."
                    )
                )
            else:
                ssl_label.config(
                    text=(
                        "⚠️ SSL-сертификат Минцифры не найден. "
                        "GigaChat может не отвечать. [Как исправить]"
                    )
                )

        shown = bool(ssl_banner.winfo_manager())
        if need_banner:
            if not shown:
                ssl_banner.pack(fill="x", padx=20, pady=4)
        else:
            if shown:
                ssl_banner.pack_forget()
    except Exception as exc:
        logger.debug("Не удалось обновить SSL-плашку: %s", exc)


ai_control_frame = tk.Frame(window, bg=BG_COLOR)
ai_control_frame.pack(pady=5)

ai_button = tk.Button(
    ai_control_frame,
    text=f"{'✅' if USE_AI else '❌'} ИИ: {'ВКЛ' if USE_AI else 'ВЫКЛ'}",
    font=("Arial", 9, "bold"),
    bg="#4CAF50" if USE_AI else "#f44336",
    fg="white",
    padx=10,
    pady=3,
    cursor="hand2",
    command=toggle_ai,
)
ai_button.pack(side="left", padx=5)

tk.Label(
    ai_control_frame, text=f"Модель: {GIGACHAT_MODEL}", font=("Arial", 9), bg=BG_COLOR
).pack(side="left", padx=5)

tk.Button(
    ai_control_frame,
    text="⚡ Прогрев",
    font=("Arial", 9),
    bg="#2196F3",
    fg="white",
    padx=8,
    pady=3,
    cursor="hand2",
    command=warmup_async,
).pack(side="left", padx=2)

tk.Button(
    ai_control_frame,
    text="🔑 Ключ",
    font=("Arial", 9),
    bg="#3F51B5",
    fg="white",
    padx=8,
    pady=3,
    cursor="hand2",
    command=enter_api_key,
).pack(side="left", padx=2)

tk.Button(
    ai_control_frame,
    text="🗑️ Кэш",
    font=("Arial", 9),
    bg="#FF9800",
    fg="white",
    padx=8,
    pady=3,
    cursor="hand2",
    command=clear_cache,
).pack(side="left", padx=5)

tk.Button(
    ai_control_frame,
    text="📋 Диагностика",
    font=("Arial", 9),
    bg="#455A64",
    fg="white",
    padx=8,
    pady=3,
    cursor="hand2",
    command=show_diagnostics,
).pack(side="left", padx=2)

tk.Button(
    ai_control_frame,
    text="❓",
    font=("Arial", 9),
    bg="#607D8B",
    fg="white",
    padx=8,
    pady=3,
    cursor="hand2",
    command=show_gigachat_help,
).pack(side="left", padx=2)

creativity_frame = tk.Frame(window, bg=BG_COLOR)
creativity_frame.pack(pady=5)

creativity_label = tk.Label(
    creativity_frame,
    text=f"Креативность: {int(CREATIVITY * 100)}%",
    font=("Arial", 9),
    bg=BG_COLOR,
)
creativity_label.pack(side="left", padx=10)

creativity_slider = tk.Scale(
    creativity_frame,
    from_=0.2,
    to=1.0,
    resolution=0.05,
    orient="horizontal",
    length=300,
    bg=BG_COLOR,
    command=update_creativity,
)
creativity_slider.set(CREATIVITY)
creativity_slider.pack(side="left", padx=10)

tk.Label(
    creativity_frame,
    text="(0.2 - шаблонно, 1.0 - креативно)",
    font=("Arial", 8),
    fg="#888",
    bg=BG_COLOR,
).pack(side="left", padx=5)

notebook = ttk.Notebook(window)
notebook.pack(fill="x", padx=20, pady=5)

tab_south = tk.Frame(notebook, bg=BG_COLOR)
notebook.add(tab_south, text="📍 Юга")

tab_vladivostok = tk.Frame(notebook, bg=BG_COLOR)
notebook.add(tab_vladivostok, text="📍 Владивосток")

tab_east = tk.Frame(notebook, bg=BG_COLOR)
notebook.add(tab_east, text="📍 Восток")

notebook.bind("<<NotebookTabChanged>>", switch_region)

city_vars = {}

# ===== ЮГА =====
tk.Label(
    tab_south, text="📍 Выберите города:", font=("Arial", 10, "bold"), bg=BG_COLOR
).grid(row=0, column=0, columnspan=3, pady=5)
row = 1
for i, city in enumerate(ALL_CITIES_SOUTH):
    var = IntVar(value=1 if city in SELECTED_CITIES and CURRENT_REGION == "Юга" else 0)
    city_vars[city] = var
    Checkbutton(
        tab_south, text=city, variable=var, font=("Arial", 10), bg=BG_COLOR, cursor="hand2"
    ).grid(row=row, column=i % 3, padx=15, pady=3, sticky="w")
    if i % 3 == 2:
        row += 1

# ===== ВЛАДИВОСТОК =====
tk.Label(
    tab_vladivostok, text="📍 Выберите направления:", font=("Arial", 10, "bold"), bg=BG_COLOR
).grid(row=0, column=0, columnspan=3, pady=5)
row_vl = 1
for i, city in enumerate(ALL_CITIES_VLADIVOSTOK):
    var = IntVar(
        value=1 if city in SELECTED_CITIES and CURRENT_REGION == "Владивосток" else 0
    )
    city_vars[city] = var
    Checkbutton(
        tab_vladivostok,
        text=city,
        variable=var,
        font=("Arial", 10),
        bg=BG_COLOR,
        cursor="hand2",
    ).grid(row=row_vl, column=i % 3, padx=15, pady=3, sticky="w")
    if i % 3 == 2:
        row_vl += 1

# ===== ВОСТОК =====
tk.Label(
    tab_east, text="📍 Выберите города:", font=("Arial", 10, "bold"), bg=BG_COLOR
).grid(row=0, column=0, columnspan=4, pady=5)

var_creative = IntVar(
    value=1 if CREATIVE_MODE in SELECTED_CITIES and CURRENT_REGION == "Восток" else 0
)
city_vars[CREATIVE_MODE] = var_creative
Checkbutton(
    tab_east,
    text=CREATIVE_MODE,
    variable=var_creative,
    font=("Arial", 9, "italic"),
    bg=BG_COLOR,
    fg="#FF5722",
    cursor="hand2",
).grid(row=1, column=0, columnspan=4, padx=10, pady=5, sticky="w")

row_east = 2
for i, city in enumerate(ALL_CITIES_EAST):
    var = IntVar(value=1 if city in SELECTED_CITIES and CURRENT_REGION == "Восток" else 0)
    city_vars[city] = var
    Checkbutton(
        tab_east, text=city, variable=var, font=("Arial", 9), bg=BG_COLOR, cursor="hand2"
    ).grid(row=row_east, column=i % 4, padx=10, pady=2, sticky="w")
    if i % 4 == 3:
        row_east += 1

tk.Button(
    window,
    text="✅ ПРИМЕНИТЬ ВЫБОР ГОРОДОВ",
    font=("Arial", 9, "bold"),
    bg="#4CAF50",
    fg="white",
    padx=20,
    pady=5,
    cursor="hand2",
    command=update_cities,
).pack(pady=5)

tk.Frame(window, height=2, bg="#ccc").pack(fill="x", padx=20, pady=5)

refresh_button = tk.Button(
    window,
    text="🔄 ПОКАЗАТЬ НОВЫЙ ВАРИАНТ",
    font=("Arial", 10, "bold"),
    bg="#9C27B0",
    fg="white",
    padx=20,
    pady=8,
    cursor="hand2",
    command=refresh_preview_async,
)
refresh_button.pack(pady=5)

# Статус фоновой генерации
gen_status_label = tk.Label(window, text="", font=("Arial", 9), fg="#FF9800", bg=BG_COLOR)
gen_status_label.pack(pady=2)

text_preview = tk.Text(
    window, height=13, font=("Arial", 10), wrap="word", relief="solid", bd=1, bg="white"
)
text_preview.pack(fill="both", padx=20, pady=10, expand=True)
text_preview.config(state="disabled")

tk.Label(
    window,
    text="💡 Для Востока: отметьте 'Без города' для креативного режима",
    font=("Arial", 9),
    bg=BG_COLOR,
    fg="#888",
).pack(pady=2)

button_frame = tk.Frame(window, bg=BG_COLOR)
button_frame.pack(pady=5)

copy_button = tk.Button(
    button_frame,
    text="📋 КОПИРОВАТЬ",
    font=("Arial", 11, "bold"),
    bg="#4CAF50",
    fg="white",
    padx=20,
    pady=10,
    width=18,
    cursor="hand2",
    command=copy_to_clipboard,
)
copy_button.grid(row=0, column=0, padx=5, pady=5)

word_button = tk.Button(
    button_frame,
    text="📄 СОХРАНИТЬ WORD",
    font=("Arial", 11, "bold"),
    bg="#2196F3",
    fg="white",
    padx=20,
    pady=10,
    width=18,
    cursor="hand2",
    command=save_to_word,
)
word_button.grid(row=0, column=1, padx=5, pady=5)

variants_button = tk.Button(
    button_frame,
    text="📚 3 ВАРИАНТА",
    font=("Arial", 11, "bold"),
    bg="#FF9800",
    fg="white",
    padx=20,
    pady=10,
    width=18,
    cursor="hand2",
    command=save_multiple_variants,
)
variants_button.grid(row=0, column=2, padx=5, pady=5)

tk.Button(
    window,
    text="✖ ВЫХОД",
    font=("Arial", 10),
    bg="#f44336",
    fg="white",
    padx=15,
    pady=5,
    width=15,
    cursor="hand2",
    command=on_exit_button,
).pack(pady=10)

logger.info("Интерфейс загружен")

# Плашку SSL выставляем сразу после построения окна: если сертификата нет,
# пользователь видит предупреждение ещё до первого запроса к GigaChat.
# Повторная (уже без записи в лог) проверка — в _run_gui перед mainloop().
update_ssl_banner()


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
