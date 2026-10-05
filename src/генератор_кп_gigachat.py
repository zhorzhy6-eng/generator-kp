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
import threading
import time
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Импортируем общий модуль: пути проекта и логгер
from logger_config import setup_logger, SETTINGS_PATH

# Единый провайдер LLM (GigaChat + Ollama)
from llm_provider import (
    DEFAULT_SYSTEM_PROMPT,
    generate_gigachat,
    get_status_info,
    is_gigachat_configured,
    warmup_gigachat,
)

# Настраиваем логгер
logger = setup_logger(__name__)
logger.info("=" * 70)
logger.info("ЗАПУСК ГЕНЕРАТОРА КП С GIGACHAT")
logger.info("=" * 70)

# ============================================
# КОНФИГУРАЦИЯ ГЕНЕРАЦИИ (параметры под скорость)
# ============================================

AI_TEMPERATURE = 0.85      # креативность
AI_MAX_TOKENS = 280        # короче ответ — быстрее приходит
AI_TIMEOUT = 30            # секунд, задаётся также в .env (GIGACHAT_TIMEOUT)

# Короткий системный промпт: меньше токенов на вход — быстрее ответ
GIGACHAT_SYSTEM_PROMPT = (
    DEFAULT_SYSTEM_PROMPT + " Отвечай строго 3 предложениями и коротким призывом к действию."
)

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
    """Блокирует/разблокирует кнопки и переключает статус на время генерации."""
    global _generation_in_progress
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

        # Один сетевой запрос за раз: клиент GigaChat не рассчитан на параллельность
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
            on_done(text, error)

        # Возвращаемся в главный поток Tk
        try:
            window.after(0, finish)
        except Exception:
            # Окно уже закрыто
            pass

    threading.Thread(target=worker, name="gigachat-generate", daemon=True).start()
    return True


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


def _save_docx(text: str, filename: str) -> str:
    """Создаёт Word-файл с текстом КП на рабочем столе."""
    doc = Document()
    title = doc.add_heading("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    date_paragraph = doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y')}")
    date_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph()

    for line in text.split("\n"):
        if line.strip():
            doc.add_paragraph(line)

    filepath = os.path.join(os.path.expanduser("~"), "Desktop", filename)
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
            _save_docx(text, filename)
            messagebox.showinfo(
                "Готово!", f"✅ Word-файл создан на рабочем столе!\n\n{filename}"
            )
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
                _save_docx(text, filename)
                saved_files.append(filename)
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
            messagebox.showinfo(
                "Готово!", f"✅ {len(saved_files)} файла созданы на рабочем столе!"
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
        "Ключ хранится ТОЛЬКО в файле .env в корне проекта и никогда не попадает в Git.\n\n"
        "🌐 Если в логах ошибка SSL (certificate verify failed), укажите в .env\n"
        "путь к корневому сертификату Минцифры:\n"
        "GIGACHAT_CA_BUNDLE_FILE=C:\\certs\\russian_trusted_root_ca.cer\n\n"
        "Скачать: https://gu-st.ru/content/Other/doc/russian_trusted_root_ca.cer",
    )


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

tk.Label(
    header_frame,
    text=f"GigaChat: {status_text}",
    font=("Arial", 9),
    fg=status_color,
    bg=BG_COLOR,
).pack()

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
    command=window.quit,
).pack(pady=10)

logger.info("Интерфейс загружен")

# Первый предпросмотр и прогрев авторизации — в фоне, окно открывается сразу
window.after(100, refresh_preview_async)

if GIGACHAT_READY:
    threading.Thread(target=warmup_gigachat, name="gigachat-startup-warmup", daemon=True).start()

window.mainloop()
logger.info("Программа завершена")
