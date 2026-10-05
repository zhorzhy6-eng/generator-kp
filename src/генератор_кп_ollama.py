#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Генератор коммерческих предложений с ИИ (Ollama)
Версия: 4.2 - Уникальный текст при каждой генерации (исправленная)
"""

import pyperclip
import tkinter as tk
from tkinter import messagebox, Checkbutton, IntVar, ttk
from datetime import datetime
import os
import random
import json
import re
import time
from typing import List, Dict, Optional, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Импортируем общий модуль: пути проекта и логгер
from logger_config import setup_logger, SETTINGS_PATH

# Настраиваем логгер
logger = setup_logger(__name__)
logger.info("=" * 70)
logger.info("ЗАПУСК ГЕНЕРАТОРА КП С OLLAMA (v4.2 - УНИКАЛЬНЫЕ ТЕКСТЫ)")
logger.info("=" * 70)

# ============================================
# ПРОВЕРКА НАЛИЧИЯ OLLAMA
# ============================================

def check_ollama_available() -> Tuple[bool, str]:
    """Проверяет, доступен ли Ollama и возвращает первую доступную модель"""
    try:
        import requests
        response = requests.get("http://localhost:11434/api/tags", timeout=2)
        if response.status_code == 200:
            data = response.json()
            models = data.get('models', [])
            if models:
                model_name = models[0]['name']
                logger.info(f"✅ Ollama доступен. Используем модель: {model_name}")
                return True, model_name
            else:
                logger.warning("⚠️ Ollama доступен, но нет установленных моделей")
                return False, ""
        return False, ""
    except Exception as e:
        logger.warning(f"⚠️ Ollama не доступен: {e}")
        return False, ""

def get_all_models() -> List[str]:
    """Возвращает список всех доступных моделей Ollama"""
    try:
        import requests
        response = requests.get("http://localhost:11434/api/tags", timeout=2)
        if response.status_code == 200:
            data = response.json()
            return [m['name'] for m in data.get('models', [])]
        return []
    except:
        return []

# Проверяем доступность Ollama
OLLAMA_AVAILABLE, DEFAULT_MODEL = check_ollama_available()
logger.info(f"Ollama статус: {OLLAMA_AVAILABLE}, модель по умолчанию: {DEFAULT_MODEL}")

# ============================================
# НАСТРОЙКИ
# ============================================

SETTINGS_FILE = SETTINGS_PATH

def load_settings():
    """Загружает настройки из файла"""
    try:
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return None

def save_settings(selected_cities, region, use_ai=True, model_name=None, creativity=0.8):
    """Сохраняет настройки в файл"""
    try:
        if model_name is None:
            model_name = DEFAULT_MODEL if OLLAMA_AVAILABLE else "llama3.2"
        
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump({
                "cities": selected_cities,
                "region": region,
                "use_ai": use_ai,
                "model": model_name,
                "creativity": creativity
            }, f, ensure_ascii=False, indent=2)
        logger.info(f"Настройки сохранены: регион={region}, городов={len(selected_cities)}")
    except Exception as e:
        logger.error(f"Ошибка сохранения настроек: {e}")

# Загружаем сохранённые настройки
saved = load_settings()
if saved:
    SELECTED_CITIES = saved.get("cities", [])
    CURRENT_REGION = saved.get("region", "Юга")
    USE_AI = saved.get("use_ai", True)
    AI_MODEL = saved.get("model", DEFAULT_MODEL if OLLAMA_AVAILABLE else "llama3.2")
    CREATIVITY = saved.get("creativity", 0.8)
else:
    SELECTED_CITIES = []
    CURRENT_REGION = "Юга"
    USE_AI = True
    AI_MODEL = DEFAULT_MODEL if OLLAMA_AVAILABLE else "llama3.2"
    CREATIVITY = 0.8

logger.info(f"Загружена модель: {AI_MODEL}, креативность: {CREATIVITY}")

# ============================================
# СПИСКИ ГОРОДОВ
# ============================================

ALL_CITIES_SOUTH = [
    "Ростов-на-Дону",
    "Краснодар",
    "Новороссийск",
    "Адыгея",
    "Пятигорск",
    "Минеральные Воды"
]

ALL_CITIES_VLADIVOSTOK = [
    "Владивосток → Москва",
    "Владивосток → Новосибирск",
    "Владивосток → Краснодар",
    "Владивосток → Екатеринбург",
    "Владивосток → Казань",
    "Владивосток → Санкт-Петербург"
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
    "Чебоксары"
]

CREATIVE_MODE = "✨ Без города (креативный режим)"

# ============================================
# ВАРИАНТЫ СТИЛЕЙ ДЛЯ УНИКАЛЬНОСТИ
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
# ИИ-ГЕНЕРАЦИЯ
# ============================================

class AIGuard:
    """Класс для защиты от галлюцинаций ИИ"""
    
    @staticmethod
    def validate_ai_response(text: str) -> Tuple[bool, str]:
        """Валидация ответа ИИ"""
        if not text or len(text.strip()) < 20:
            return False, "Текст слишком короткий"
        
        # Проверка на ключевые слова
        required_keywords = ['машин', 'ло', 'груз', 'авто', 'транспорт']
        has_keywords = any(kw in text.lower() for kw in required_keywords)
        if not has_keywords:
            return False, "Нет ключевых слов по теме"
        
        # Проверка что мы не предлагаем "ваши грузы"
        wrong_phrases = ['ваши грузы', 'ваш груз', 'вашего груза', 'погрузим ваши']
        has_wrong = any(phrase in text.lower() for phrase in wrong_phrases)
        if has_wrong:
            return False, "Неправильная формулировка"
        
        if len(text) > 2000:
            return False, "Текст слишком длинный"
        
        return True, "OK"
    
    @staticmethod
    def fix_ai_response(text: str) -> str:
        """Исправляет типичные проблемы в ответе ИИ"""
        text = re.sub(r'\n{3,}', '\n\n', text)
        text = re.sub(r' {2,}', ' ', text)
        
        # Удаляем подписи
        text = re.sub(r'С уважением,?\s*\[.*?\]', '', text, flags=re.IGNORECASE)
        text = re.sub(r'\(название компании\)', '', text, flags=re.IGNORECASE)
        text = re.sub(r'\[Название компании\]', '', text, flags=re.IGNORECASE)
        text = re.sub(r'С уважением,?\s*$', '', text, flags=re.IGNORECASE)
        
        return text.strip()

# ============================================
# УМНЫЙ КЭШ (с учётом времени и случайности)
# ============================================

class AICache:
    """Кэш с учётом времени жизни"""
    _cache: Dict[str, Tuple[str, float]] = {}
    
    @classmethod
    def get(cls, key: str) -> Optional[str]:
        if key in cls._cache:
            value, timestamp = cls._cache[key]
            # Кэш живёт 30 секунд - чтобы каждый раз был новый текст
            if time.time() - timestamp < 30:
                return value
            else:
                del cls._cache[key]
        return None
    
    @classmethod
    def set(cls, key: str, value: str):
        cls._cache[key] = (value, time.time())
    
    @classmethod
    def clear(cls):
        cls._cache.clear()
        logger.info("Кэш очищен")

def generate_with_ollama(prompt: str, model: str = None, timeout: int = 20, creativity: float = 0.8) -> Optional[str]:
    """
    Генерация текста через Ollama.
    """
    if not OLLAMA_AVAILABLE:
        return None
    
    if model is None:
        model = AI_MODEL
    
    available_models = get_all_models()
    if model not in available_models:
        if available_models:
            model = available_models[0]
            logger.info(f"Автоматически выбрана модель: {model}")
        else:
            return None
    
    try:
        import requests
        
        temperature = 0.6 + (creativity * 0.4)  # от 0.6 до 1.0
        
        payload = {
            "model": model,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": temperature,
                "top_p": 0.9,
                "top_k": 50,
                "num_predict": 300,
                "repeat_penalty": 1.2,
                "presence_penalty": 0.2,
                "frequency_penalty": 0.2,
            }
        }
        
        start_time = time.time()
        response = requests.post(
            "http://localhost:11434/api/generate",
            json=payload,
            timeout=timeout
        )
        
        elapsed = time.time() - start_time
        logger.info(f"Запрос к Ollama выполнен за {elapsed:.2f} сек, статус: {response.status_code}")
        
        if response.status_code == 200:
            result = response.json()
            generated_text = result.get('response', '').strip()
            
            is_valid, reason = AIGuard.validate_ai_response(generated_text)
            if is_valid:
                logger.info(f"✅ ИИ сгенерировал текст (длина: {len(generated_text)} символов)")
                return AIGuard.fix_ai_response(generated_text)
            else:
                logger.warning(f"⚠️ Ответ ИИ не прошёл валидацию: {reason}")
                return None
        else:
            logger.error(f"Ошибка API Ollama: {response.status_code}")
            return None
            
    except requests.Timeout:
        logger.error(f"Таймаут запроса к Ollama ({timeout} сек)")
        return None
    except Exception as e:
        logger.error(f"Ошибка при запросе к Ollama: {e}")
        return None

# ============================================
# ФУНКЦИИ ГЕНЕРАЦИИ (УНИКАЛЬНЫЙ ТЕКСТ)
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

def generate_ai_text(region: str, cities: List[str]) -> Optional[str]:
    """
    Генерирует уникальный текст с помощью ИИ.
    """
    if not USE_AI or not OLLAMA_AVAILABLE:
        return None
    
    # УНИКАЛЬНЫЙ КЛЮЧ КЭША - добавляем случайное число и время
    city_str = ",".join(sorted(cities))
    random_seed = random.randint(1, 1000)
    cache_key = f"{region}_{city_str}_{AI_MODEL}_{int(CREATIVITY*10)}_{random_seed}"
    
    # Проверяем кэш (он живёт 30 секунд)
    cached = AICache.get(cache_key)
    if cached:
        logger.info("✅ Использован кэшированный ответ ИИ")
        return cached
    
    # Формируем города
    city_names = format_cities_list(cities) if cities else "все города"
    
    # Описание региона
    region_desc = {
        "Юга": "южные регионы (Краснодар, Ростов-на-Дону)",
        "Владивосток": "все направления из Владивостока",
        "Восток": "восточные регионы (Урал, Сибирь, Дальний Восток)"
    }.get(region, "регионы России")
    
    # РАЗНЫЕ ВАРИАНТЫ ПРОМПТОВ для уникальности
    prompt_variants = [
        f"""Напиши живое коммерческое предложение для перевозок автотранспортом.

Маршрут: Москва -> {region_desc}.
Города: {city_names}.

Стиль: энергичный, как в чате перевозчиков.
Напиши про наличие машин, условия оплаты, договор.
Объём: 3-4 предложения.
В конце призыв к действию.

Пример: "Лоты сформированы! Машины в наличии, готовы к загрузке. Работаем по безналу, договор сразу. Ждём ваших ставок!""",
        
        f"""Напиши коммерческое предложение для грузоперевозок.

Направление: Москва -> {region_desc}.
Города: {city_names}.

Требования: деловой стиль, но не официальный.
Упомяни: машины в наличии, оплата безналом, договор.
4-5 предложений.

Пример: "Машины готовы к отправке! Лоты сформированы, загрузка быстрая. Оплата безналичная, работаем по договору. Пишите!""",
        
        f"""Напиши краткое коммерческое предложение.

Маршрут: Москва -> {region_desc}.
Города: {city_names}.

Коротко и по делу: машины есть, условия прозрачные, работаем честно.
3 предложения максимум.

Пример: "Лоты в наличии! Машины готовы, загрузка сегодня. Безнал, договор. Пишите!""",
    ]
    
    prompt = random.choice(prompt_variants)
    
    logger.info(f"Генерация ИИ текста для региона {region}, креативность: {CREATIVITY}")
    
    # Генерируем текст
    generated = generate_with_ollama(prompt, AI_MODEL, creativity=CREATIVITY)
    
    if generated:
        AICache.set(cache_key, generated)
        return generated
    
    return None

def generate_creative_text() -> str:
    """
    Генерация уникального текста при каждом вызове.
    """
    region = get_current_region()
    cities = get_selected_cities()
    
    # Убираем креативный режим
    route_cities = [c for c in cities if c != CREATIVE_MODE]
    cities_text = format_cities_list(route_cities) if route_cities else ""
    
    logger.debug(f"Генерация текста для региона {region}, городов: {len(route_cities)}")
    
    # Пробуем ИИ
    ai_text = generate_ai_text(region, route_cities)
    
    if ai_text:
        logger.info("✅ Использован ИИ-текст")
        if cities_text:
            return f"{ai_text}\n\n{cities_text}"
        return ai_text
    
    # Fallback на РАНДОМНЫЕ шаблоны
    logger.info("🔄 Использован шаблонный текст (fallback)")
    
    # Собираем случайные фразы
    opener = random.choice(OPENERS)
    status = random.choice(STATUSES)
    payment = random.choice(PAYMENTS)
    legal = random.choice(LEGALS)
    cta = random.choice(CTA)
    
    # Собираем в текст с разными вариациями
    variants = [
        f"{opener} {status} {payment} {legal} {cta}",
        f"{opener} {payment} {status} {legal} {cta}",
        f"{opener} {status} {legal} {payment} {cta}",
    ]
    
    result = random.choice(variants)
    
    # Добавляем города
    if cities_text:
        result = f"{result}\n\n{cities_text}"
    
    return result

# ============================================
# ИНТЕРФЕЙС
# ============================================

def copy_to_clipboard():
    try:
        text = generate_creative_text()
        pyperclip.copy(text)
        messagebox.showinfo("Готово!", "✅ Текст скопирован в буфер обмена!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"❌ {str(e)}")

def save_to_word():
    try:
        text = generate_creative_text()
        doc = Document()
        title = doc.add_heading('КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ', level=1)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        date_paragraph = doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y')}")
        date_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        doc.add_paragraph()
        
        for line in text.split('\n'):
            if line.strip():
                doc.add_paragraph(line)
        
        region = get_current_region()
        filename = f"КП_{region}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        filepath = os.path.join(os.path.expanduser('~'), 'Desktop', filename)
        doc.save(filepath)
        
        messagebox.showinfo("Готово!", f"✅ Word-файл создан на рабочем столе!\n\n{filename}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"❌ {str(e)}")

def save_multiple_variants():
    try:
        region = get_current_region()
        for i in range(3):
            text = generate_creative_text()
            doc = Document()
            title = doc.add_heading('КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ', level=1)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            date_paragraph = doc.add_paragraph(f"Дата: {datetime.now().strftime('%d.%m.%Y')}")
            date_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            doc.add_paragraph()
            
            for line in text.split('\n'):
                if line.strip():
                    doc.add_paragraph(line)
            
            filename = f"КП_{region}_вариант_{i+1}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            filepath = os.path.join(os.path.expanduser('~'), 'Desktop', filename)
            doc.save(filepath)
        
        messagebox.showinfo("Готово!", "✅ 3 файла созданы на рабочем столе!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"❌ {str(e)}")

def show_preview():
    try:
        text = generate_creative_text()
        text_preview.config(state="normal")
        text_preview.delete("1.0", tk.END)
        text_preview.insert("1.0", text)
        text_preview.config(state="disabled")
    except Exception as e:
        logger.error(f"Ошибка предпросмотра: {e}")

def update_cities():
    selected = get_selected_cities()
    region = get_current_region()
    
    if selected:
        global SELECTED_CITIES
        SELECTED_CITIES = selected
        save_settings(selected, region, USE_AI, AI_MODEL, CREATIVITY)
        show_preview()
        
        if len(selected) == 1:
            if selected[0] == CREATIVE_MODE:
                msg = "✅ Выбран креативный режим 'Без города'"
            else:
                msg = f"✅ Выбран город:\n\n{selected[0]}"
        else:
            msg = f"✅ Выбрано: {len(selected)}\n\n{', '.join(selected)}"
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
        save_settings(selected, region, USE_AI, AI_MODEL, CREATIVITY)
    
    show_preview()

def toggle_ai():
    global USE_AI
    USE_AI = not USE_AI
    ai_button.config(
        text=f"{'✅' if USE_AI else '❌'} ИИ: {'ВКЛ' if USE_AI else 'ВЫКЛ'}",
        bg="#4CAF50" if USE_AI else "#f44336"
    )
    save_settings(get_selected_cities(), get_current_region(), USE_AI, AI_MODEL, CREATIVITY)
    show_preview()

def clear_cache():
    AICache.clear()
    messagebox.showinfo("Готово!", "✅ Кэш ИИ очищен!")
    show_preview()

def refresh_models():
    models = get_all_models()
    model_combo['values'] = models
    if models:
        current = model_combo.get()
        if current not in models:
            model_combo.set(models[0])
            change_model()
    else:
        model_combo.set('Нет доступных моделей')
    messagebox.showinfo("Готово!", f"✅ Найдено моделей: {len(models)}")

def change_model():
    global AI_MODEL
    selected = model_combo.get()
    if selected and selected != 'Нет доступных моделей':
        AI_MODEL = selected
        save_settings(get_selected_cities(), get_current_region(), USE_AI, AI_MODEL, CREATIVITY)
        AICache.clear()
        show_preview()
        messagebox.showinfo("Готово!", f"✅ Модель изменена на: {AI_MODEL}")

def update_creativity(val):
    global CREATIVITY
    CREATIVITY = float(val)
    creativity_label.config(text=f"Креативность: {int(CREATIVITY * 100)}%")
    save_settings(get_selected_cities(), get_current_region(), USE_AI, AI_MODEL, CREATIVITY)
    AICache.clear()

# ============================================
# ГРАФИЧЕСКИЙ ИНТЕРФЕЙС
# ============================================

window = tk.Tk()
window.title("Генератор КП - Автоперевозки (с ИИ)")
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
    bg=BG_COLOR
).pack()

status_text = f"✅ ИИ доступен ({DEFAULT_MODEL})" if OLLAMA_AVAILABLE else "❌ ИИ не доступен"
status_color = "#4CAF50" if OLLAMA_AVAILABLE else "#f44336"
tk.Label(
    header_frame,
    text=f"Ollama: {status_text}",
    font=("Arial", 9),
    fg=status_color,
    bg=BG_COLOR
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
    command=toggle_ai
)
ai_button.pack(side="left", padx=5)

tk.Label(ai_control_frame, text="Модель:", font=("Arial", 9), bg=BG_COLOR).pack(side="left", padx=5)

available_models = get_all_models()
model_combo = ttk.Combobox(ai_control_frame, values=available_models, width=20)
if available_models:
    model_combo.set(AI_MODEL if AI_MODEL in available_models else available_models[0])
else:
    model_combo.set('Нет доступных моделей')
model_combo.pack(side="left", padx=5)
model_combo.bind('<<ComboboxSelected>>', lambda e: change_model())

tk.Button(
    ai_control_frame,
    text="🔄",
    font=("Arial", 10),
    bg="#2196F3",
    fg="white",
    padx=8,
    pady=3,
    cursor="hand2",
    command=refresh_models
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
    command=clear_cache
).pack(side="left", padx=5)

creativity_frame = tk.Frame(window, bg=BG_COLOR)
creativity_frame.pack(pady=5)

creativity_label = tk.Label(creativity_frame, text=f"Креативность: {int(CREATIVITY * 100)}%", font=("Arial", 9), bg=BG_COLOR)
creativity_label.pack(side="left", padx=10)

creativity_slider = tk.Scale(
    creativity_frame,
    from_=0.2, to=1.0,
    resolution=0.05,
    orient="horizontal",
    length=300,
    bg=BG_COLOR,
    command=update_creativity
)
creativity_slider.set(CREATIVITY)
creativity_slider.pack(side="left", padx=10)

tk.Label(
    creativity_frame,
    text="(0.2 - шаблонно, 1.0 - креативно)",
    font=("Arial", 8),
    fg="#888",
    bg=BG_COLOR
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
tk.Label(tab_south, text="📍 Выберите города:", font=("Arial", 10, "bold"), bg=BG_COLOR).grid(row=0, column=0, columnspan=3, pady=5)
row = 1
for i, city in enumerate(ALL_CITIES_SOUTH):
    var = IntVar(value=1 if city in SELECTED_CITIES and CURRENT_REGION == "Юга" else 0)
    city_vars[city] = var
    Checkbutton(tab_south, text=city, variable=var, font=("Arial", 10), bg=BG_COLOR, cursor="hand2").grid(row=row, column=i % 3, padx=15, pady=3, sticky="w")
    if i % 3 == 2:
        row += 1

# ===== ВЛАДИВОСТОК =====
tk.Label(tab_vladivostok, text="📍 Выберите направления:", font=("Arial", 10, "bold"), bg=BG_COLOR).grid(row=0, column=0, columnspan=3, pady=5)
row_vl = 1
for i, city in enumerate(ALL_CITIES_VLADIVOSTOK):
    var = IntVar(value=1 if city in SELECTED_CITIES and CURRENT_REGION == "Владивосток" else 0)
    city_vars[city] = var
    Checkbutton(tab_vladivostok, text=city, variable=var, font=("Arial", 10), bg=BG_COLOR, cursor="hand2").grid(row=row_vl, column=i % 3, padx=15, pady=3, sticky="w")
    if i % 3 == 2:
        row_vl += 1

# ===== ВОСТОК =====
tk.Label(tab_east, text="📍 Выберите города:", font=("Arial", 10, "bold"), bg=BG_COLOR).grid(row=0, column=0, columnspan=4, pady=5)

var_creative = IntVar(value=1 if CREATIVE_MODE in SELECTED_CITIES and CURRENT_REGION == "Восток" else 0)
city_vars[CREATIVE_MODE] = var_creative
Checkbutton(tab_east, text=CREATIVE_MODE, variable=var_creative, font=("Arial", 9, "italic"), bg=BG_COLOR, fg="#FF5722", cursor="hand2").grid(row=1, column=0, columnspan=4, padx=10, pady=5, sticky="w")

row_east = 2
for i, city in enumerate(ALL_CITIES_EAST):
    var = IntVar(value=1 if city in SELECTED_CITIES and CURRENT_REGION == "Восток" else 0)
    city_vars[city] = var
    Checkbutton(tab_east, text=city, variable=var, font=("Arial", 9), bg=BG_COLOR, cursor="hand2").grid(row=row_east, column=i % 4, padx=10, pady=2, sticky="w")
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
    command=update_cities
).pack(pady=5)

tk.Frame(window, height=2, bg="#ccc").pack(fill="x", padx=20, pady=5)

tk.Button(
    window,
    text="🔄 ПОКАЗАТЬ НОВЫЙ ВАРИАНТ",
    font=("Arial", 10, "bold"),
    bg="#9C27B0",
    fg="white",
    padx=20,
    pady=8,
    cursor="hand2",
    command=show_preview
).pack(pady=5)

text_preview = tk.Text(window, height=13, font=("Arial", 10), wrap="word", relief="solid", bd=1, bg="white")
text_preview.pack(fill="both", padx=20, pady=10, expand=True)

window.after(100, show_preview)

tk.Label(
    window,
    text="💡 Для Востока: отметьте 'Без города' для креативного режима",
    font=("Arial", 9),
    bg=BG_COLOR,
    fg="#888"
).pack(pady=2)

button_frame = tk.Frame(window, bg=BG_COLOR)
button_frame.pack(pady=5)

tk.Button(
    button_frame,
    text="📋 КОПИРОВАТЬ",
    font=("Arial", 11, "bold"),
    bg="#4CAF50",
    fg="white",
    padx=20,
    pady=10,
    width=18,
    cursor="hand2",
    command=copy_to_clipboard
).grid(row=0, column=0, padx=5, pady=5)

tk.Button(
    button_frame,
    text="📄 СОХРАНИТЬ WORD",
    font=("Arial", 11, "bold"),
    bg="#2196F3",
    fg="white",
    padx=20,
    pady=10,
    width=18,
    cursor="hand2",
    command=save_to_word
).grid(row=0, column=1, padx=5, pady=5)

tk.Button(
    button_frame,
    text="📚 3 ВАРИАНТА",
    font=("Arial", 11, "bold"),
    bg="#FF9800",
    fg="white",
    padx=20,
    pady=10,
    width=18,
    cursor="hand2",
    command=save_multiple_variants
).grid(row=0, column=2, padx=5, pady=5)

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
    command=window.quit
).pack(pady=10)

logger.info("Интерфейс загружен")
window.mainloop()
logger.info("Программа завершена")