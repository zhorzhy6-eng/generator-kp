import pyperclip
import tkinter as tk
from tkinter import messagebox, Checkbutton, IntVar, ttk
from datetime import datetime
import os
import random
import json
import logging

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Пути проекта и настройка логирования (общий модуль с версией для Ollama)
from logger_config import setup_logger, SETTINGS_PATH, LEGACY_LOG_FILE

logging.getLogger("Генератор_КП").setLevel(logging.DEBUG)
logger = setup_logger("Генератор_КП", level=logging.DEBUG)

logger.info("=" * 60)
logger.info("ЗАПУСК ПРОГРАММЫ ГЕНЕРАТОР КП")
logger.info("=" * 60)

# ============================================
# ФАЙЛ ДЛЯ СОХРАНЕНИЯ НАСТРОЕК
# ============================================

SETTINGS_FILE = SETTINGS_PATH
LOG_FILE = LEGACY_LOG_FILE

def load_settings():
    """Загружает настройки из JSON файла"""
    try:
        with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            logger.info(f"Настройки загружены: города={len(data.get('cities', []))}, регион={data.get('region', 'не указан')}")
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

def save_settings(selected_cities, region):
    """Сохраняет настройки в JSON файл"""
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump({"cities": selected_cities, "region": region}, f, ensure_ascii=False, indent=2)
        logger.info(f"Настройки сохранены: города={len(selected_cities)}, регион={region}")
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

saved = load_settings()
if saved:
    SELECTED_CITIES = saved.get("cities", ALL_CITIES_SOUTH.copy())
    CURRENT_REGION = saved.get("region", "Юга")
    logger.info(f"Загружены настройки пользователя: регион={CURRENT_REGION}, города={SELECTED_CITIES}")
else:
    SELECTED_CITIES = ALL_CITIES_SOUTH.copy()
    CURRENT_REGION = "Юга"
    logger.info("Использую настройки по умолчанию: регион=Юга")

# ============================================
# ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ДЛЯ GUI
# ============================================

city_vars = {}
notebook = None
text_preview = None

# ============================================
# ФУНКЦИИ
# ============================================

def get_selected_cities():
    """Возвращает список выбранных городов"""
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
    """Определяет текущий активный регион по вкладке"""
    try:
        tab = notebook.index(notebook.select())
        if tab == 0:
            return "Юга"
        elif tab == 1:
            return "Владивосток"
        else:
            return "Восток"
    except:
        return CURRENT_REGION  # Возвращаем сохранённый регион

def switch_region(event):
    """Переключает регион при смене вкладки"""
    region = get_current_region()
    global CURRENT_REGION, SELECTED_CITIES
    CURRENT_REGION = region
    logger.info(f"Переключение на регион: {region}")
    
    if region == "Владивосток":
        cities = ALL_CITIES_VLADIVOSTOK
    elif region == "Восток":
        cities = ALL_CITIES_EAST + [CREATIVE_MODE]
    else:
        cities = ALL_CITIES_SOUTH
    
    # Обновляем чекбоксы
    for city, var in city_vars.items():
        var.set(1 if city in cities else 0)
    
    selected = get_selected_cities()
    if selected:
        SELECTED_CITIES = selected
        save_settings(selected, region)
    show_preview()

# ============================================
# ОФОРМЛЕНИЕ ГОРОДОВ - РАСШИРЕННЫЙ КРЕАТИВ
# ============================================

def format_route():
    """Форматирует маршрут с креативным оформлением"""
    cities = get_selected_cities()
    region = get_current_region()
    
    is_creative = CREATIVE_MODE in cities and len(cities) == 1
    
    if is_creative and region == "Восток":
        creative_phrases = [
            "🚛 Москва → Восточные регионы — новые горизонты открываются!",
            "📦 Москва → На Восток — держим курс на приключения!",
            "🚚 Москва → Восточная Сибирь и Дальний Восток — масштабный охват!",
            "📊 Москва → Восточное направление — стратегический маршрут!",
            "🚛 Москва → Урал, Сибирь, Дальний Восток — охватываем всё!",
            "📦 Москва → Восточные города России — от столицы до океана!",
            "🚚 Москва → От Урала до Тихого океана — трансконтинентальный маршрут!",
            "📊 Москва → Восточные регионы РФ — надёжная логистика!",
            "🌏 Москва → Восток — соединяем столицу с Дальним Востоком!",
            "⚡ Москва → Восточные рубежи — доставка в любую точку!"
        ]
        return random.choice(creative_phrases)
    
    if not cities:
        return "Нет выбранных городов"
    
    cities = [c for c in cities if c != CREATIVE_MODE]
    if not cities:
        if region == "Восток":
            creative_phrases = [
                "🚛 Москва → Восточные регионы — открываем новые возможности!",
                "📦 Москва → На Восток — надёжный логистический коридор!",
                "🚚 Москва → Восточная Сибирь и Дальний Восток — полный спектр!",
                "📊 Москва → Восточное направление — ваш надёжный партнёр!",
                "🚛 Москва → Урал, Сибирь, Дальний Восток — максимальный охват!",
                "📦 Москва → Восточные города России — доставим в любую точку!",
                "🚚 Москва → От Урала до Тихого океана — полный маршрут!",
                "📊 Москва → Восточные регионы РФ — опыт и надёжность!"
            ]
            return random.choice(creative_phrases)
        return "Нет выбранных городов"
    
    if len(cities) == 1:
        city = cities[0]
        one_city_styles = [
            f"🎯 Маршрут на {city} — точное попадание!",
            f"🚀 Направление: {city} — старт дан!",
            f"📍 Маршрут следования: {city} — главная цель!",
            f"⭐ Пункт назначения: {city} — ждём вас!",
            f"🏁 Маршрут до {city} — финишная прямая!",
            f"🎯 Направление на {city} — в прицеле!",
            f"📍 {city} — точка на карте!",
            f"🚀 Маршрут: {city} — движение вперёд!",
            f"⭐ Точка назначения: {city} — наш ориентир!",
            f"🏁 Курс: {city} — держим путь!"
        ]
        return random.choice(one_city_styles)
    
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
    return random.choice(styles)()

# ============================================
# КРЕАТИВНЫЕ ФРАЗЫ - РАСШИРЕННЫЕ
# ============================================

INTROS_SOUTH = [
    "🔥 Отличные лоты с Москвы на Юга — успейте забрать!",
    "🚛 Свежие лоты из Москвы на Южное направление — спешите!",
    "💥 Топовые предложения с Москвы на Юга — лучший выбор!",
    "⭐ Новые лоты с Москвы на популярные южные маршруты!",
    "🚀 Горячие предложения с Москвы на Юга — не упустите!",
    "📦 Отличные варианты с Москвы до Черноморского побережья!",
    "🏆 Лучшие предложения с Москвы на юг России — премиум-класс!",
    "⚡ Срочные лоты с Москвы на курорты Юга — загрузка сегодня!",
    "🎯 Точечные предложения с Москвы на Южные регионы!",
    "🌟 Супер-лоты с Москвы на Юга — только свежие!",
    "🚚 Свежие машины из Москвы в южные города — встречайте!",
    "💎 Эксклюзивные предложения с Москвы на Юга — VIP-лоты!",
    "📢 Новое поступление лотов с Москвы на Юга — успейте!",
    "🌊 Южное направление — свежие лоты из Москвы!",
    "☀️ Солнечные маршруты: Москва → Юга — лучшие предложения!",
]

CARGO_SOUTH = [
    "🚗 Груз — кроссоверы, лоты сформированы, машины в наличии, загрузка быстрая.",
    "🚙 Предлагаем кроссоверы — машины готовы, загрузка оперативная.",
    "🚘 В наличии кроссоверы — быстрая загрузка, лоты полностью готовы.",
    "🚗 Кроссоверы в наличии, лоты готовы к отправке, загрузка без задержек.",
    "🚙 Свежие кроссоверы — все лоты в наличии, загрузка моментальная.",
    "🚘 Кроссоверы ждут загрузки — лоты полностью укомплектованы.",
    "🚗 Отличные кроссоверы — машины на месте, загрузка по графику.",
    "🚙 Автомобили премиум-класса — ждут своего часа!",
    "🚘 Только свежие кроссоверы — идеальное состояние!",
]

PAYMENT_SOUTH = [
    "💳 Оплата — безнал, сроки индивидуальные, рассматриваем предоплату.",
    "💰 Способы оплаты: безналичный расчёт, гибкие сроки, возможна предоплата.",
    "💵 Работаем по безналу, сроки обсуждаем, рассматриваем частичную предоплату.",
    "🏦 Оплата безналичная, условия индивидуальные, предоплата приветствуется.",
    "💳 Безналичный расчёт — сроки подбираем под клиента, возможна предоплата.",
    "💰 Оплата без наличных — индивидуальные сроки, предоплата обсуждаема.",
    "💵 Гибкая система оплаты — подстроимся под ваши условия!",
]

LEGAL_SOUTH = [
    "📋 Работаем по договору, платим без задержек.",
    "📄 Заключаем договор, выплаты строго в срок, без просрочек.",
    "🤝 Официальное оформление, оплата вовремя, гарантируем надёжность.",
    "📑 Работаем официально, платежи без задержек, договор гарантирует защиту.",
    "📋 Все сделки по договору, выплаты чётко в срок, без задержек.",
    "📄 Юридически чистая работа — договор, оплата по факту, без просрочек.",
    "⚖️ Полное юридическое сопровождение — ваша безопасность в приоритете!",
]

CTA_SOUTH = [
    "✍️ Все детали и ставки — в ЛС. Пишите! 📩",
    "📩 Ждём ваших сообщений для обсуждения деталей и ставок!",
    "💬 Пишите в личные сообщения — обсудим все условия и ставки!",
    "📲 Свяжитесь с нами для уточнения деталей и получения ставок!",
    "✍️ Детали и стоимость — в личку. Ждём ваших вопросов!",
    "📩 Обсудим все нюансы и ставки в личных сообщениях!",
    "📱 Ждём ваших сообщений — готовы ответить на все вопросы!",
]

INTROS_VLADIVOSTOK = [
    "🔥 Ловите горячие лоты с Владивостока — по всей России!",
    "🚛 Свежие лоты из Владивостока по всей России — успейте!",
    "💥 Топовые предложения с Владивостока на все направления!",
    "⭐ Новые лоты с Владивостока — успейте забрать лучшие!",
    "🚀 Горячие предложения с Дальнего Востока — только сегодня!",
    "📦 Отличные варианты из Владивостока в любую точку страны!",
    "🏆 Лучшие лоты с Владивостока на рынке — эксклюзив!",
    "⚡ Срочные лоты с Владивостока — загрузка сегодня же!",
    "🎯 Точечные предложения с Владивостока для вашего бизнеса!",
    "🌟 Супер-лоты с Владивостока по выгодным ценам!",
    "🚚 Свежие машины из Владивостока в города России — встречайте!",
    "💎 Эксклюзивные предложения с Владивостока — VIP-сервис!",
    "📢 Новое поступление лотов с Владивостока — спешите!",
    "🌊 Владивосток — точка силы. Лучшие лоты для вас!",
    "⚓ Владивосток — начало великого маршрута!",
]

CARGO_VLADIVOSTOK = [
    "✅ Все лоты готовы к отгрузке в любой момент.",
    "🚗 Машины в наличии, загрузка быстрая и чёткая.",
    "🚙 Лоты полностью готовы к отправке в любой момент.",
    "🚘 Все авто в наличии, документы в порядке, загрузка сразу.",
    "🚗 Свежие лоты — машины на месте, готовы к отгрузке.",
    "🚙 Полная готовность — ждём только вашего сигнала!",
]

PAYMENT_VLADIVOSTOK = [
    "💰 Оплата — 50/50 наличными при погрузке.",
    "💳 Оплата — 50% при загрузке, 50% по факту доставки.",
    "💵 Работаем по схеме 50/50 — предоплата при погрузке.",
    "🏦 Оплата наличными — 50% при погрузке, 50% после.",
    "💸 Гибкая система оплаты — обсудим индивидуально!",
]

LEGAL_VLADIVOSTOK = [
    "📦 Работаем чисто, без задержек. Подробности по каждому лоту — в личные сообщения.",
    "📋 Работаем по договору, все сделки прозрачны.",
    "📄 Гарантируем чистоту сделок, работаем без задержек.",
    "🤝 Надёжные партнёры — оплата вовремя, без задержек.",
    "📑 Полная юридическая прозрачность — ваше спокойствие гарантируем!",
]

CTA_VLADIVOSTOK = [
    "✍️ Пишите — обсудим ставки, даты и условия!",
    "📩 Обсудим детали в личных сообщениях — ждём ваших вопросов!",
    "💬 Свяжитесь с нами для уточнения ставок и дат!",
    "📲 Пишите — договоримся о всех условиях!",
    "📱 Ждём вашего сообщения — готовы к диалогу!",
]

INTROS_EAST = [
    "🚛 Предложения с Москвы на Восток — охват максимальный!",
    "📦 Лоты с Москвы в восточные регионы — спешите!",
    "🚚 Надёжные поставки с Москвы на Восток — проверенный маршрут!",
    "📊 Оптимальные маршруты Москва → Восток — логистика высшего уровня!",
    "🚛 Транспортные решения Москва — Восточная Сибирь — полный сервис!",
    "📦 Грузоперевозки Москва → Урал → Сибирь → Дальний Восток!",
    "🚚 Логистика Москва — Восточные регионы РФ — надёжно и в срок!",
    "📊 Москва → Восток: надёжно и в срок — ваш выбор!",
    "🚛 Доставка грузов из Москвы в восточные города — быстро и качественно!",
    "📦 Коммерческие предложения по маршруту Москва — Восток — лучшие условия!",
    "🌏 Москва → Восток — трансконтинентальный маршрут!",
    "⚡ Москва → Восток — логистика нового поколения!",
]

CARGO_EAST = [
    "🚗 Лоты сформированы, машины в наличии — ждут отправки.",
    "🚙 Предлагаем авто — полная готовность к отправке.",
    "🚘 В наличии автомобили — загрузка по графику.",
    "🚗 Машины готовы к отгрузке в восточном направлении.",
    "🚙 Полный парк готов к отправке — выбирайте!",
]

PAYMENT_EAST = [
    "💳 Оплата — безнал, сроки обсуждаем индивидуально.",
    "💰 Способы оплаты: безналичный расчёт, гибкие условия.",
    "💵 Работаем по безналу — быстро и прозрачно.",
    "🏦 Оплата безналичная, условия под каждого клиента.",
    "💸 Гибкие условия оплаты — подстроимся под вас!",
]

LEGAL_EAST = [
    "📋 Работаем по договору, выплаты в срок.",
    "📄 Заключаем договор — гарантия надёжности.",
    "🤝 Официальное оформление, оплата вовремя.",
    "📑 Все сделки по договору, без задержек.",
    "⚖️ Полное юридическое сопровождение сделки!",
]

CTA_EAST = [
    "✍️ Детали и ставки — в ЛС. Ждём ваших сообщений!",
    "📩 Обсудим условия и ставки в личных сообщениях.",
    "💬 Пишите — рассчитаем оптимальные условия!",
    "📲 Свяжитесь с нами для уточнения деталей!",
    "📱 Ждём вашего звонка или сообщения!",
]

# ============================================
# ФУНКЦИЯ ГЕНЕРАЦИИ
# ============================================

def generate_creative_text():
    """Генерирует креативный текст коммерческого предложения"""
    region = get_current_region()
    route = format_route()
    logger.debug(f"Генерация текста для региона: {region}")
    
    if region == "Владивосток":
        intro = random.choice(INTROS_VLADIVOSTOK)
        cargo = random.choice(CARGO_VLADIVOSTOK)
        payment = random.choice(PAYMENT_VLADIVOSTOK)
        legal = random.choice(LEGAL_VLADIVOSTOK)
        cta = random.choice(CTA_VLADIVOSTOK)
    elif region == "Восток":
        intro = random.choice(INTROS_EAST)
        cargo = random.choice(CARGO_EAST)
        payment = random.choice(PAYMENT_EAST)
        legal = random.choice(LEGAL_EAST)
        cta = random.choice(CTA_EAST)
    else:
        intro = random.choice(INTROS_SOUTH)
        cargo = random.choice(CARGO_SOUTH)
        payment = random.choice(PAYMENT_SOUTH)
        legal = random.choice(LEGAL_SOUTH)
        cta = random.choice(CTA_SOUTH)
    
    result = f"{intro}\n\n{route}\n\n{cargo}\n{payment}\n{legal}\n\n{cta}"
    logger.debug("Текст успешно сгенерирован")
    return result

# ============================================
# КНОПКИ
# ============================================

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
        logger.info(f"Word-файл сохранён: {filepath}")
        messagebox.showinfo("Готово!", f"✅ Word-файл создан на рабочем столе!\n\n{filename}")
    except Exception as e:
        logger.error(f"Ошибка сохранения в Word: {e}")
        messagebox.showerror("Ошибка", f"❌ {str(e)}")

def save_multiple_variants():
    """Создаёт 3 разных варианта в Word"""
    logger.info("Создание 3 вариантов в Word")
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
            logger.debug(f"Вариант {i+1} сохранён: {filepath}")
        messagebox.showinfo("Готово!", "✅ 3 файла созданы на рабочем столе!")
    except Exception as e:
        logger.error(f"Ошибка создания 3 вариантов: {e}")
        messagebox.showerror("Ошибка", f"❌ {str(e)}")

def show_preview():
    """Обновляет предпросмотр текста"""
    logger.debug("Обновление предпросмотра")
    try:
        text = generate_creative_text()
        text_preview.config(state="normal")
        text_preview.delete("1.0", tk.END)
        text_preview.insert("1.0", text)
        text_preview.config(state="disabled")
    except Exception as e:
        logger.error(f"Ошибка обновления предпросмотра: {e}")

def create_installer():
    """Создаёт установщик с помощью Inno Setup"""
    logger.info("Запуск создания установщика")
    try:
        import subprocess
        import shutil
        current_dir = os.path.dirname(os.path.abspath(__file__))
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        messagebox.showinfo("Установщик", "🔨 Сборка...\n\nПодождите 2-3 минуты.")
        subprocess.run(['pyinstaller', '--onefile', '--console', '--name', 'Генератор_КП', os.path.join(current_dir, 'генератор_кп.py')], check=True)
        exe_source = os.path.join(current_dir, 'dist', 'Генератор_КП.exe')
        setup_folder = os.path.join(desktop, 'Setup')
        if not os.path.exists(setup_folder):
            os.makedirs(setup_folder)
        shutil.copy(exe_source, os.path.join(setup_folder, 'Генератор_КП.exe'))
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
        with open(os.path.join(setup_folder, 'setup.iss'), 'w', encoding='utf-8') as f:
            f.write(iss_content)
        os.startfile(setup_folder)
        logger.info("Папка Setup создана на рабочем столе")
        messagebox.showinfo("Готово!", "✅ Папка Setup создана на рабочем столе!\n\nОткройте setup.iss в Inno Setup и нажмите F9")
    except Exception as e:
        logger.error(f"Ошибка создания установщика: {e}")
        messagebox.showerror("Ошибка", f"❌ {str(e)}\n\nУстановите PyInstaller:\npip install pyinstaller")

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
# ГРАФИЧЕСКИЙ ИНТЕРФЕЙС
# ============================================

logger.info("Создание графического интерфейса")

window = tk.Tk()
window.title("Генератор КП - Автоперевозки")
window.geometry("750x850")
window.resizable(False, False)
BG_COLOR = "#f0f4f8"
window.configure(bg=BG_COLOR)

# Обработчик закрытия окна
window.protocol("WM_DELETE_WINDOW", on_closing)

tk.Label(window, text="🚛 ГЕНЕРАТОР КОММЕРЧЕСКИХ ПРЕДЛОЖЕНИЙ", font=("Arial", 16, "bold"), bg=BG_COLOR, pady=10).pack()

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
tk.Label(tab_south, text="📍 Выберите города для маршрута:", font=("Arial", 10, "bold"), bg=BG_COLOR).grid(row=0, column=0, columnspan=3, pady=5)
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
tk.Label(tab_east, text="📍 Выберите города для маршрута (в порядке следования):", font=("Arial", 10, "bold"), bg=BG_COLOR).grid(row=0, column=0, columnspan=4, pady=5)

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

# Устанавливаем правильную вкладку при запуске
if CURRENT_REGION == "Юга":
    notebook.select(0)
elif CURRENT_REGION == "Владивосток":
    notebook.select(1)
else:
    notebook.select(2)

tk.Button(window, text="✅ ПРИМЕНИТЬ ВЫБОР ГОРОДОВ", font=("Arial", 9, "bold"), bg="#4CAF50", fg="white", padx=20, pady=5, cursor="hand2", command=update_cities).pack(pady=5)
tk.Frame(window, height=2, bg="#ccc").pack(fill="x", padx=20, pady=5)

tk.Button(window, text="🔄 ПОКАЗАТЬ НОВЫЙ ВАРИАНТ", font=("Arial", 10, "bold"), bg="#9C27B0", fg="white", padx=20, pady=8, cursor="hand2", command=show_preview).pack(pady=5)

text_preview = tk.Text(window, height=12, font=("Arial", 10), wrap="word", relief="solid", bd=1, bg="white")
text_preview.pack(fill="both", padx=20, pady=10, expand=True)
show_preview()

tk.Label(window, text="📌 Для Востока: отметьте 'Без города' для креативного режима без перечисления городов", font=("Arial", 9), bg=BG_COLOR, fg="#888").pack(pady=2)

button_frame = tk.Frame(window, bg=BG_COLOR)
button_frame.pack(pady=5)

tk.Button(button_frame, text="📋 КОПИРОВАТЬ В БУФЕР", font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", padx=20, pady=10, width=20, cursor="hand2", command=copy_to_clipboard).grid(row=0, column=0, padx=5, pady=5)
tk.Button(button_frame, text="📄 СОХРАНИТЬ В WORD", font=("Arial", 11, "bold"), bg="#2196F3", fg="white", padx=20, pady=10, width=20, cursor="hand2", command=save_to_word).grid(row=0, column=1, padx=5, pady=5)

tk.Button(window, text="📚 СОЗДАТЬ 3 РАЗНЫХ ВАРИАНТА В WORD", font=("Arial", 10, "bold"), bg="#FF9800", fg="white", padx=20, pady=10, width=45, cursor="hand2", command=save_multiple_variants).pack(pady=5)

tk.Button(window, text="🛠 СОЗДАТЬ УСТАНОВЩИК (Inno Setup)", font=("Arial", 10, "bold"), bg="#FF5722", fg="white", padx=20, pady=10, width=45, cursor="hand2", command=create_installer).pack(pady=5)

tk.Button(window, text="✖ ВЫХОД", font=("Arial", 10), bg="#f44336", fg="white", padx=15, pady=5, width=15, cursor="hand2", command=on_closing).pack(pady=10)

logger.info("Графический интерфейс загружен. Запуск mainloop()")
window.mainloop()