#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Универсальный провайдер LLM для «Генератора КП».

Поддерживает два бэкенда:
  * GigaChat (облачный, Сбер) — ключ берётся ТОЛЬКО из .env
  * Ollama   (локальный)      — http://localhost:11434

Принципы безопасности:
  * API-ключ никогда не логируется: любые упоминания маскируются как "***"
  * Ключ не хранится в коде — только в .env (который в .gitignore)
  * SSL-верификация включена по умолчанию
  * Ни одна функция не выбрасывает исключение наружу: при ошибке возвращается None

Принципы скорости:
  * Клиент GigaChat создаётся лениво и кэшируется (singleton) — токен и
    TLS-соединение переиспользуются между запросами
  * Генерация идёт потоково (stream=True): текст начинает приходить сразу
  * Есть отдельная функция warmup() для прогрева авторизации в фоне
"""

import os
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

# ============================================
# ЗАГРУЗКА .env (безопасно, если dotenv не установлен)
# ============================================
# .env ищем в корне проекта относительно этого файла (<корень>/src/llm_provider.py)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"

try:
    from dotenv import load_dotenv

    # override=False: реальные переменные окружения имеют приоритет над .env
    load_dotenv(ENV_PATH, override=False)
except ImportError:  # pragma: no cover — python-dotenv не установлен
    pass

# ============================================
# ЛОГГЕР
# ============================================
# Используем общий логгер проекта. Если импорт не удался (например, модуль
# запущен вне структуры проекта) — падаем на стандартный logging, чтобы
# провайдер оставался полностью автономным.

try:
    from logger_config import setup_logger

    logger = setup_logger(__name__)
except Exception:  # pragma: no cover
    import logging

    logger = logging.getLogger(__name__)

# ============================================
# КОНСТАНТЫ
# ============================================

DEFAULT_GIGACHAT_SCOPE = "GIGACHAT_API_PERS"
DEFAULT_GIGACHAT_MODEL = "GigaChat"
DEFAULT_GIGACHAT_TIMEOUT = 30.0
DEFAULT_OLLAMA_HOST = "http://localhost:11434"
DEFAULT_OLLAMA_MODEL = "llama3.2"
DEFAULT_OLLAMA_TIMEOUT = 20

# Значения, которые считаем «ключ не задан»
_PLACEHOLDER_VALUES = {
    "",
    "your_key_here",
    "your_key",
    "changeme",
    "none",
    "null",
    "вставьте_ключ",
    "тут_ключ",
}

# Имена исключений gigachat, означающие проблему с авторизацией/токеном.
# Только они приводят к пересозданию клиента; сетевые и SSL-ошибки — нет.
_AUTH_ERROR_NAMES = {
    "ResponseError",
    "GigaChatException",
    "AuthenticationError",
    "AuthorizationError",
}

# ============================================
# МАСКИРОВАНИЕ СЕКРЕТОВ
# ============================================


def mask_secret(value: Optional[str], keep: int = 4) -> str:
    """
    Маскирует секрет для логов.

    Возвращает "***" для пустого или короткого значения, иначе "abcd...wxyz".

    ВНИМАНИЕ: по умолчанию этот вариант НЕ используется для API-ключа —
    см. mask_credentials(). Функция оставлена для отладочных случаев,
    когда нужно сверить, ТОТ ли ключ загружен.
    """
    if not value:
        return "***"
    value = str(value)
    if len(value) <= keep * 2:
        return "***"
    return f"{value[:keep]}...{value[-keep:]}"


# Политика маскирования ключа GigaChat.
# None — показывать только "***" (значение по умолчанию, безопасный режим).
# Целое число — показывать первые/последние N символов (отладка).
# Меняйте только осознанно: любой фрагмент ключа в логе — это утечка.
GIGACHAT_MASK_KEEP: Optional[int] = None


def mask_credentials(credentials: Optional[str]) -> str:
    """
    Маска для Authorization Key GigaChat.

    По умолчанию возвращает ровно "***" — ключ не попадает в лог ни одним
    символом. Возвращает "***" и для пустого значения, чтобы по логу нельзя
    было отличить «ключ не задан» от «ключ задан».
    """
    if GIGACHAT_MASK_KEEP is None:
        return "***"
    return mask_secret(credentials, keep=GIGACHAT_MASK_KEEP)


def scrub_text(text: str) -> str:
    """
    Вычищает секреты из произвольного текста (например, текста исключения).

    Нужна, потому что некоторые библиотеки и HTTP-клиенты любят включать
    заголовки/URL с токеном в сообщение об ошибке.
    """
    if not text:
        return ""
    scrubbed = str(text)

    # Прямые значения секретов из окружения
    for name in ("GIGACHAT_CREDENTIALS", "GIGACHAT_PASSWORD"):
        secret = os.environ.get(name)
        if secret and len(secret) >= 6:
            scrubbed = scrubbed.replace(secret, "***")

    # Общие шаблоны: Bearer <token>, Authorization: Basic <base64>, access_token=<...>
    scrubbed = re.sub(r"(?i)(bearer\s+)[A-Za-z0-9._\-+/=]{8,}", r"\1***", scrubbed)
    scrubbed = re.sub(r"(?i)(basic\s+)[A-Za-z0-9+/=]{8,}", r"\1***", scrubbed)
    scrubbed = re.sub(
        r"(?i)(access_token[\"'\s:=]+)[A-Za-z0-9._\-+/=]{8,}", r"\1***", scrubbed
    )
    scrubbed = re.sub(
        r"(?i)(authorization[\"'\s:=]+)[A-Za-z0-9._\-+/=]{8,}", r"\1***", scrubbed
    )

    return scrubbed


# ============================================
# GIGACHAT — КОНФИГУРАЦИЯ
# ============================================


def get_gigachat_credentials() -> str:
    """Возвращает Authorization Key из окружения (после загрузки .env)."""
    return (os.environ.get("GIGACHAT_CREDENTIALS") or "").strip()


def is_gigachat_configured() -> bool:
    """
    Проверяет, задан ли ключ GigaChat.

    Сама библиотека gigachat при этом НЕ требуется — проверка чисто по .env,
    поэтому функция безопасна для вызова до установки зависимостей.
    """
    return get_gigachat_credentials().lower() not in _PLACEHOLDER_VALUES


# ============================================
# GIGACHAT — ЛЕНИВЫЙ SINGLETON КЛИЕНТ
# ============================================

_gigachat_client: Optional[Any] = None
_gigachat_config_key: Optional[tuple] = None


def _current_config_key() -> tuple:
    """Отпечаток конфигурации: при её смене клиент пересоздаётся."""
    return (
        get_gigachat_credentials(),
        os.environ.get("GIGACHAT_SCOPE", DEFAULT_GIGACHAT_SCOPE),
        os.environ.get("GIGACHAT_MODEL", DEFAULT_GIGACHAT_MODEL),
        os.environ.get("GIGACHAT_VERIFY_SSL_CERTS", "true"),
        os.environ.get("GIGACHAT_CA_BUNDLE_FILE", ""),
    )


def reset_gigachat_client() -> None:
    """
    Сбрасывает кэшированный клиент (например, после смены ключа без перезапуска
    программы). Старое соединение закрывается, ошибки игнорируются.
    """
    global _gigachat_client, _gigachat_config_key

    if _gigachat_client is not None:
        try:
            _gigachat_client.close()
        except Exception:
            pass

    _gigachat_client = None
    _gigachat_config_key = None
    logger.info("Клиент GigaChat сброшен")


def _get_gigachat_client() -> Optional[Any]:
    """
    Возвращает клиент GigaChat, создавая его лениво при первом вызове.

    Ленивость важна: отсутствие ключа или библиотеки не должно ломать
    программу на старте — GUI должен открыться и показать статус.
    """
    global _gigachat_client, _gigachat_config_key

    if not is_gigachat_configured():
        logger.warning("GigaChat не настроен: GIGACHAT_CREDENTIALS не задан в .env")
        return None

    config_key = _current_config_key()

    # Клиент уже создан и конфигурация не менялась — переиспользуем
    if _gigachat_client is not None and _gigachat_config_key == config_key:
        return _gigachat_client

    # Конфигурация изменилась — пересоздаём
    if _gigachat_client is not None:
        logger.info("Конфигурация GigaChat изменилась — пересоздаю клиент")
        reset_gigachat_client()

    try:
        from gigachat import GigaChat
    except ImportError:
        logger.error(
            "Библиотека gigachat не установлена. "
            "Установите: pip install gigachat python-dotenv"
        )
        return None

    credentials = get_gigachat_credentials()
    scope = os.environ.get("GIGACHAT_SCOPE", DEFAULT_GIGACHAT_SCOPE).strip()
    model = os.environ.get("GIGACHAT_MODEL", DEFAULT_GIGACHAT_MODEL).strip()

    try:
        timeout = float(os.environ.get("GIGACHAT_TIMEOUT", DEFAULT_GIGACHAT_TIMEOUT))
    except (TypeError, ValueError):
        logger.warning("Некорректный GIGACHAT_TIMEOUT — использую значение по умолчанию")
        timeout = DEFAULT_GIGACHAT_TIMEOUT

    # SSL-верификация включена, если явно не отключена ("false")
    verify_ssl = (
        os.environ.get("GIGACHAT_VERIFY_SSL_CERTS", "true").strip().lower() != "false"
    )
    ca_bundle = (os.environ.get("GIGACHAT_CA_BUNDLE_FILE") or "").strip() or None

    try:
        _gigachat_client = GigaChat(
            credentials=credentials,
            scope=scope,
            model=model,
            timeout=timeout,
            verify_ssl_certs=verify_ssl,
            ca_bundle_file=ca_bundle,
        )
        _gigachat_config_key = config_key

        # ВАЖНО: сам ключ не логируем — только маску
        logger.info(
            "Клиент GigaChat создан | credentials=%s | scope=%s | model=%s | "
            "timeout=%sс | verify_ssl=%s | ca_bundle=%s",
            mask_credentials(credentials),
            scope,
            model,
            timeout,
            verify_ssl,
            ca_bundle or "системный",
        )
        return _gigachat_client

    except Exception as e:
        logger.error("Не удалось создать клиент GigaChat: %s", scrub_text(str(e)))
        _gigachat_client = None
        _gigachat_config_key = None
        return None


def warmup_gigachat() -> bool:
    """
    Прогревает авторизацию GigaChat (получает токен заранее).

    Вызывать в фоновом потоке на старте GUI — тогда первый реальный запрос
    не будет ждать OAuth-хендшейк. Возвращает True при успехе.
    """
    client = _get_gigachat_client()
    if client is None:
        return False

    try:
        start = time.time()
        client.get_token()
        logger.info("Прогрев GigaChat выполнен за %.2f сек", time.time() - start)
        return True
    except Exception as e:
        logger.warning("Прогрев GigaChat не удался: %s", scrub_text(str(e)))
        return False


# ============================================
# GIGACHAT — ГЕНЕРАЦИЯ
# ============================================


def _extract_message_content(response: Any) -> str:
    """Достаёт текст ответа из любого поддерживаемого формата ответа."""
    if response is None:
        return ""

    # Основной путь: response.choices[0].message.content
    choices = getattr(response, "choices", None)
    if choices:
        message = getattr(choices[0], "message", None)
        content = getattr(message, "content", None)
        if content:
            return str(content).strip()
        # На случай, если пришёл "сырой" dict
        if isinstance(choices[0], dict):
            return str(choices[0].get("message", {}).get("content", "")).strip()

    # Фолбэк для dict-подобных ответов
    if isinstance(response, dict):
        try:
            return str(response["choices"][0]["message"]["content"]).strip()
        except (KeyError, IndexError, TypeError):
            return ""

    return ""


def _generate_gigachat_streaming(
    client: Any, system_prompt: str, prompt: str, temperature: float, max_tokens: int
) -> str:
    """
    Потоковая генерация: текст приходит по мере готовности.

    Это быстрее по ощущениям (и экономит время на разбор полного ответа),
    поэтому используется как основной путь.
    """
    from gigachat.models import Chat, Messages, MessagesRole

    payload = Chat(
        messages=[
            Messages(role=MessagesRole.SYSTEM, content=system_prompt),
            Messages(role=MessagesRole.USER, content=prompt),
        ],
        temperature=temperature,
        max_tokens=max_tokens,
        stream=True,
    )

    chunks: List[str] = []
    for chunk in client.stream(payload):
        content = _extract_message_content(chunk)
        if content:
            chunks.append(content)

    return "".join(chunks).strip()


def _generate_gigachat_sync(
    client: Any, system_prompt: str, prompt: str, temperature: float, max_tokens: int
) -> str:
    """Обычная (непотоковая) генерация — фолбэк, если стриминг не сработал."""
    from gigachat.models import Chat, Messages, MessagesRole

    payload = Chat(
        messages=[
            Messages(role=MessagesRole.SYSTEM, content=system_prompt),
            Messages(role=MessagesRole.USER, content=prompt),
        ],
        temperature=temperature,
        max_tokens=max_tokens,
    )

    return _extract_message_content(client.chat(payload))


def generate_gigachat(
    prompt: str,
    system_prompt: Optional[str] = None,
    temperature: float = 0.85,
    max_tokens: int = 280,
    use_stream: bool = True,
) -> Optional[str]:
    """
    Генерирует текст через GigaChat.

    Args:
        prompt:        пользовательский запрос
        system_prompt: системная роль (по умолчанию — копирайтер по автоперевозкам)
        temperature:   креативность (0.0–2.0)
        max_tokens:    ограничение длины ответа (меньше = быстрее)
        use_stream:    потоковая генерация (быстрее отдаёт первый токен)

    Returns:
        Текст ответа или None при любой ошибке. Исключения не выбрасываются.
    """
    if not prompt or not str(prompt).strip():
        logger.warning("generate_gigachat: пустой prompt")
        return None

    if system_prompt is None:
        system_prompt = DEFAULT_SYSTEM_PROMPT

    if not is_gigachat_configured():
        logger.warning("generate_gigachat: ключ GigaChat не задан (см. .env)")
        return None

    client = _get_gigachat_client()
    if client is None:
        return None

    try:
        temperature = float(temperature)
    except (TypeError, ValueError):
        temperature = 0.85

    try:
        max_tokens = int(max_tokens)
    except (TypeError, ValueError):
        max_tokens = 280

    start = time.time()

    try:
        text = ""
        if use_stream:
            try:
                text = _generate_gigachat_streaming(
                    client, system_prompt, str(prompt), temperature, max_tokens
                )
            except Exception as stream_error:
                # Стриминг может не поддерживаться прокси/сетью — идём обычным путём
                logger.warning(
                    "Потоковая генерация GigaChat не удалась (%s), "
                    "переключаюсь на обычный режим",
                    scrub_text(str(stream_error)),
                )
                text = _generate_gigachat_sync(
                    client, system_prompt, str(prompt), temperature, max_tokens
                )
        else:
            text = _generate_gigachat_sync(
                client, system_prompt, str(prompt), temperature, max_tokens
            )

        elapsed = time.time() - start

        if not text:
            logger.warning("GigaChat вернул пустой ответ за %.2f сек", elapsed)
            return None

        logger.info(
            "GigaChat: ответ получен за %.2f сек (символов: %d)", elapsed, len(text)
        )
        return text

    except Exception as e:
        elapsed = time.time() - start
        error_name = type(e).__name__
        logger.error(
            "Ошибка GigaChat за %.2f сек: %s: %s",
            elapsed,
            error_name,
            scrub_text(str(e)),
        )

        # Сбрасываем клиент ТОЛЬКО при ошибке авторизации (протух токен,
        # неверный ключ). Сетевые/SSL-ошибки клиент не портят — пересоздавать
        # его нельзя, иначе теряется прогрев и TLS-соединение.
        if error_name in _AUTH_ERROR_NAMES or "401" in scrub_text(str(e)):
            logger.info("Сбрасываю клиент GigaChat после ошибки авторизации")
            reset_gigachat_client()

        return None


DEFAULT_SYSTEM_PROMPT = (
    "Ты — опытный копирайтер транспортной компании. "
    "Пиши короткие живые коммерческие предложения по автоперевозкам на русском языке. "
    "Стиль энергичный, как в чате перевозчиков, но без ошибок и лишней воды. "
    "Не выдумывай цены, названия компаний и сроки. Не добавляй подписи и обращения."
)


# ============================================
# OLLAMA — ПРОВЕРКА И ГЕНЕРАЦИЯ
# ============================================


def _ollama_host() -> str:
    return os.environ.get("OLLAMA_HOST", DEFAULT_OLLAMA_HOST).rstrip("/")


def get_ollama_models(timeout: float = 2.0) -> List[str]:
    """Возвращает список доступных моделей Ollama (пустой список при ошибке)."""
    try:
        import requests

        response = requests.get(f"{_ollama_host()}/api/tags", timeout=timeout)
        if response.status_code == 200:
            return [m.get("name", "") for m in response.json().get("models", [])]
        return []
    except Exception:
        return []


def is_ollama_available(timeout: float = 2.0) -> bool:
    """Проверяет, запущен ли Ollama и есть ли хотя бы одна модель."""
    models = get_ollama_models(timeout=timeout)
    if models:
        return True

    # Сервер может быть запущен, но без моделей — проверим сам факт доступности
    try:
        import requests

        response = requests.get(f"{_ollama_host()}/api/tags", timeout=timeout)
        return response.status_code == 200 and bool(response.json().get("models"))
    except Exception:
        return False


def generate_ollama(
    prompt: str,
    model: Optional[str] = None,
    temperature: float = 0.8,
    timeout: int = DEFAULT_OLLAMA_TIMEOUT,
    max_tokens: int = 300,
    system_prompt: Optional[str] = None,
) -> Optional[str]:
    """
    Генерирует текст через локальный Ollama.

    Returns:
        Текст ответа или None при любой ошибке. Исключения не выбрасываются.
    """
    if not prompt or not str(prompt).strip():
        logger.warning("generate_ollama: пустой prompt")
        return None

    try:
        import requests
    except ImportError:
        logger.error("Библиотека requests не установлена — Ollama недоступен")
        return None

    models = get_ollama_models()
    if not models:
        logger.warning("generate_ollama: Ollama недоступен или нет моделей")
        return None

    if not model or model not in models:
        model = models[0]
        logger.info("Автоматически выбрана модель Ollama: %s", model)

    try:
        temperature = float(temperature)
    except (TypeError, ValueError):
        temperature = 0.8

    payload: Dict[str, Any] = {
        "model": model,
        "prompt": str(prompt),
        "stream": False,
        "options": {
            "temperature": temperature,
            "top_p": 0.9,
            "top_k": 50,
            "num_predict": int(max_tokens),
            "repeat_penalty": 1.2,
        },
    }
    if system_prompt:
        payload["system"] = system_prompt

    start = time.time()
    try:
        response = requests.post(
            f"{_ollama_host()}/api/generate", json=payload, timeout=timeout
        )
        elapsed = time.time() - start
        logger.info(
            "Ollama: ответ за %.2f сек (HTTP %s, модель %s)",
            elapsed,
            response.status_code,
            model,
        )

        if response.status_code != 200:
            logger.error("Ollama вернул HTTP %s", response.status_code)
            return None

        text = (response.json().get("response") or "").strip()
        if not text:
            logger.warning("Ollama вернул пустой ответ")
            return None

        return text

    except requests.Timeout:
        logger.error("Таймаут запроса к Ollama (%s сек)", timeout)
        return None
    except Exception as e:
        logger.error("Ошибка запроса к Ollama: %s", scrub_text(str(e)))
        return None


# ============================================
# ЕДИНАЯ ТОЧКА ВХОДА
# ============================================


def get_available_provider() -> str:
    """
    Определяет лучший доступный провайдер.

    Returns:
        "gigachat" | "ollama" | "none"
    """
    if is_gigachat_configured():
        return "gigachat"
    if is_ollama_available():
        return "ollama"
    return "none"


def generate_text(
    prompt: str,
    provider: str = "gigachat",
    system_prompt: Optional[str] = None,
    temperature: float = 0.85,
    max_tokens: int = 280,
    ollama_model: Optional[str] = None,
    ollama_timeout: int = DEFAULT_OLLAMA_TIMEOUT,
) -> Optional[str]:
    """
    Единая точка генерации текста.

    provider="gigachat" — GigaChat, при неудаче автоматический откат на Ollama
    provider="ollama"   — только Ollama
    provider="auto"     — GigaChat, если настроен, иначе Ollama

    Returns:
        Текст или None, если оба провайдера не сработали.
    """
    if provider == "auto":
        provider = get_available_provider()

    if provider == "gigachat":
        result = generate_gigachat(
            prompt=prompt,
            system_prompt=system_prompt,
            temperature=temperature,
            max_tokens=max_tokens,
        )
        if result:
            return result

        logger.info("GigaChat не дал результат — пробую Ollama как запасной вариант")
        return generate_ollama(
            prompt=prompt,
            model=ollama_model,
            temperature=temperature,
            timeout=ollama_timeout,
            max_tokens=max_tokens,
            system_prompt=system_prompt,
        )

    if provider == "ollama":
        return generate_ollama(
            prompt=prompt,
            model=ollama_model,
            temperature=temperature,
            timeout=ollama_timeout,
            max_tokens=max_tokens,
            system_prompt=system_prompt,
        )

    logger.warning("generate_text: нет доступного провайдера (provider=%s)", provider)
    return None


def get_status_info() -> Dict[str, Any]:
    """Краткая сводка о состоянии провайдеров — для логов и UI."""
    try:
        model = os.environ.get("GIGACHAT_MODEL", DEFAULT_GIGACHAT_MODEL)
    except Exception:
        model = DEFAULT_GIGACHAT_MODEL

    return {
        "gigachat_configured": is_gigachat_configured(),
        "gigachat_model": model,
        "gigachat_credentials": mask_credentials(get_gigachat_credentials()),
        "ollama_available": is_ollama_available(),
        "ollama_models": get_ollama_models(),
        "env_path": str(ENV_PATH),
        "env_exists": ENV_PATH.exists(),
    }


# ============================================
# САМОПРОВЕРКА
# ============================================

if __name__ == "__main__":
    # Консоль Windows часто в cp1251 — переключаем вывод в UTF-8,
    # чтобы русский текст и эмодзи печатались, а не падали с UnicodeEncodeError
    try:
        import sys

        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    logger.info("=" * 70)
    logger.info("САМОПРОВЕРКА llm_provider")
    logger.info("=" * 70)

    info = get_status_info()
    print(f"Файл .env:            {info['env_path']} (существует: {info['env_exists']})")
    print(f"GigaChat настроен:    {'✅ да' if info['gigachat_configured'] else '❌ нет'}")
    print(f"GigaChat credentials: {info['gigachat_credentials']}")
    print(f"GigaChat модель:      {info['gigachat_model']}")
    print(f"Ollama доступен:      {'✅ да' if info['ollama_available'] else '❌ нет'}")
    print(f"Ollama модели:        {', '.join(info['ollama_models']) or '—'}")
    print(f"Активный провайдер:   {get_available_provider()}")
