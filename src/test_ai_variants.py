"""Проверки полного ИИ-абзаца, истории и HTTP-запроса без сети и ключа.

Запуск: python -m unittest discover -s src -p test_ai_variants.py -v
AST позволяет проверить реальные функции, не запуская глобальное окно Tkinter.
GigaChat и httpx используются уже установленные, новые зависимости не нужны.
"""

import ast
import json
import logging
import os
import re
import time
import unittest
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from unittest.mock import patch

import text_variants as tv

SRC = Path(__file__).parent


def load_definitions(filename, names, constants, namespace):
    tree = ast.parse((SRC / filename).read_text(encoding="utf-8"))
    nodes = []
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.ClassDef)) and node.name in names:
            nodes.append(node)
        elif isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id in constants
            for target in node.targets
        ):
            nodes.append(node)
    exec(compile(ast.Module(body=nodes, type_ignores=[]), filename, "exec"), namespace)


class AIVariantsTest(unittest.TestCase):
    def setUp(self):
        tv._history.clear()
        self.g = dict(
            Dict=Dict, List=List, Optional=Optional, Tuple=Tuple, Any=Any,
            re=re, json=json, time=time, os=os,
            logger=logging.getLogger("test_ai_variants"),
            DEFAULT_PAYMENT_KEY=tv.DEFAULT_PAYMENT_KEY,
            CREATIVE_MODE=tv.CREATIVE_MODE, CREATIVITY=0.9,
            GIGACHAT_READY=True, USE_AI=True,
            _AI_PREVIOUS={}, _AI_CHOICES=tv.RecentHistory(depth=4),
            _PHRASES=tv._history, STATUSES=tv.STATUSES, LEGALS=tv.LEGALS,
            assemble_text=tv.assemble_text, build_route_block=tv.build_route_block,
            payment_line=tv.payment_line, build_cta=tv.build_cta,
            normalize_cities=tv.normalize_cities, route_origin=tv.route_origin,
            destinations=tv.destinations,
            pick_opener=tv.pick_opener, get_current_region=lambda: "Юга",
            _generation_in_progress=False,
        )
        load_definitions("llm_provider.py", {
            "_generate_gigachat_sync", "_extract_message_content",
        }, {"PAYMENT_TYPES", "DEFAULT_GIGACHAT_MODEL"}, self.g)
        load_definitions("генератор_кп_gigachat.py", {
            "AIGuard", "AICache", "_ai_key", "_remember_ai_text", "_ai_facts",
            "_build_prompt", "_word_overlap", "_opening_key", "_validate_ai_paragraph",
            "generate_with_gigachat", "generate_text_sync", "_build_fallback_paragraph", "format_cities_list",
            "_guess_city_count", "build_template_text", "new_variant",
        }, {
            "_AI_STYLES", "_AI_STRUCTURES", "GIGACHAT_SYSTEM_PROMPT", "AI_TEMPERATURE",
            "AI_MAX_TOKENS", "AI_MAX_ATTEMPTS", "ALL_CITIES_SOUTH",
            "ALL_CITIES_EAST", "ALL_CITIES_VLADIVOSTOK",
        }, self.g)
        self.calls = []
        self.answer = (
            "Машины — груз для перевозки автовозом. "
            "Москва → Краснодар → Ростов-на-Дону; Оплата по безналу (с НДС). "
            "Напишите в личку, обсудим детали."
        )
        self.answers = []

        def fake_api(**kwargs):
            self.calls.append(kwargs)
            return self.answers.pop(0) if self.answers else self.answer

        self.g["generate_gigachat"] = fake_api

    def generate(self, cities=None, pay="beznal_nds", use_cache=False):
        return self.g["generate_text_sync"](
            "Юга", cities or ["Краснодар", "Ростов-на-Дону"],
            use_cache=use_cache, pay_key=pay,
        )

    def test_exact_temperature_ignores_old_settings(self):
        for value in (0.1, 0.85, 0.95, 1.0, float("nan")):
            self.g["CREATIVITY"] = value
            self.g["generate_with_gigachat"]("Юга", ["Краснодар", "Ростов-на-Дону"])
            self.assertEqual(0.9, self.calls[-1]["temperature"])
            self.assertEqual(300, self.calls[-1]["max_tokens"])

    def test_whole_paragraph_returned_without_template_assembly(self):
        result = self.generate()
        self.assertEqual(self.answer, result)
        self.assertNotIn("\n", result)

    def test_prompt_has_five_negative_examples_and_no_fixed_example(self):
        previous = [f"Предыдущий текст {i}" for i in range(7)]
        prompts = [self.g["_build_prompt"](
            "Владивосток", ["Владивосток → Москва"], "beznal_nds", previous,
        ) for _ in range(4)]
        self.assertEqual(4, len({p.split("Стиль: ")[1].splitlines()[0] for p in prompts}))
        self.assertIn(json.dumps(previous[-5:], ensure_ascii=False), prompts[0])
        self.assertNotIn("Предыдущий текст 0", prompts[0])
        self.assertIn('"route": "Владивосток → Москва"', prompts[0])
        self.assertIn("один цельный абзац", prompts[0])
        self.assertNotIn("Пример удачного ответа", prompts[0])

    def test_overlap_threshold(self):
        overlap = self.g["_word_overlap"]
        self.assertEqual(1, overlap("Те же слова!", "слова те же"))
        self.assertAlmostEqual(0.7, overlap("а б в г д е ж", "а б в г д е ж з и к"))
        self.assertGreater(overlap("а б в г д е ж з", "а б в г д е ж з и к"), 0.7)
        self.assertEqual(0, overlap("", ""))

    def test_repeat_of_older_history_entry_retries_then_falls_back(self):
        key = self.g["_ai_key"]("Юга", ["Краснодар", "Ростов-на-Дону"], "beznal_nds")
        for text in (self.answer, "Другой ранний текст", "Третий текст", "Четвёртый", "Последний"):
            self.g["_remember_ai_text"](key, text)
        result = self.generate()
        self.assertEqual(3, len(self.calls))
        self.assertIn("Слишком похоже на предыдущий текст, перепиши другими словами",
                      self.calls[1]["prompt"])
        self.assertIn(json.dumps(self.answer, ensure_ascii=False), self.calls[0]["prompt"])
        self.assertNotEqual(self.answer, result)
        self.assertIn("Краснодар", result)
        self.assertIn("НДС", result)

    def test_rewritten_answer_accepted_after_repeat(self):
        self.generate()
        fresh = (
            "Коллеги, ищем перевозчика: под автовоз есть автомобили. "
            "Оплата по безналу (с НДС); Москва → Краснодар → Ростов-на-Дону. "
            "За подробностями обращайтесь личным сообщением."
        )
        self.answers = [self.answer, fresh]
        self.assertEqual(fresh, self.generate(use_cache=True))
        self.assertEqual(3, len(self.calls))

    def test_missing_route_payment_or_wrong_vat_rejected(self):
        facts = self.g["_ai_facts"]("Юга", ["Краснодар", "Ростов-на-Дону"], "beznal_nds")
        for bad in (
            self.answer.replace("Ростов-на-Дону", "Омск"),
            self.answer.replace("Оплата по безналу (с НДС)", "Оплата наличными"),
            self.answer + " Можно без НДС.",
            self.answer + " Возможны наличные.",
            self.answer + " Также Тюмень.",
            self.answer.replace(". Москва", ".\nМосква"),
        ):
            with self.subTest(bad=bad):
                self.assertFalse(self.g["_validate_ai_paragraph"](
                    bad, facts, ["Краснодар", "Ростов-на-Дону"],
                )[0])

    def test_prices_brands_and_invented_conditions_rejected(self):
        facts = self.g["_ai_facts"]("Юга", ["Краснодар", "Ростов-на-Дону"], "beznal_nds")
        for extra in ("цена сто тысяч рублей", "Toyota", "Лада", "сегодня", "50000 ₽"):
            with self.subTest(extra=extra):
                self.assertFalse(self.g["_validate_ai_paragraph"](
                    self.answer + " " + extra, facts, ["Краснодар", "Ростов-на-Дону"],
                )[0])
        self.assertTrue(self.g["AIGuard"].validate_ai_response(self.answer + " Машины со склада.")[0])
        self.assertTrue(self.g["AIGuard"].validate_ai_response(self.answer + " Доставка машин автовозом.")[0])
        facts = self.g["_ai_facts"]("Юга", ["Краснодар", "Ростов-на-Дону"], "beznal_nds")
        self.assertTrue(self.g["_validate_ai_paragraph"](
            self.answer + " Безналичный расчёт.", facts, ["Краснодар", "Ростов-на-Дону"],
        )[0])

    def test_invalid_answer_also_retries(self):
        self.answers = ["Нет маршрута и оплаты", self.answer]
        self.assertEqual(self.answer, self.generate())
        self.assertEqual(2, len(self.calls))
        self.assertIn("Исправь ошибку", self.calls[1]["prompt"])

    def test_unavailable_api_immediate_fallback(self):
        self.answer = None
        text = self.generate()
        self.assertEqual(1, len(self.calls))
        self.assertIn("автовоз", text.lower())
        self.assertIn("Краснодар", text)
        samples = [text] + [self.generate() for _ in range(4)]
        self.assertEqual(5, len(set(samples)))
        self.assertEqual(5, len({self.g["_opening_key"](sample) for sample in samples}))
        self.assertTrue(all("\n" not in sample for sample in samples))
        self.assertTrue(all("готовы к отправке" not in sample.lower() for sample in samples))

    def test_all_payments_and_creative_route(self):
        for pay in self.g["PAYMENT_TYPES"]:
            facts = self.g["_ai_facts"]("Юга", ["Краснодар"], pay)
            text = f"Автомобили — груз для автовоза. {facts['route']}; {facts['payment']}. Пишите в личку."
            self.assertTrue(self.g["_validate_ai_paragraph"](text, facts, ["Краснодар"])[0])
        self.assertEqual("Москва → Восток", self.g["_ai_facts"]("Восток", [], "cash")["route"])
        self.assertEqual("Владивосток → Россия", self.g["_ai_facts"]("Владивосток", [], "cash")["route"])

    def test_real_world_payment_and_role_regressions(self):
        facts = self.g["_ai_facts"]("Восток", [], "beznal_nds")
        text = (
            "Есть свободный автовоз? Нужна перевозка наших машин из Москвы на восток. "
            "Оплата безналичным расчетом с учетом НДС. Напишите в личку, договоримся."
        )
        self.assertTrue(self.g["_validate_ai_paragraph"](text, facts, [])[0])
        self.assertFalse(self.g["_validate_ai_paragraph"](
            text.replace("из Москвы на восток", "с востока в Москву"), facts, [],
        )[0])
        self.assertFalse(self.g["_validate_ai_paragraph"](
            text + " Везём на нашем автовозе.", facts, [],
        )[0])
        self.assertFalse(self.g["_validate_ai_paragraph"](
            "Предлагает перевозчику автовоз. " + text, facts, [],
        )[0])
        self.assertFalse(self.g["_validate_ai_paragraph"](
            "Уважаемые коллеги! " + text, facts, [],
        )[0])
        self.assertFalse(self.g["_validate_ai_paragraph"](
            text + " Машины готовы к отправке.", facts, [],
        )[0])

    def test_new_variant_keeps_ai_history(self):
        history = self.g["_AI_PREVIOUS"]
        history[("Юга", ("Краснодар",), "beznal_nds")] = [self.answer]
        tv._history.remember("probe", "а")
        self.g["refresh_preview_async"] = lambda: None
        self.g["new_variant"]()
        self.assertEqual([self.answer], next(iter(history.values())))
        self.assertEqual("б", tv._history.pick("probe", ["а", "б"]))

    def test_history_bounded_to_five_texts_and_32_keys(self):
        for index in range(100):
            self.g["_remember_ai_text"]((index,), "текст")
        self.assertEqual(32, len(self.g["_AI_PREVIOUS"]))
        for index in range(8):
            self.g["_remember_ai_text"]((99,), f"Вариант {index}")
        self.assertEqual([f"Вариант {i}" for i in range(3, 8)], self.g["_AI_PREVIOUS"][(99,)])

    def test_real_sdk_http_json_contains_temperature_and_messages(self):
        # Используем настоящий SDK и его HTTP-сериализацию. Только транспорт
        # подставной: сетевой запрос и авторизация не выполняются.
        import httpx
        from gigachat import GigaChat
        requests = []
        previous = [f"Старый вариант {i}" for i in range(5)]
        key = self.g["_ai_key"]("Юга", ["Краснодар", "Ростов-на-Дону"], "beznal_nds")
        self.g["_AI_PREVIOUS"][key] = previous.copy()

        def handle(request):
            requests.append(json.loads(request.content))
            return httpx.Response(200, json={
                "choices": [{"message": {"role": "assistant", "content": self.answer},
                             "index": 0, "finish_reason": "stop"}],
                "model": "GigaChat",
                "usage": {"prompt_tokens": 10, "completion_tokens": 20, "total_tokens": 30},
                "created": 1, "object": "chat.completion",
            })

        with patch.dict(os.environ, {"GIGACHAT_MODEL": "GigaChat"}):
            with GigaChat(access_token="offline-test-token", model="GigaChat") as client:
                client._client_instance = httpx.Client(
                    base_url="https://offline.invalid/api/v1/",
                    transport=httpx.MockTransport(handle),
                )
                def real_sdk_boundary(**kwargs):
                    return self.g["_generate_gigachat_sync"](
                        client, kwargs["system_prompt"], kwargs["prompt"],
                        kwargs["temperature"], kwargs["max_tokens"],
                    )
                self.g["generate_gigachat"] = real_sdk_boundary
                self.assertEqual(self.answer, self.generate())
        self.assertEqual(1, len(requests))
        self.assertEqual(0.9, requests[0]["temperature"])
        self.assertEqual(300, requests[0]["max_tokens"])
        self.assertEqual("GigaChat", requests[0]["model"])
        self.assertEqual(["system", "user"], [m["role"] for m in requests[0]["messages"]])
        self.assertIn("previous_texts", requests[0]["messages"][1]["content"])
        self.assertIn(json.dumps(previous, ensure_ascii=False), requests[0]["messages"][1]["content"])


if __name__ == "__main__":
    unittest.main()
