"""Офлайн-тесты LLM-клиента без пакета openai и без сети.

Запуск:  python tests/llm_test.py
"""

from __future__ import annotations

import sys
import tempfile
import types
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from doqa_ai_cases.config import Settings  # noqa: E402
from doqa_ai_cases.llm import LLMClient, LLMError, parse_json_payload  # noqa: E402
from doqa_ai_cases.schemas import CHECKLIST_SCHEMA  # noqa: E402


class _FakeCompletions:
    def __init__(self, owner: "_FakeOpenAI") -> None:
        self._owner = owner

    def create(self, **params):
        return self._owner.behavior(params)


class _FakeOpenAI:
    """Минимальная замена openai.OpenAI: поведение задаётся снаружи."""

    behavior = staticmethod(lambda params: _response('{"items": []}'))
    instances: list["_FakeOpenAI"] = []

    def __init__(self, **kwargs) -> None:
        self.kwargs = kwargs
        self.chat = SimpleNamespace(completions=_FakeCompletions(self))
        _FakeOpenAI.instances.append(self)


def _response(content: str, prompt_tokens: int = 10, completion_tokens: int = 5):
    return SimpleNamespace(
        choices=[SimpleNamespace(message=SimpleNamespace(content=content, refusal=None))],
        usage=SimpleNamespace(prompt_tokens=prompt_tokens, completion_tokens=completion_tokens),
    )


def _install_fake_openai() -> "types.ModuleType":
    module = types.ModuleType("openai")
    module.OpenAI = _FakeOpenAI
    sys.modules["openai"] = module
    return module


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def _client(**overrides) -> LLMClient:
    settings = Settings(
        openai_api_key="test-key",
        openai_model="test-model",
        openai_temperature=0.2,
        llm_cache=False,
        **overrides,
    )
    return LLMClient(settings)


def test_temperature_fallback() -> None:
    calls: list[dict] = []

    def behavior(params):
        calls.append(dict(params))
        if "temperature" in params:
            raise RuntimeError(
                "Unsupported parameter: 'temperature' is not supported with this model."
            )
        return _response('{"items": []}')

    _FakeOpenAI.behavior = staticmethod(behavior)
    client = _client()
    result = client.json_call(
        system="s", user="u", schema=CHECKLIST_SCHEMA, schema_name="checklist"
    )
    check(result == {"items": []}, "не разобран ответ после fallback по temperature")
    check(len(calls) == 2, f"ожидалось 2 запроса, выполнено {len(calls)}")
    check("temperature" not in calls[1], "temperature не убран из повторного запроса")
    check(client.usage.api_calls == 2, "неверно посчитаны вызовы API")


def test_json_mode_fallback() -> None:
    modes: list[str] = []

    def behavior(params):
        mode = (params.get("response_format") or {}).get("type", "text")
        modes.append(mode)
        if mode == "json_schema":
            raise RuntimeError(
                "Invalid parameter: 'response_format' json_schema is not supported by this gateway"
            )
        return _response('```json\n{"cases": []}\n```')

    _FakeOpenAI.behavior = staticmethod(behavior)
    result = _client().json_call(
        system="s", user="u", schema=CHECKLIST_SCHEMA, schema_name="cases"
    )
    check(result == {"cases": []}, "не разобрали JSON в markdown-обёртке")
    check(modes == ["json_schema", "json_object"], f"неверная цепочка режимов: {modes}")


def test_max_tokens_adaptation() -> None:
    calls: list[dict] = []

    def behavior(params):
        calls.append(dict(params))
        if "max_tokens" in params:
            raise RuntimeError(
                "Unsupported parameter: 'max_tokens' is not supported with this model. "
                "Use 'max_completion_tokens' instead."
            )
        check(params.get("max_completion_tokens") == 128, "max_completion_tokens не передан")
        return _response('{"items": []}')

    _FakeOpenAI.behavior = staticmethod(behavior)
    result = _client(openai_max_tokens=128).json_call(
        system="s", user="u", schema=CHECKLIST_SCHEMA, schema_name="checklist"
    )
    check(result == {"items": []}, "ответ после адаптации max_tokens не разобран")
    check(len(calls) == 2, "не было повтора после ошибки max_tokens")


def test_invalid_json_repair() -> None:
    responses = ["это не json", '{"items": [{"area": "A"}]}']
    calls: list[dict] = []

    def behavior(params):
        calls.append(dict(params))
        return _response(responses.pop(0))

    _FakeOpenAI.behavior = staticmethod(behavior)
    result = _client().json_call(
        system="s", user="u", schema=CHECKLIST_SCHEMA, schema_name="checklist"
    )
    check(result["items"][0]["area"] == "A", "не разобран ответ после исправления JSON")
    check(len(calls) == 2, "не был выполнен запрос на исправление JSON")
    messages = calls[1]["messages"]
    check(any(message["role"] == "assistant" for message in messages), "нет assistant-сообщения")
    check(messages[-1]["role"] == "user", "нет корректирующей инструкции")


def test_invalid_json_twice_raises() -> None:
    _FakeOpenAI.behavior = staticmethod(lambda params: _response("мусор"))
    client = _client()
    try:
        client.json_call(system="s", user="u", schema=CHECKLIST_SCHEMA, schema_name="checklist")
    except LLMError as exc:
        check("невалидный json" in str(exc).lower(), f"неожиданный текст ошибки: {exc}")
    else:
        raise AssertionError("LLMError не поднят после двух невалидных ответов")


def test_cache(tmp: Path) -> None:
    calls = {"count": 0}

    def behavior(params):
        calls["count"] += 1
        return _response('{"items": []}')

    _FakeOpenAI.behavior = staticmethod(behavior)
    settings = Settings(
        openai_api_key="test-key",
        openai_model="test-model",
        llm_cache=True,
        llm_cache_dir=tmp / "cache",
    )
    client = LLMClient(settings)
    client.json_call(system="s", user="u", schema=CHECKLIST_SCHEMA, schema_name="checklist")
    client.json_call(system="s", user="u", schema=CHECKLIST_SCHEMA, schema_name="checklist")
    check(calls["count"] == 1, f"кэш не сработал: {calls['count']} запросов")
    check(client.usage.cache_hits == 1, "не посчитан cache hit")
    check(len(list((tmp / "cache").glob("*.json"))) == 1, "файл кэша не создан")


def test_parse_json_payload() -> None:
    check(parse_json_payload('{"a": 1}') == {"a": 1}, "обычный JSON не разобран")
    check(
        parse_json_payload("Вот результат:\n```json\n{\"a\": 2}\n```") == {"a": 2},
        "JSON из markdown не разобран",
    )
    check(
        parse_json_payload('Пояснение {"a": 3} конец') == {"a": 3},
        "JSON внутри текста не разобран",
    )


def main() -> int:
    _install_fake_openai()
    try:
        test_temperature_fallback()
        test_json_mode_fallback()
        test_max_tokens_adaptation()
        test_invalid_json_repair()
        test_invalid_json_twice_raises()
        test_parse_json_payload()
        with tempfile.TemporaryDirectory(prefix="doqa_llm_") as tmp_dir:
            test_cache(Path(tmp_dir))
    finally:
        sys.modules.pop("openai", None)
    print("LLM TESTS OK: все проверки пройдены")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
