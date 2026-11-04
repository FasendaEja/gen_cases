from __future__ import annotations

import hashlib
import json
import logging
import re
from dataclasses import dataclass
from typing import Any

from .config import Settings


class LLMError(RuntimeError):
    """Ошибка при обращении к модели."""


class _NextMode(Exception):
    """Провайдер не поддерживает выбранный способ запроса JSON."""


@dataclass
class Usage:
    api_calls: int = 0
    cache_hits: int = 0
    prompt_tokens: int = 0
    completion_tokens: int = 0
    failed_calls: int = 0

    @property
    def total_tokens(self) -> int:
        return self.prompt_tokens + self.completion_tokens

    def as_dict(self) -> dict[str, int]:
        return {
            "api_calls": self.api_calls,
            "cache_hits": self.cache_hits,
            "failed_calls": self.failed_calls,
            "prompt_tokens": self.prompt_tokens,
            "completion_tokens": self.completion_tokens,
            "total_tokens": self.total_tokens,
        }


_FENCE_RE = re.compile(r"^```(?:json)?\s*(.*?)\s*```$", re.DOTALL | re.IGNORECASE)


def parse_json_payload(content: str) -> dict[str, Any]:
    """Достаёт JSON-объект из ответа модели (в т.ч. из ```json ... ```)."""
    text = (content or "").strip()
    if not text:
        raise ValueError("Ответ модели пуст")

    fenced = _FENCE_RE.match(text)
    if fenced:
        text = fenced.group(1).strip()

    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start == -1 or end <= start:
            raise ValueError("В ответе модели нет JSON-объекта")
        payload = json.loads(text[start:end + 1])

    if not isinstance(payload, dict):
        raise ValueError("Ожидался JSON-объект, получен " + type(payload).__name__)
    return payload


class LLMClient:
    """Обёртка над OpenAI chat.completions с JSON-схемой, ретраями и кэшем."""

    def __init__(
        self,
        settings: Settings,
        logger: "logging.Logger | None" = None,
    ) -> None:
        self.settings = settings
        self.log = logger or logging.getLogger(__name__)
        self.usage = Usage()
        self._client = None
        self._cache_dir = settings.llm_cache_dir if settings.llm_cache else None
        if self._cache_dir is not None:
            self._cache_dir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------ client
    def _openai_client(self):  # noqa: ANN202 - тип из внешней библиотеки
        if self._client is not None:
            return self._client
        try:
            from openai import OpenAI  # type: ignore
        except ImportError as exc:  # pragma: no cover
            raise LLMError(
                "Не установлен пакет openai. Выполните: pip install -r requirements.txt"
            ) from exc

        kwargs: dict[str, Any] = {
            "api_key": self.settings.openai_api_key,
            "timeout": self.settings.openai_timeout,
            "max_retries": self.settings.openai_max_retries,
        }
        if self.settings.openai_base_url:
            kwargs["base_url"] = self.settings.openai_base_url
        self._client = OpenAI(**kwargs)
        return self._client

    # ------------------------------------------------------------------ кэш
    def _cache_key(self, system: str, user: str, schema: dict, temperature: "float | None") -> str:
        payload = json.dumps(
            {
                "model": self.settings.openai_model,
                "temperature": temperature,
                "system": system,
                "user": user,
                "schema": schema,
            },
            ensure_ascii=False,
            sort_keys=True,
        )
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def _load_cache(self, key: str) -> "dict | None":
        if self._cache_dir is None:
            return None
        path = self._cache_dir / f"{key}.json"
        if not path.is_file():
            return None
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            return None
        response = payload.get("response")
        return response if isinstance(response, dict) else None

    def _store_cache(self, key: str, response: dict) -> None:
        if self._cache_dir is None:
            return
        path = self._cache_dir / f"{key}.json"
        try:
            path.write_text(
                json.dumps({"response": response}, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except OSError as exc:  # pragma: no cover
            self.log.debug("Не удалось записать кэш ответа: %s", exc)

    # --------------------------------------------------------------- промпты
    # Рендеры промптов сохраняет Pipeline (он знает каталог запуска и настройку
    # SAVE_PROMPTS) — так они пишутся независимо от реализации LLM-клиента.

    # ------------------------------------------------------------------ вызов
    def json_call(
        self,
        *,
        system: str,
        user: str,
        schema: dict,
        schema_name: str,
        prompt_name: str = "request",
        temperature: "float | None | str" = "default",
    ) -> dict[str, Any]:
        api_temperature = (
            self.settings.openai_temperature if temperature == "default" else temperature  # type: ignore[comparison-overlap]
        )
        cache_key = self._cache_key(system, user, schema, api_temperature)
        cached = self._load_cache(cache_key)
        if cached is not None:
            self.usage.cache_hits += 1
            self.log.debug("Ответ модели из кэша: %s", prompt_name)
            return cached

        errors: list[str] = []
        for mode in self._modes():
            try:
                response = self._request(system, user, schema, schema_name, mode, api_temperature)
            except _NextMode as exc:
                errors.append(f"{mode}: {exc}")
                self.log.debug("Режим %s не поддержан провайдером, пробую следующий", mode)
                continue
            self._store_cache(cache_key, response)
            return response

        raise LLMError(
            "Не удалось получить корректный JSON от модели. Попытки: " + "; ".join(errors[-3:])
        )

    def _modes(self) -> list[str]:
        mode = self.settings.openai_json_mode
        if mode == "auto":
            return ["json_schema", "json_object", "text"]
        return [mode]

    def _request(
        self,
        system: str,
        user: str,
        schema: dict,
        schema_name: str,
        mode: str,
        temperature: "float | None",
    ) -> dict[str, Any]:
        client = self._openai_client()
        messages: list[dict[str, str]] = [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ]
        params: dict[str, Any] = {
            "model": self.settings.openai_model,
            "messages": messages,
        }
        if temperature is not None:
            params["temperature"] = temperature
        if self.settings.openai_max_tokens:
            params["max_tokens"] = self.settings.openai_max_tokens
        if mode == "json_schema":
            params["response_format"] = {
                "type": "json_schema",
                "json_schema": {
                    "name": schema_name,
                    "strict": True,
                    "schema": schema,
                },
            }
        elif mode == "json_object":
            params["response_format"] = {"type": "json_object"}

        adjusted_temperature = False
        adjusted_max_tokens = False
        repair_attempts = 0

        while True:
            self.usage.api_calls += 1
            try:
                response = client.chat.completions.create(**params)
            except Exception as exc:  # noqa: BLE001 - SDK кидает разные типы
                message = str(exc)
                lowered = message.lower()

                if (
                    not adjusted_temperature
                    and "temperature" in lowered
                    and any(word in lowered for word in ("unsupported", "not support", "does not support", "invalid"))
                    and "temperature" in params
                ):
                    params.pop("temperature", None)
                    adjusted_temperature = True
                    self.log.debug("Модель не принимает temperature — параметр убран")
                    continue

                if (
                    not adjusted_max_tokens
                    and "max_tokens" in lowered
                    and "max_completion_tokens" in lowered
                ):
                    if "max_tokens" in params:
                        params["max_completion_tokens"] = params.pop("max_tokens")
                    adjusted_max_tokens = True
                    self.log.debug("Провайдер требует max_completion_tokens")
                    continue

                if self._is_format_error(lowered):
                    raise _NextMode(message) from exc

                self.usage.failed_calls += 1
                raise LLMError(
                    f"Ошибка запроса к модели {self.settings.openai_model}: {message}"
                ) from exc

            self._accumulate_usage(response)
            content = self._content_of(response)

            try:
                return parse_json_payload(content)
            except ValueError as exc:
                repair_attempts += 1
                if repair_attempts > 1:
                    self.usage.failed_calls += 1
                    snippet = (content or "")[:500]
                    raise LLMError(
                        "Модель вернула невалидный JSON дважды. "
                        f"Фрагмент ответа: {snippet!r}"
                    ) from exc
                self.log.warning("Невалидный JSON от модели, запрашиваю исправление")
                params["messages"] = [
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                    {"role": "assistant", "content": content},
                    {
                        "role": "user",
                        "content": (
                            "Предыдущий ответ не является валидным JSON. "
                            "Верни ТОЛЬКО корректный JSON-объект по заданной схеме, "
                            "без markdown, комментариев и текста вокруг."
                        ),
                    },
                ]

    # ---------------------------------------------------------------- helpers
    @staticmethod
    def _is_format_error(lowered: str) -> bool:
        mentions_format = "response_format" in lowered or "json_schema" in lowered or "json schema" in lowered
        mentions_unsupported = any(
            word in lowered
            for word in ("unsupported", "not supported", "does not support", "invalid", "unknown")
        )
        return mentions_format and mentions_unsupported

    def _accumulate_usage(self, response: Any) -> None:
        usage = getattr(response, "usage", None)
        if usage is None:
            return
        self.usage.prompt_tokens += int(getattr(usage, "prompt_tokens", 0) or 0)
        self.usage.completion_tokens += int(getattr(usage, "completion_tokens", 0) or 0)

    @staticmethod
    def _content_of(response: Any) -> str:
        try:
            message = response.choices[0].message
        except (AttributeError, IndexError, TypeError) as exc:
            raise LLMError("Неожиданный формат ответа модели (нет choices[0])") from exc
        content = getattr(message, "content", None)
        if content is None:
            refusal = getattr(message, "refusal", None)
            raise LLMError(f"Модель отказалась отвечать: {refusal or 'пустой ответ'}")
        return content


class DryRunLLM:
    """Заглушка для --dry-run: не обращается к сети и не тратит токены.

    Промпты сохраняет Pipeline, поэтому их можно посмотреть после запуска.
    """

    def __init__(self) -> None:
        self.usage = Usage()
        self.counter = 0
        self.calls: list[str] = []

    def json_call(
        self,
        *,
        system: str,
        user: str,
        schema: dict,
        schema_name: str,
        prompt_name: str = "request",
        temperature: "float | None | str" = "default",
    ) -> dict[str, Any]:
        self.counter += 1
        self.calls.append(prompt_name)
        if schema_name.startswith("checklist"):
            return {"items": []}
        return {"cases": []}
