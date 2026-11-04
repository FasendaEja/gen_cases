from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field
from typing import Any


class ModelError(ValueError):
    """Данные от модели/из выгрузки не соответствуют ожидаемой структуре."""


# --------------------------------------------------------------------------- #
#  Нормализация значений
# --------------------------------------------------------------------------- #

_PRIORITY_ALIASES = {
    "high": "high",
    "critical": "high",
    "blocker": "high",
    "высок": "high",
    "критич": "high",
    "medium": "medium",
    "normal": "medium",
    "major": "medium",
    "средн": "medium",
    "обычн": "medium",
    "low": "low",
    "minor": "low",
    "trivial": "low",
    "низк": "low",
    "незнач": "low",
}

_TYPE_ALIASES = {
    "positive": "positive",
    "позитив": "positive",
    "negative": "negative",
    "негатив": "negative",
    "boundary": "boundary",
    "гранич": "boundary",
    "integration": "integration",
    "интеграц": "integration",
    "security": "security",
    "безопас": "security",
    "ui": "ui",
    "ux": "ui",
    "интерфейс": "ui",
    "performance": "performance",
    "производ": "performance",
    "нагруз": "performance",
    "other": "other",
    "другое": "other",
}


def text_of(value: Any) -> str:
    """Аккуратно приводит значение из JSON/Excel к строке."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (list, tuple)):
        return "\n".join(item for item in (text_of(v) for v in value) if item)
    return str(value).strip()


def _dedupe(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        key = value.strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(value.strip())
    return result


def list_of(value: Any) -> list[str]:
    """Список строк из строки/списка. Строка режется по переводам строк и ';'."""
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        result: list[str] = []
        for item in value:
            result.extend(list_of(item))
        return _dedupe(result)
    text = text_of(value)
    if not text:
        return []
    parts = re.split(r"[\r\n;]+", text)
    cleaned = []
    for part in parts:
        part = re.sub(r"^\s*(?:[-*•]+|\d+[.)])\s*", "", part).strip()
        if part:
            cleaned.append(part)
    return _dedupe(cleaned)


def normalize_priority(value: Any) -> str:
    text = text_of(value).lower()
    if not text:
        return "medium"
    for alias, priority in _PRIORITY_ALIASES.items():
        if alias in text:
            return priority
    return "medium"


def normalize_type(value: Any) -> str:
    text = text_of(value).lower()
    for alias, type_name in _TYPE_ALIASES.items():
        if alias in text:
            return type_name
    return "other"


def normalize_title(value: Any) -> str:
    """Ключ для сравнения/дедупликации названий."""
    text = text_of(value).lower().replace("ё", "е")
    text = re.sub(r"[^0-9a-zа-я\s]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()[:240]


def strip_step_number(value: str) -> str:
    return re.sub(r"^\s*(?:шаг|step)?\s*\d+\s*[.)\-:]?\s*", "", value, flags=re.IGNORECASE).strip()


def _first(data: dict[str, Any], *keys: str) -> Any:
    for key in keys:
        if key in data and data[key] not in (None, "", [], {}):
            return data[key]
    return None


# --------------------------------------------------------------------------- #
#  Чек-лист
# --------------------------------------------------------------------------- #


@dataclass
class ChecklistItem:
    title: str = ""
    area: str = ""
    type: str = "other"
    priority: str = "medium"
    requirements: list[str] = field(default_factory=list)
    preconditions: str = ""
    checks: list[str] = field(default_factory=list)
    expected: str = ""
    notes: str = ""
    id: str = ""

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ChecklistItem":
        if not isinstance(data, dict):
            raise ModelError(f"Пункт чек-листа должен быть объектом, получено: {type(data).__name__}")

        title = text_of(_first(data, "title", "name", "check", "check_title", "проверка", "название"))
        if not title:
            raise ModelError("В пункте чек-листа отсутствует title")

        checks = list_of(_first(data, "checks", "steps", "actions", "проверки", "действия"))
        if not checks:
            checks = [title]

        item = cls(
            title=title,
            area=text_of(_first(data, "area", "module", "section", "область", "модуль", "раздел")),
            type=normalize_type(_first(data, "type", "check_type", "тип")),
            priority=normalize_priority(_first(data, "priority", "приоритет")),
            requirements=list_of(_first(data, "requirements", "requirement", "source", "требования")),
            preconditions=text_of(
                _first(data, "preconditions", "precondition", "предусловия", "предусловие")
            ),
            checks=checks,
            expected=text_of(
                _first(data, "expected", "expected_result", "expectedResult", "ожидаемый_результат")
            ),
            notes=text_of(_first(data, "notes", "note", "comment", "примечания")),
            id=text_of(_first(data, "id", "checklist_id")),
        )
        if not item.area:
            item.area = "Общее"
        if not item.expected:
            item.expected = "Поведение соответствует требованиям"
        return item

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @property
    def query_text(self) -> str:
        return " ".join(
            part
            for part in (
                self.title,
                self.area,
                " ".join(self.checks),
                self.expected,
                self.preconditions,
            )
            if part
        )


# --------------------------------------------------------------------------- #
#  Шаги и тест-кейсы
# --------------------------------------------------------------------------- #


@dataclass
class CaseStep:
    action: str = ""
    expected_result: str = ""
    test_data: str = ""

    @classmethod
    def from_dict(cls, data: Any) -> "CaseStep":
        if isinstance(data, str):
            return cls(action=strip_step_number(data))
        if not isinstance(data, dict):
            raise ModelError("Шаг должен быть объектом или строкой")

        action = text_of(
            _first(data, "action", "step", "steps", "description", "действие", "шаг")
        )
        expected = text_of(
            _first(
                data,
                "expected_result",
                "expectedResult",
                "expected",
                "result",
                "ожидаемый_результат",
                "ожидаемый",
            )
        )
        if not action and expected:
            action = expected
            expected = ""
        if not action:
            raise ModelError("В шаге отсутствует действие (action)")
        return cls(
            action=strip_step_number(action) if re.match(r"^\s*\d+[.)]", action) else action,
            expected_result=expected,
            test_data=text_of(_first(data, "test_data", "testData", "data", "тестовые_данные")),
        )

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class GeneratedCase:
    checklist_id: str = ""
    title: str = ""
    description: str = ""
    area: str = ""
    priority: str = "medium"
    preconditions: str = ""
    postconditions: str = ""
    steps: list[CaseStep] = field(default_factory=list)
    test_data: str = ""
    automation_status: str = "manual"
    source_cases: list[str] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "GeneratedCase":
        if not isinstance(data, dict):
            raise ModelError("Тест-кейс должен быть объектом")

        title = text_of(_first(data, "title", "name", "название"))
        if not title:
            raise ModelError("В тест-кейсе отсутствует title")

        raw_steps = _first(data, "steps", "step", "test_steps", "шаги")
        steps: list[CaseStep] = []
        if isinstance(raw_steps, str):
            for line in raw_steps.splitlines():
                line = line.strip()
                if not line:
                    continue
                action, expected = line, ""
                for separator in ("->", "→", "|"):
                    if separator in line:
                        action, expected = (part.strip() for part in line.split(separator, 1))
                        break
                steps.append(CaseStep(action=strip_step_number(action), expected_result=expected))
        elif isinstance(raw_steps, (list, tuple)):
            for raw_step in raw_steps:
                try:
                    steps.append(CaseStep.from_dict(raw_step))
                except ModelError:
                    continue

        if not steps:
            raise ModelError(f"В тест-кейсе «{title}» нет шагов")

        case = cls(
            checklist_id=text_of(_first(data, "checklist_id", "checklistId", "check_id")),
            title=title,
            description=text_of(_first(data, "description", "objective", "summary", "описание")),
            area=text_of(_first(data, "area", "path", "folder", "section", "область")),
            priority=normalize_priority(_first(data, "priority", "приоритет")),
            preconditions=text_of(
                _first(data, "preconditions", "precondition", "предусловия")
            ),
            postconditions=text_of(
                _first(data, "postconditions", "postcondition", "постусловия")
            ),
            steps=steps,
            test_data=text_of(_first(data, "test_data", "testData", "тестовые_данные")),
            automation_status=text_of(
                _first(data, "automation_status", "automationStatus", "автоматизация")
            )
            or "manual",
            source_cases=[str(v) for v in _as_iterable(_first(data, "source_cases", "sources"))],
        )
        return case

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["steps"] = [step.to_dict() for step in self.steps]
        return payload

    @property
    def dedupe_key(self) -> str:
        return normalize_title(self.title)


def _as_iterable(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return list(value)
    return [value]


@dataclass
class ExportedCase:
    """Тест-кейс из выгрузки DoQA (нужен как эталон стиля и источник шагов)."""

    id: str = ""
    number: str = ""
    path: str = ""
    title: str = ""
    priority: str = "medium"
    status: str = ""
    description: str = ""
    parameters: str = ""
    preconditions: str = ""
    postconditions: str = ""
    test_data: str = ""
    automation_status: str = ""
    attributes: str = ""
    steps: list[CaseStep] = field(default_factory=list)
    row: int = 0

    @property
    def search_text(self) -> str:
        parts = [
            self.title,
            self.title,
            self.path,
            self.description,
            self.preconditions,
            self.test_data,
        ]
        for step in self.steps:
            parts.append(step.action)
            parts.append(step.expected_result)
        return "\n".join(part for part in parts if part)

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["steps"] = [step.to_dict() for step in self.steps]
        return payload
