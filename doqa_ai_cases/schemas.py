from __future__ import annotations


CHECKLIST_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["items"],
    "properties": {
        "items": {
            "type": "array",
            "description": "Пункты чек-листа. Пустой список, если во фрагменте нет проверяемых требований.",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": [
                    "area",
                    "title",
                    "type",
                    "priority",
                    "requirements",
                    "preconditions",
                    "checks",
                    "expected",
                    "notes",
                ],
                "properties": {
                    "area": {
                        "type": "string",
                        "description": "Функциональный блок или модуль, например «Авторизация».",
                    },
                    "title": {
                        "type": "string",
                        "description": "Короткое название атомарной проверки без слов «проверить, что».",
                    },
                    "type": {
                        "type": "string",
                        "enum": [
                            "positive",
                            "negative",
                            "boundary",
                            "integration",
                            "security",
                            "ui",
                            "performance",
                            "other",
                        ],
                    },
                    "priority": {"type": "string", "enum": ["high", "medium", "low"]},
                    "requirements": {
                        "type": "array",
                        "description": "Ссылки на источники: путь к документу и раздел/цитата.",
                        "items": {"type": "string"},
                    },
                    "preconditions": {
                        "type": "string",
                        "description": "Состояние системы и данные до начала проверки.",
                    },
                    "checks": {
                        "type": "array",
                        "description": "1–5 конкретных действий или условий проверки.",
                        "items": {"type": "string"},
                    },
                    "expected": {
                        "type": "string",
                        "description": "Однозначный проверяемый ожидаемый результат.",
                    },
                    "notes": {
                        "type": "string",
                        "description": "Неоднозначности, вопросы к требованиям, тестовые данные.",
                    },
                },
            },
        }
    },
}


CASES_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["cases"],
    "properties": {
        "cases": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": [
                    "checklist_id",
                    "title",
                    "description",
                    "area",
                    "priority",
                    "preconditions",
                    "postconditions",
                    "steps",
                    "test_data",
                    "automation_status",
                    "source_cases",
                ],
                "properties": {
                    "checklist_id": {
                        "type": "string",
                        "description": "ID пункта чек-листа строго из входных данных, например CHK-004.",
                    },
                    "title": {"type": "string"},
                    "description": {
                        "type": "string",
                        "description": "1–3 предложения: что и зачем проверяем, со ссылкой на требование.",
                    },
                    "area": {
                        "type": "string",
                        "description": "Функциональный блок/папка, как в эталонных кейсах.",
                    },
                    "priority": {"type": "string", "enum": ["high", "medium", "low"]},
                    "preconditions": {"type": "string"},
                    "postconditions": {"type": "string"},
                    "steps": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "required": ["action", "expected_result", "test_data"],
                            "properties": {
                                "action": {"type": "string"},
                                "expected_result": {"type": "string"},
                                "test_data": {"type": "string"},
                            },
                        },
                    },
                    "test_data": {
                        "type": "string",
                        "description": "Общие тестовые данные кейса; пустая строка, если не нужны.",
                    },
                    "automation_status": {
                        "type": "string",
                        "enum": ["manual", "planned", "automated"],
                    },
                    "source_cases": {
                        "type": "array",
                        "description": "ID эталонных кейсов, на которые опирались.",
                        "items": {"type": "string"},
                    },
                },
            },
        }
    },
}
