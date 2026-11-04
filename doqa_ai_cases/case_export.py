from __future__ import annotations

import logging
import math
import re
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .models import CaseStep, ExportedCase, GeneratedCase, strip_step_number


logger = logging.getLogger(__name__)


class CaseExportError(RuntimeError):
    """Ошибка чтения/записи файла тест-кейсов."""


# --------------------------------------------------------------------------- #
#  Колонки
# --------------------------------------------------------------------------- #

CANONICAL_COLUMNS = [
    "number",
    "id",
    "path",
    "title",
    "creator",
    "creationDate",
    "priority",
    "status",
    "keys",
    "description",
    "parameters",
    "preconditions",
    "postconditions",
    "attachments",
    "step",
    "expectedResult",
    "testData",
    "automationStatus",
    "autotest",
    "attributes",
]

_ALIASES: dict[str, list[str]] = {
    "number": ["number", "no", "num", "номер", "пп", "номерпп", "номерстроки"],
    "id": ["id", "ид", "caseid", "тесткейсид", "идтесткейса", "идентификатор", "idтесткейса"],
    "path": ["path", "folder", "folderpath", "путь", "папка", "раздел", "каталог"],
    "title": [
        "title",
        "name",
        "casetitle",
        "testcase",
        "testcasename",
        "название",
        "заголовок",
        "тема",
        "наименование",
        "сценарий",
    ],
    "creator": ["creator", "author", "автор", "создатель", "пользователь"],
    "creationDate": ["creationdate", "createdat", "created", "датасоздания", "создан", "дата"],
    "priority": ["priority", "приоритет", "важность"],
    "status": ["status", "статус", "состояние"],
    "keys": ["keys", "key", "tags", "теги", "ключи", "метки"],
    "description": ["description", "описание", "краткоеописание"],
    "parameters": ["parameters", "parameter", "параметры", "параметр"],
    "preconditions": ["preconditions", "precondition", "предусловия", "предусловие"],
    "postconditions": ["postconditions", "postcondition", "постусловия", "постусловие"],
    "attachments": ["attachments", "attachment", "вложения", "вложение", "файлы"],
    "testData": ["testdata", "тестовыеданные", "данные", "тестданные"],
    "automationStatus": ["automationstatus", "статусавтоматизации", "автоматизация"],
    "autotest": ["autotest", "автотест", "ссылканаавтотест"],
    "attributes": ["attributes", "attribute", "атрибуты", "атрибут"],
    "expectedResult": ["expectedresult", "ожидаемыйрезультат", "ожидаемыйрезультатшага"],
}

_STEP_RE = re.compile(r"^(?:step|steps|шаг|шаги|действие|действия)(\d*)$")
_EXPECTED_RE = re.compile(
    r"^(?:expectedresult|expectedresults|expected|result|ожидаемыйрезультат|"
    r"ожидаемыйрезультаты|ожидаемый|результат|ожидаемые)(\d*)$"
)


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower().replace("ё", "е")
    return re.sub(r"[^0-9a-zа-я]+", "", text)


def _alias_lookup() -> dict[str, str]:
    lookup: dict[str, str] = {}
    for key, aliases in _ALIASES.items():
        for alias in aliases:
            lookup[normalize_header(alias)] = key
    return lookup


_ALIAS_TO_KEY = _alias_lookup()

# ключ -> может повторяться в шапке несколько раз
_MULTI_KEYS = {"step", "expectedResult"}


@dataclass
class ColumnSpec:
    index: int
    header: str
    key: str = ""
    step_no: "int | None" = None


@dataclass
class ExportedBundle:
    path: Path
    format: str
    cases: list[ExportedCase]
    headers: list[str]
    columns: list[ColumnSpec]
    header_row: int
    sheet_name: str
    freeze_panes: str = ""


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        if value.time() == datetime.min.time():
            return value.date().isoformat()
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


# --------------------------------------------------------------------------- #
#  Разбор шапки
# --------------------------------------------------------------------------- #


def _classify_header(raw: Any) -> "tuple[str, int | None] | None":
    raw_text = str(raw if raw is not None else "").strip()
    if not raw_text:
        return None
    if raw_text.lower() in {"#", "№", "n", "no", "no."}:
        return ("number", None)
    normalized = normalize_header(raw_text)
    if not normalized:
        return None

    match = _STEP_RE.match(normalized)
    if match:
        return ("step", int(match.group(1)) if match.group(1) else None)

    match = _EXPECTED_RE.match(normalized)
    if match:
        return ("expectedResult", int(match.group(1)) if match.group(1) else None)

    key = _ALIAS_TO_KEY.get(normalized)
    if key:
        return (key, None)
    return None


def _build_columns(header_values: list[str]) -> list[ColumnSpec]:
    columns: list[ColumnSpec] = []
    for index, raw in enumerate(header_values, start=1):
        classified = _classify_header(raw)
        if not classified:
            continue
        key, step_no = classified
        columns.append(ColumnSpec(index=index, header=str(raw).strip(), key=key, step_no=step_no))

    for key in ("step", "expectedResult"):
        specs = [spec for spec in columns if spec.key == key]
        explicit = [spec.step_no for spec in specs if spec.step_no is not None]
        next_number = max(explicit, default=0)
        for spec in specs:
            if spec.step_no is None:
                next_number += 1
                spec.step_no = next_number
    return columns


def _detect_header(ws) -> "tuple[int, list[str], list[ColumnSpec]] | None":  # noqa: ANN001
    best: "tuple[int, int, list[str], list[ColumnSpec]] | None" = None
    max_row = min(getattr(ws, "max_row", 0) or 0, 40)
    max_column = min(getattr(ws, "max_column", 0) or 0, 500)
    for row_index in range(1, max_row + 1):
        values = [_cell_text(ws.cell(row=row_index, column=col).value) for col in range(1, max_column + 1)]
        while values and not values[-1]:
            values.pop()
        if len([value for value in values if value]) < 2:
            continue
        columns = _build_columns(values)
        keys = {column.key for column in columns}
        if len([column for column in columns if column.key]) < 3:
            continue
        score = len([column for column in columns if column.key])
        if "title" in keys:
            score += 5
        if "step" in keys:
            score += 3
        if best is None or score > best[0]:
            best = (score, row_index, values, columns)
    if best is None:
        return None
    _, row_index, values, columns = best
    return row_index, values, columns


# --------------------------------------------------------------------------- #
#  Разбор строк
# --------------------------------------------------------------------------- #


def _split_lines(text: str) -> list[str]:
    return [line.strip() for line in str(text or "").splitlines() if line.strip()]


def _steps_from_single_pair(action_text: str, expected_text: str) -> list[CaseStep]:
    action_lines = _split_lines(action_text)
    expected_lines = _split_lines(expected_text)

    if not action_lines and not expected_lines:
        return []
    if len(action_lines) <= 1 and len(expected_lines) <= 1:
        return [
            CaseStep(
                action=(action_lines[0] if action_lines else ""),
                expected_result=(expected_lines[0] if expected_lines else ""),
            )
        ]
    if len(action_lines) == len(expected_lines) and len(action_lines) > 1:
        return [
            CaseStep(
                action=strip_step_number(action),
                expected_result=strip_step_number(expected),
            )
            for action, expected in zip(action_lines, expected_lines)
        ]
    steps: list[CaseStep] = []
    for position, action in enumerate(action_lines):
        expected = strip_step_number(expected_text) if position == 0 and len(action_lines) > 1 else ""
        steps.append(CaseStep(action=strip_step_number(action), expected_result=expected))
    if not steps and expected_lines:
        steps.append(CaseStep(action="", expected_result="\n".join(expected_lines)))
    return steps


def _row_to_case(
    values: list[str],
    columns: list[ColumnSpec],
    row_index: int,
) -> ExportedCase:
    data: dict[str, str] = {}
    for column in columns:
        if column.key in _MULTI_KEYS:
            continue
        if column.index - 1 >= len(values):
            continue
        value = values[column.index - 1]
        if value and column.key not in data:
            data[column.key] = value

    step_specs = [column for column in columns if column.key == "step"]
    expected_specs = [column for column in columns if column.key == "expectedResult"]

    if len(step_specs) == 1:
        action_text = values[step_specs[0].index - 1] if step_specs[0].index - 1 < len(values) else ""
        expected_text = (
            values[expected_specs[0].index - 1]
            if expected_specs and expected_specs[0].index - 1 < len(values)
            else ""
        )
        steps = _steps_from_single_pair(action_text, expected_text)
    else:
        actions: dict[int, str] = {}
        expecteds: dict[int, str] = {}
        for column in step_specs:
            position = column.index - 1
            if position < len(values) and values[position]:
                actions[column.step_no or 0] = values[position]
        for column in expected_specs:
            position = column.index - 1
            if position < len(values) and values[position]:
                expecteds[column.step_no or 0] = values[position]
        steps = []
        for number in sorted(set(actions) | set(expecteds)):
            action = actions.get(number, "")
            expected = expecteds.get(number, "")
            if action or expected:
                steps.append(CaseStep(action=action, expected_result=expected))

    return ExportedCase(
        id=data.get("id", ""),
        number=data.get("number", ""),
        path=data.get("path", ""),
        title=data.get("title", ""),
        priority=(data.get("priority", "") or "medium"),
        status=data.get("status", ""),
        description=data.get("description", ""),
        parameters=data.get("parameters", ""),
        preconditions=data.get("preconditions", ""),
        postconditions=data.get("postconditions", ""),
        test_data=data.get("testData", ""),
        automation_status=data.get("automationStatus", ""),
        attributes=data.get("attributes", ""),
        steps=steps,
        row=row_index,
    )


def load_export(path: "Path | str") -> ExportedBundle:
    """Читает xlsx-выгрузку DoQA и возвращает кейсы + структуру шапки."""
    export_path = Path(path)
    if not export_path.is_file():
        raise CaseExportError(f"Файл выгрузки не найден: {export_path}")
    if export_path.suffix.lower() == ".pdf":
        raise CaseExportError(
            "PDF-выгрузка не поддерживается как источник кейсов: нужен xlsx "
            "(format=xlsx), иначе нельзя разобрать шаги."
        )

    workbook = load_workbook(filename=str(export_path), data_only=True, read_only=False)
    try:
        worksheet = workbook.worksheets[0]
        header = _detect_header(worksheet)
        if header is None:
            sample = []
            for row in worksheet.iter_rows(min_row=1, max_row=min(5, worksheet.max_row), values_only=True):
                sample.append(" | ".join(str(value) for value in row if value is not None))
            raise CaseExportError(
                "Не удалось найти строку заголовков в выгрузке DoQA. "
                "Проверьте, что экспортирован xlsx и в нём есть колонки "
                "«Название/Title», «Шаг/Step», «Ожидаемый результат/Expected result».\n"
                "Первые строки файла:\n" + "\n".join(sample)
            )

        header_row, headers, columns = header
        keys = {column.key for column in columns}
        if "title" not in keys:
            recognized = ", ".join(
                f"«{column.header}» -> {column.key}" for column in columns
            )
            raise CaseExportError(
                "В шапке выгрузки не найдена колонка названия тест-кейса. "
                "Ожидается «Название», «Title» или «Case title». "
                f"Распознанные колонки: {recognized or 'нет'}."
            )
        if "step" not in keys:
            logger.warning(
                "В выгрузке нет колонки шагов: эталонные кейсы будут без шагов"
            )

        header_keys = {normalize_header(column.header) for column in columns}
        last_column = max(
            [len(headers)] + [column.index for column in columns] + [1]
        )
        cases: list[ExportedCase] = []
        for row_index in range(header_row + 1, worksheet.max_row + 1):
            values = [
                _cell_text(worksheet.cell(row=row_index, column=col).value)
                for col in range(1, last_column + 1)
            ]
            if not any(values):
                continue
            repeated_header = sum(
                1 for value in values if value and normalize_header(value) in header_keys
            )
            if repeated_header >= 3:
                continue
            case = _row_to_case(values, columns, row_index)
            if case.title or case.steps or case.description:
                cases.append(case)

        return ExportedBundle(
            path=export_path,
            format=export_path.suffix.lower().lstrip("."),
            cases=cases,
            headers=headers,
            columns=columns,
            header_row=header_row,
            sheet_name=worksheet.title,
            freeze_panes=worksheet.freeze_panes or "",
        )
    finally:
        workbook.close()


# --------------------------------------------------------------------------- #
#  Запись итоговых кейсов в формате выгрузки
# --------------------------------------------------------------------------- #


@dataclass
class _HeaderSlot:
    key: str
    step_no: "int | None"
    header: str
    source_index: "int | None" = None
    width: "float | None" = None
    style: Any = None


_DEFAULT_WIDTHS = {
    "number": 8,
    "id": 10,
    "path": 28,
    "title": 50,
    "creator": 18,
    "creationDate": 14,
    "priority": 10,
    "status": 10,
    "keys": 16,
    "description": 60,
    "parameters": 30,
    "preconditions": 45,
    "postconditions": 45,
    "attachments": 20,
    "step": 55,
    "expectedResult": 55,
    "testData": 30,
    "automationStatus": 16,
    "autotest": 20,
    "attributes": 25,
}

_THIN = Side(style="thin", color="B7C3D6")
_DEFAULT_HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
_DEFAULT_HEADER_FILL = PatternFill("solid", fgColor="305496")
_DEFAULT_HEADER_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _default_slots() -> list[_HeaderSlot]:
    return [
        _HeaderSlot(key=key, step_no=1 if key == "step" else (1 if key == "expectedResult" else None),
                    header=key)
        for key in CANONICAL_COLUMNS
    ]


def _slots_from_template(template: ExportedBundle) -> list[_HeaderSlot]:
    workbook = load_workbook(filename=str(template.path), data_only=False, read_only=False)
    try:
        worksheet = workbook.worksheets[0]
        spec_by_index = {column.index: column for column in template.columns}
        slots: list[_HeaderSlot] = []
        for column_index, header in enumerate(template.headers, start=1):
            spec = spec_by_index.get(column_index)
            slot = _HeaderSlot(
                key=spec.key if spec else "",
                step_no=spec.step_no if spec else None,
                header=header,
                source_index=column_index,
            )
            letter = get_column_letter(column_index)
            dimension = worksheet.column_dimensions.get(letter)
            slot.width = dimension.width if dimension and dimension.width else None
            try:
                slot.style = copy(worksheet.cell(row=template.header_row, column=column_index)._style)
            except Exception:  # noqa: BLE001
                slot.style = None
            slots.append(slot)
        return slots
    finally:
        workbook.close()


def _last_pair_position(slots: list[_HeaderSlot]) -> int:
    positions = [index for index, slot in enumerate(slots) if slot.key in _MULTI_KEYS]
    return (max(positions) + 1) if positions else len(slots)


def _expand_slots(slots: list[_HeaderSlot], max_steps: int) -> None:
    step_slots = [slot for slot in slots if slot.key == "step" and slot.step_no]
    if not step_slots:
        return

    existing = max((slot.step_no or 0) for slot in step_slots)
    if max_steps <= existing:
        return

    last_step = max(step_slots, key=lambda slot: slot.step_no or 0)
    expected_slots = [slot for slot in slots if slot.key == "expectedResult" and slot.step_no]
    last_expected = max(expected_slots, key=lambda slot: slot.step_no or 0) if expected_slots else None

    def base_name(header: str, fallback: str) -> str:
        stripped = re.sub(r"\s*\d+\s*$", "", header).strip()
        return stripped or fallback

    step_base = base_name(last_step.header, "step")
    expected_base = base_name(last_expected.header if last_expected else "", "expectedResult")

    insert_at = _last_pair_position(slots)
    for number in range(existing + 1, max_steps + 1):
        step_slot = _HeaderSlot(
            key="step",
            step_no=number,
            header=f"{step_base} {number}",
            width=last_step.width,
            style=last_step.style,
        )
        expected_slot = _HeaderSlot(
            key="expectedResult",
            step_no=number,
            header=f"{expected_base} {number}",
            width=last_expected.width if last_expected else None,
            style=last_expected.style if last_expected else None,
        )
        slots[insert_at:insert_at] = [step_slot, expected_slot]
        insert_at += 2


def _numbered_steps(slots: list[_HeaderSlot]) -> bool:
    step_slots = [slot for slot in slots if slot.key == "step"]
    if len(step_slots) > 1:
        return True
    return bool(step_slots and re.search(r"\d", step_slots[0].header))


def _case_value_for(
    slot: _HeaderSlot,
    case: GeneratedCase,
    row_number: int,
    numbered: bool,
    today: str,
) -> Any:
    key = slot.key
    if key == "number":
        return row_number
    if key == "id":
        return ""
    if key == "path":
        return case.area
    if key == "title":
        return case.title
    if key == "creator":
        return ""
    if key == "creationDate":
        return today
    if key == "priority":
        return case.priority
    if key == "status":
        return "draft"
    if key == "keys":
        return ""
    if key == "description":
        return case.description
    if key == "parameters":
        return ""
    if key == "preconditions":
        return case.preconditions
    if key == "postconditions":
        return case.postconditions
    if key == "attachments":
        return ""
    if key == "step":
        if numbered:
            number = slot.step_no or 1
            if 1 <= number <= len(case.steps):
                return case.steps[number - 1].action
            return ""
        return "\n".join(
            f"{index}. {step.action}"
            for index, step in enumerate(case.steps, start=1)
            if step.action
        )
    if key == "expectedResult":
        if numbered:
            number = slot.step_no or 1
            if 1 <= number <= len(case.steps):
                return case.steps[number - 1].expected_result
            return ""
        return "\n".join(
            f"{index}. {step.expected_result}"
            for index, step in enumerate(case.steps, start=1)
            if step.expected_result
        )
    if key == "testData":
        if case.test_data:
            return case.test_data
        data_lines = [
            f"{index}. {step.test_data}"
            for index, step in enumerate(case.steps, start=1)
            if step.test_data
        ]
        return "\n".join(data_lines)
    if key == "automationStatus":
        return case.automation_status or "manual"
    if key == "autotest":
        return ""
    if key == "attributes":
        return ""
    return ""


def _row_height(values: list[Any], widths: list[float]) -> float:
    lines = 1
    for value, width in zip(values, widths):
        if not isinstance(value, str) or not value:
            continue
        height_lines = value.count("\n") + 1
        wrapped_lines = math.ceil(len(value) / max(12.0, width * 1.05))
        lines = max(lines, height_lines, wrapped_lines)
    return min(15.0 * lines + 3.0, 260.0)


def write_cases_xlsx(
    cases: list[GeneratedCase],
    out_path: "Path | str",
    template: "ExportedBundle | None" = None,
    sheet_name: "str | None" = None,
) -> Path:
    """Пишет кейсы в xlsx, максимально повторяя структуру выгрузки DoQA."""
    output = Path(out_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    slots = _slots_from_template(template) if template is not None else _default_slots()
    if not slots:
        slots = _default_slots()

    max_steps = max((len(case.steps) for case in cases), default=1)
    numbered = _numbered_steps(slots)
    if numbered:
        _expand_slots(slots, max_steps)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = (sheet_name or (template.sheet_name if template else "Test cases"))[:31]

    for position, slot in enumerate(slots, start=1):
        cell = worksheet.cell(row=1, column=position, value=slot.header)
        if slot.style is not None:
            try:
                cell._style = copy(slot.style)
            except Exception:  # noqa: BLE001
                cell.font = _DEFAULT_HEADER_FONT
                cell.fill = _DEFAULT_HEADER_FILL
                cell.border = _DEFAULT_HEADER_BORDER
                cell.alignment = Alignment(wrap_text=True, vertical="center")
        else:
            cell.font = _DEFAULT_HEADER_FONT
            cell.fill = _DEFAULT_HEADER_FILL
            cell.border = _DEFAULT_HEADER_BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        width = slot.width or _DEFAULT_WIDTHS.get(slot.key, 20)
        worksheet.column_dimensions[get_column_letter(position)].width = width

    worksheet.row_dimensions[1].height = 32
    today = date.today().isoformat()
    widths = [slot.width or _DEFAULT_WIDTHS.get(slot.key, 20) for slot in slots]

    for row_offset, case in enumerate(cases, start=1):
        row_index = row_offset + 1
        row_values: list[Any] = []
        for position, slot in enumerate(slots, start=1):
            value = _case_value_for(slot, case, row_offset, numbered, today)
            row_values.append(value)
            cell = worksheet.cell(row=row_index, column=position, value=value)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        worksheet.row_dimensions[row_index].height = _row_height(row_values, widths)

    freeze = template.freeze_panes if template and template.freeze_panes else "A2"
    worksheet.freeze_panes = freeze
    workbook.save(str(output))
    return output
