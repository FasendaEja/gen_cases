from __future__ import annotations

import json
import logging
import re
import time
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable

from .case_export import CaseExportError, ExportedBundle, load_export, write_cases_xlsx
from .doqa_client import DoQAClient
from .models import (
    ChecklistItem,
    ExportedCase,
    GeneratedCase,
    ModelError,
    normalize_title,
)
from .prompts import PromptLibrary
from .requirements_io import RequirementChunk, RequirementDoc
from .retrieval import CaseIndex, TextIndex
from .schemas import CASES_SCHEMA, CHECKLIST_SCHEMA


class PipelineError(RuntimeError):
    """Ошибка выполнения конвейера."""


def _now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _unique(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        key = (value or "").strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(value.strip())
    return result


def _truncate(text: str, limit: int) -> str:
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 1)].rstrip() + "…"


@dataclass
class RunPaths:
    root: Path

    def __post_init__(self) -> None:
        self.root = Path(self.root)
        self.ensure()

    def ensure(self) -> None:
        for sub in ("doqa", "checklist", "cases", "prompts", "logs"):
            (self.root / sub).mkdir(parents=True, exist_ok=True)

    @property
    def doqa(self) -> Path:
        return self.root / "doqa"

    @property
    def checklist(self) -> Path:
        return self.root / "checklist"

    @property
    def cases(self) -> Path:
        return self.root / "cases"

    @property
    def prompts(self) -> Path:
        return self.root / "prompts"

    @property
    def logs(self) -> Path:
        return self.root / "logs"


@dataclass
class Coverage:
    total: int = 0
    covered: dict[str, int] = field(default_factory=dict)
    missing: list[str] = field(default_factory=list)

    @property
    def covered_count(self) -> int:
        return sum(1 for item_id in self.covered if item_id)

    @property
    def percent(self) -> float:
        if not self.total:
            return 0.0
        return round(100.0 * self.covered_count / self.total, 1)

    def as_dict(self) -> dict[str, Any]:
        return {
            "checklist_total": self.total,
            "checklist_covered": self.covered_count,
            "percent": self.percent,
            "missing": self.missing,
            "cases_by_checklist": self.covered,
        }


class Pipeline:
    """Оркестратор: выгрузка DoQA -> чек-лист -> тест-кейсы."""

    def __init__(
        self,
        settings,
        llm,
        paths: RunPaths,
        logger: "logging.Logger | None" = None,
        prompts: "PromptLibrary | None" = None,
    ) -> None:
        self.settings = settings
        self.llm = llm
        self.paths = paths
        self.log = logger or logging.getLogger(__name__)
        self.prompts = prompts or PromptLibrary(settings.prompts_dir)
        self.warnings: list[str] = []
        self.timings: dict[str, float] = {}
        self.export_bundle: "ExportedBundle | None" = None
        self.raw_export_path: "Path | None" = None
        self.docs: list[RequirementDoc] = []
        self.chunks: list[RequirementChunk] = []
        self.items: list[ChecklistItem] = []
        self.cases: list[GeneratedCase] = []
        self.space_id: "int | None" = None
        self._prompt_counter = 0

    # ------------------------------------------------------------------ утилиты
    def warn(self, message: str) -> None:
        self.warnings.append(message)
        self.log.warning(message)

    def _write_json(self, path: Path, payload: Any) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _save_prompt(self, prompt_name: str, system: str, user: str) -> None:
        if not self.settings.save_prompts:
            return
        self._prompt_counter += 1
        safe_name = re.sub(r"[^0-9a-zA-Zа-яА-Я_.-]+", "_", prompt_name)[:80]
        path = self.paths.prompts / f"{self._prompt_counter:03d}_{safe_name}.md"
        path.write_text(
            f"<!-- {prompt_name} | model: {self.settings.openai_model} -->\n\n"
            f"## SYSTEM\n\n{system}\n\n## USER\n\n{user}\n",
            encoding="utf-8",
        )

    # --------------------------------------------------------------- выгрузка
    def run_export(self, space_id: int, filters: "dict[str, Any] | None" = None) -> ExportedBundle:
        start = time.perf_counter()
        self.space_id = int(space_id)
        body: dict[str, Any] = {"spaceId": int(space_id), "format": "xlsx"}
        for key, value in (filters or {}).items():
            if value not in (None, "", [], {}):
                body[key] = value

        client = DoQAClient(self.settings, self.log)
        result = client.export_cases(body)

        raw_path = self.paths.doqa / f"cases_export{result.extension}"
        raw_path.write_bytes(result.content)
        self.raw_export_path = raw_path
        self.log.info("Выгрузка сохранена: %s", raw_path)

        bundle: "ExportedBundle | None" = None
        if result.extension == ".xlsx":
            if not zipfile.is_zipfile(raw_path):
                raise CaseExportError(
                    "DoQA вернул файл с расширением .xlsx, но это не ZIP-контейнер. "
                    "Проверьте DOQA_BASE_URL и права доступа."
                )
            bundle = load_export(raw_path)
            self._dump_normalized_cases(bundle)
            if not bundle.cases:
                self.warn("В выгрузке DoQA не найдено ни одного тест-кейса — генерация будет без эталонов")
        else:
            self.warn(
                f"DoQA вернул формат {result.extension}: разобрать шаги и использовать кейсы "
                "как эталон нельзя. Для полноценной работы нужен format=xlsx"
            )

        self.export_bundle = bundle
        self.timings["export_sec"] = round(time.perf_counter() - start, 2)
        return bundle  # type: ignore[return-value]

    def attach_export_file(self, path: "Path | str") -> ExportedBundle:
        """Использовать уже скачанный xlsx вместо нового запроса к DoQA."""
        bundle = load_export(path)
        self._dump_normalized_cases(bundle)
        self.export_bundle = bundle
        self.raw_export_path = Path(path)
        if not bundle.cases:
            self.warn("В файле выгрузки не найдено тест-кейсов — генерация будет без эталонов")
        self.log.info("Используется готовая выгрузка: %s (%d кейсов)", path, len(bundle.cases))
        return bundle

    def _dump_normalized_cases(self, bundle: ExportedBundle) -> None:
        self._write_json(
            self.paths.doqa / "cases.normalized.json",
            {
                "source_file": str(bundle.path),
                "sheet": bundle.sheet_name,
                "count": len(bundle.cases),
                "cases": [case.to_dict() for case in bundle.cases],
            },
        )

    # --------------------------------------------------------------- чек-лист
    def build_checklist(self, chunks: list[RequirementChunk], docs: "list[RequirementDoc] | None" = None) -> list[ChecklistItem]:
        start = time.perf_counter()
        if not chunks:
            raise PipelineError("Нет фрагментов требований для построения чек-листа")
        self.chunks = list(chunks)
        self.docs = list(docs or [])

        system = self.prompts.system("checklist")
        items: list[ChecklistItem] = []
        total = len(chunks)
        for position, chunk in enumerate(chunks, start=1):
            user = self.prompts.render(
                "checklist",
                doc_path=chunk.doc_rel_path,
                chunk_id=chunk.id,
                chunk_index=chunk.index,
                chunk_total=chunk.total,
                heading=chunk.heading,
                requirements=chunk.text,
            )
            self._save_prompt(f"checklist_{chunk.id}", system, user)
            payload = self.llm.json_call(
                system=system,
                user=user,
                schema=CHECKLIST_SCHEMA,
                schema_name="checklist",
                prompt_name=f"checklist_{chunk.id}",
            )
            raw_items = payload.get("items")
            if not isinstance(raw_items, list):
                self.warn(f"{chunk.id}: модель вернула некорректный список пунктов — фрагмент пропущен")
                continue
            added = 0
            for raw in raw_items:
                try:
                    items.append(ChecklistItem.from_dict(raw))
                    added += 1
                except ModelError as exc:
                    self.warn(f"{chunk.id}: пункт чек-листа пропущен ({exc})")
            self.log.info(
                "Чек-лист: фрагмент %d/%d (%s) -> +%d пунктов, всего %d",
                position, total, chunk.id, added, len(items),
            )

        if not items:
            raise PipelineError("Модель не вернула ни одного пункта чек-листа")

        before = len(items)
        items = self._merge_similar_items(items)
        if len(items) != before:
            self.log.info("Локальная дедупликация: %d -> %d пунктов", before, len(items))

        if self.settings.checklist_consolidate and len(items) > 1:
            items = self._consolidate_checklist(items)

        self._assign_ids(items)
        self.items = items
        self._write_checklist_files(items)
        self.timings["checklist_sec"] = round(time.perf_counter() - start, 2)
        self.log.info("Чек-лист готов: %d пунктов", len(items))
        return items

    def _merge_similar_items(self, items: list[ChecklistItem]) -> list[ChecklistItem]:
        merged: list[ChecklistItem] = []
        for item in items:
            target = None
            item_key = normalize_title(item.title)
            for existing in merged:
                existing_key = normalize_title(existing.title)
                if not item_key or not existing_key:
                    continue
                if item_key == existing_key:
                    target = existing
                    break
                # Похожие формулировки объединяем только при совпадении области
                # и типа: иначе «валидный пароль» и «невалидный пароль» склеятся
                # в один пункт и потеряют покрытие.
                same_kind = existing.area == item.area and existing.type == item.type
                if same_kind and SequenceMatcher(None, item_key, existing_key).ratio() >= 0.97:
                    target = existing
                    break
            if target is None:
                merged.append(item)
                continue
            target.requirements = _unique([*target.requirements, *item.requirements])
            target.checks = _unique([*target.checks, *item.checks])
            if len(item.expected) > len(target.expected):
                target.expected = item.expected
            if len(item.preconditions) > len(target.preconditions):
                target.preconditions = item.preconditions
            if item.notes and item.notes not in target.notes:
                target.notes = " ".join(_unique([target.notes, item.notes]))
            if _priority_rank(item.priority) > _priority_rank(target.priority):
                target.priority = item.priority
        return merged

    def _consolidate_checklist(self, items: list[ChecklistItem]) -> list[ChecklistItem]:
        items_json = json.dumps([item.to_dict() for item in items], ensure_ascii=False)
        if len(items_json) > self.settings.consolidate_max_chars:
            self.warn(
                "Чек-лист слишком велик для сведения одним запросом "
                f"({len(items_json)} символов) — оставляю результат локальной дедупликации"
            )
            return items

        system = self.prompts.system("checklist_consolidate")
        user = self.prompts.render(
            "checklist_consolidate",
            item_count=len(items),
            items_json=items_json,
        )
        self._save_prompt("checklist_consolidate", system, user)
        payload = self.llm.json_call(
            system=system,
            user=user,
            schema=CHECKLIST_SCHEMA,
            schema_name="checklist_consolidate",
            prompt_name="checklist_consolidate",
        )
        raw_items = payload.get("items")
        if not isinstance(raw_items, list) or not raw_items:
            self.warn("Сведение чек-листа вернуло пустой результат — оставляю исходные пункты")
            return items

        consolidated: list[ChecklistItem] = []
        for raw in raw_items:
            try:
                consolidated.append(ChecklistItem.from_dict(raw))
            except ModelError as exc:
                self.warn(f"Сведение чек-листа: пункт пропущен ({exc})")
        if not consolidated:
            return items
        self.log.info("Сведение чек-листа: %d -> %d пунктов", len(items), len(consolidated))
        return consolidated

    @staticmethod
    def _assign_ids(items: list[ChecklistItem]) -> None:
        for index, item in enumerate(items, start=1):
            item.id = f"CHK-{index:03d}"

    def _write_checklist_files(self, items: list[ChecklistItem]) -> None:
        payload = {
            "generated_at": _now_iso(),
            "space_id": self.space_id,
            "requirements": [
                {"path": doc.rel_path, "chars": len(doc.text)} for doc in self.docs
            ],
            "count": len(items),
            "items": [item.to_dict() for item in items],
        }
        self._write_json(self.paths.checklist / "checklist.json", payload)

        lines: list[str] = ["# Чек-лист проверок", ""]
        lines.append(f"Сформирован: {payload['generated_at']}")
        if self.space_id:
            lines.append(f"Пространство DoQA: {self.space_id}")
        if self.docs:
            lines.append(f"Документов требований: {len(self.docs)}")
        lines.append(f"Пунктов: {len(items)}")
        lines.append("")

        current_area = None
        for item in items:
            if item.area != current_area:
                lines.extend([f"## {item.area}", ""])
                current_area = item.area
            lines.append(f"- [ ] **{item.id} — {item.title}**")
            lines.append(f"  - Тип: `{item.type}`; приоритет: `{item.priority}`")
            if item.requirements:
                lines.append("  - Требования: " + "; ".join(item.requirements))
            if item.preconditions:
                lines.append("  - Предусловия: " + item.preconditions.replace("\n", " "))
            if item.checks:
                lines.append("  - Проверки:")
                for check in item.checks:
                    lines.append(f"    - {check}")
            lines.append("  - Ожидаемый результат: " + item.expected.replace("\n", " "))
            if item.notes:
                lines.append("  - Примечания: " + item.notes.replace("\n", " "))
            lines.append("")

        (self.paths.checklist / "checklist.md").write_text("\n".join(lines), encoding="utf-8")

    # ------------------------------------------------------------- тест-кейсы
    def generate_cases(
        self,
        items: "list[ChecklistItem] | None" = None,
        exported_cases: "list[ExportedCase] | None" = None,
    ) -> tuple[list[GeneratedCase], Coverage]:
        start = time.perf_counter()
        checklist = list(items if items is not None else self.items)
        if not checklist:
            raise PipelineError("Нет пунктов чек-листа для генерации тест-кейсов")

        if exported_cases is None:
            exported_cases = self.export_bundle.cases if self.export_bundle else []
        self.items = checklist

        case_index = CaseIndex(exported_cases) if exported_cases else None
        chunk_index = TextIndex(self.chunks) if self.chunks else None
        if not exported_cases:
            self.warn(
                "Эталонные тест-кейсы отсутствуют: шаги будут составлены без опоры на выгрузку DoQA"
            )

        system = self.prompts.system("cases")
        library_stats = self._library_stats(exported_cases)
        batch_size = max(1, self.settings.case_batch_size)
        batches = [checklist[i:i + batch_size] for i in range(0, len(checklist), batch_size)]

        generated: list[GeneratedCase] = []
        for position, batch in enumerate(batches, start=1):
            references = self._collect_references(batch, case_index)
            context = self._requirements_context(batch, chunk_index)
            user = self.prompts.render(
                "cases",
                requirements_context=context,
                checklist_items=json.dumps(
                    [item.to_dict() for item in batch], ensure_ascii=False, indent=2
                ),
                reference_cases=self._format_references(references),
                library_stats=library_stats,
                additional_instructions="",
            )
            self._save_prompt(f"cases_batch_{position:03d}", system, user)
            payload = self.llm.json_call(
                system=system,
                user=user,
                schema=CASES_SCHEMA,
                schema_name="cases",
                prompt_name=f"cases_batch_{position:03d}",
            )
            batch_cases = self._parse_cases(payload, batch)
            generated.extend(batch_cases)
            self.log.info(
                "Тест-кейсы: пакет %d/%d (%s) -> +%d кейсов, всего %d",
                position,
                len(batches),
                ", ".join(item.id for item in batch),
                len(batch_cases),
                len(generated),
            )

        generated = self._dedupe_cases(generated)
        coverage = self._compute_coverage(checklist, generated)

        if coverage.missing and self.settings.max_retry_items:
            generated = self._retry_missing(
                checklist, coverage, generated, case_index, chunk_index, system, library_stats
            )
            generated = self._dedupe_cases(generated)
            coverage = self._compute_coverage(checklist, generated)

        if not generated:
            raise PipelineError("Модель не вернула ни одного тест-кейса")

        for case in generated:
            if not case.area:
                for item in checklist:
                    if item.id == case.checklist_id:
                        case.area = item.area
                        break

        self.cases = generated
        self._write_cases_files(generated, coverage)
        self.timings["cases_sec"] = round(time.perf_counter() - start, 2)
        self.log.info(
            "Тест-кейсы готовы: %d, покрытие чек-листа %.1f%%",
            len(generated),
            coverage.percent,
        )
        return generated, coverage

    def _collect_references(
        self,
        items: list[ChecklistItem],
        case_index: "CaseIndex | None",
    ) -> list[ExportedCase]:
        if case_index is None:
            return []
        scored: dict[tuple[str, int], tuple[ExportedCase, float]] = {}
        for item in items:
            for case, score in case_index.search(item.query_text, limit=self.settings.similar_cases_k):
                key = (case.id, case.row)
                if key not in scored or score > scored[key][1]:
                    scored[key] = (case, score)
        ordered = sorted(scored.values(), key=lambda pair: pair[1], reverse=True)
        return [case for case, _ in ordered[: self.settings.max_reference_cases]]

    def _requirements_context(
        self,
        items: list[ChecklistItem],
        chunk_index: "TextIndex | None",
    ) -> str:
        if chunk_index is None:
            return "(фрагменты требований недоступны)"
        selected: list[RequirementChunk] = []
        seen: set[str] = set()
        for item in items:
            for chunk, _score in chunk_index.search(item.query_text, limit=2):
                if chunk.id in seen:
                    continue
                seen.add(chunk.id)
                selected.append(chunk)

        if not selected:
            return "(релевантных фрагментов требований не найдено — опирайся на формулировки чек-листа)"

        parts: list[str] = []
        used = 0
        for chunk in selected:
            body = chunk.text
            header = f"### {chunk.doc_rel_path} — раздел «{chunk.heading}» ({chunk.id})"
            block = header + "\n" + body
            remaining = self.settings.requirements_context_max_chars - used
            if remaining <= 200:
                break
            if len(block) > remaining:
                block = block[:remaining].rstrip() + "\n…(фрагмент сокращён)"
            parts.append(block)
            used += len(block)
        return "\n\n".join(parts)

    @staticmethod
    def _library_stats(cases: list[ExportedCase]) -> str:
        if not cases:
            return "- Эталонные кейсы отсутствуют"
        step_counts = [len(case.steps) for case in cases if case.steps]
        average = sum(step_counts) / len(step_counts) if step_counts else 0.0
        cyrillic = sum(1 for case in cases if any("а" <= ch <= "я" for ch in case.title.lower()))
        language = "русский" if cyrillic >= len(cases) / 2 else "английский"
        return "\n".join(
            [
                f"- Кейсов в выгрузке: {len(cases)}",
                f"- Среднее число шагов: {average:.1f}",
                f"- Язык эталонов: {language}",
                "- Ориентируйся на длину шагов и детализацию эталонов, а не на общие рекомендации",
            ]
        )

    def _format_references(self, cases: list[ExportedCase]) -> str:
        if not cases:
            return "(похожих эталонных кейсов не найдено)"
        blocks: list[str] = []
        for case in cases:
            lines = [f"### [ID {case.id or case.row}] {case.title}"]
            meta = []
            if case.path:
                meta.append(f"path: {case.path}")
            if case.priority:
                meta.append(f"priority: {case.priority}")
            if case.automation_status:
                meta.append(f"automation: {case.automation_status}")
            if meta:
                lines.append("; ".join(meta))
            if case.description:
                lines.append("Описание: " + _truncate(case.description, 400))
            if case.preconditions:
                lines.append("Предусловия: " + _truncate(case.preconditions, 350))
            if case.steps:
                lines.append("Шаги:")
                for number, step in enumerate(case.steps[:10], start=1):
                    action = _truncate(step.action, 280)
                    expected = _truncate(step.expected_result, 280)
                    lines.append(f"{number}. {action}")
                    if expected:
                        lines.append(f"   => {expected}")
            if case.test_data:
                lines.append("Тестовые данные: " + _truncate(case.test_data, 300))

            block = "\n".join(lines)
            blocks.append(_truncate(block, self.settings.reference_case_max_chars))
        return "\n\n".join(blocks)

    def _parse_cases(
        self,
        payload: dict[str, Any],
        batch_items: list[ChecklistItem],
    ) -> list[GeneratedCase]:
        raw_cases = payload.get("cases")
        if not isinstance(raw_cases, list):
            self.warn("Модель вернула некорректный список тест-кейсов — пакет пропущен")
            return []

        valid_ids = {item.id for item in batch_items}
        parsed: list[GeneratedCase] = []
        for raw in raw_cases:
            try:
                case = GeneratedCase.from_dict(raw)
            except ModelError as exc:
                self.warn(f"Тест-кейс пропущен: {exc}")
                continue
            if case.checklist_id not in valid_ids:
                matched = self._match_item(case, batch_items)
                if not matched:
                    self.warn(
                        f"Тест-кейс «{case.title}» ссылается на неизвестный пункт "
                        f"«{case.checklist_id or '—'}» — пропущен"
                    )
                    continue
                case.checklist_id = matched
            parsed.append(case)
        return parsed

    @staticmethod
    def _match_item(case: GeneratedCase, items: list[ChecklistItem]) -> "str | None":
        case_title = normalize_title(case.title)
        best_id = None
        best_ratio = 0.0
        for item in items:
            ratio = SequenceMatcher(None, case_title, normalize_title(item.title)).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_id = item.id
        if best_id and best_ratio >= 0.62:
            return best_id
        description = normalize_title(case.description)
        for item in items:
            item_title = normalize_title(item.title)
            if item_title and item_title in description:
                return item.id
        return None

    @staticmethod
    def _dedupe_cases(cases: list[GeneratedCase]) -> list[GeneratedCase]:
        seen: set[str] = set()
        result: list[GeneratedCase] = []
        for case in cases:
            key = case.dedupe_key
            if not key or key in seen:
                continue
            seen.add(key)
            result.append(case)
        return result

    def _compute_coverage(
        self,
        items: list[ChecklistItem],
        cases: list[GeneratedCase],
    ) -> Coverage:
        counts: dict[str, int] = {}
        for case in cases:
            counts[case.checklist_id] = counts.get(case.checklist_id, 0) + 1
        missing = [item.id for item in items if item.id not in counts]
        coverage = Coverage(total=len(items), covered=counts, missing=missing)
        if missing:
            self.log.warning(
                "Без тест-кейсов осталось пунктов: %d (%.1f%% покрытия)",
                len(missing),
                coverage.percent,
            )
        return coverage

    def _retry_missing(
        self,
        items: list[ChecklistItem],
        coverage: Coverage,
        generated: list[GeneratedCase],
        case_index: "CaseIndex | None",
        chunk_index: "TextIndex | None",
        system: str,
        library_stats: str,
    ) -> list[GeneratedCase]:
        by_id = {item.id: item for item in items}
        retry_items = [by_id[item_id] for item_id in coverage.missing if item_id in by_id]
        retry_items = retry_items[: self.settings.max_retry_items]
        if not retry_items:
            return generated

        self.log.info("Повторная генерация для непокрытых пунктов: %d", len(retry_items))
        collected = list(generated)
        for position, item in enumerate(retry_items, start=1):
            references = self._collect_references([item], case_index)
            context = self._requirements_context([item], chunk_index)
            user = self.prompts.render(
                "cases",
                requirements_context=context,
                checklist_items=json.dumps([item.to_dict()], ensure_ascii=False, indent=2),
                reference_cases=self._format_references(references),
                library_stats=library_stats,
                additional_instructions=(
                    "ВАЖНО: для этого пункта чек-листа в первом проходе не получилось тест-кейса. "
                    "Обязательно составь хотя бы один тест-кейс именно для него."
                ),
            )
            self._save_prompt(f"cases_retry_{item.id}", system, user)
            payload = self.llm.json_call(
                system=system,
                user=user,
                schema=CASES_SCHEMA,
                schema_name="cases",
                prompt_name=f"cases_retry_{item.id}",
            )
            new_cases = self._parse_cases(payload, [item])
            if new_cases:
                collected.extend(new_cases)
                self.log.info("Повтор %d/%d: %s -> +%d", position, len(retry_items), item.id, len(new_cases))
            else:
                self.warn(f"Повторная генерация не дала кейса для {item.id}")
        return collected

    def _write_cases_files(self, cases: list[GeneratedCase], coverage: Coverage) -> None:
        xlsx_path = self.paths.cases / "cases.xlsx"
        write_cases_xlsx(cases, xlsx_path, template=self.export_bundle)
        self.log.info("Файл тест-кейсов: %s", xlsx_path)

        payload = {
            "generated_at": _now_iso(),
            "space_id": self.space_id,
            "source_export": str(self.export_bundle.path) if self.export_bundle else None,
            "count": len(cases),
            "coverage": coverage.as_dict(),
            "cases": [case.to_dict() for case in cases],
        }
        self._write_json(self.paths.cases / "cases.json", payload)

        lines = ["# Покрытие чек-листа тест-кейсами", ""]
        lines.append(f"Всего пунктов: {coverage.total}")
        lines.append(f"Покрыто: {coverage.covered_count} ({coverage.percent}%)")
        lines.append(f"Без тест-кейсов: {len(coverage.missing)}")
        lines.append("")
        if coverage.missing:
            lines.append("## Пункты без тест-кейсов")
            lines.append("")
            by_id = {item.id: item for item in self.items}
            for item_id in coverage.missing:
                item = by_id.get(item_id)
                title = item.title if item else ""
                lines.append(f"- {item_id} — {title}")
        else:
            lines.append("Все пункты чек-листа покрыты тест-кейсами.")
        (self.paths.cases / "coverage.md").write_text("\n".join(lines), encoding="utf-8")

    # ---------------------------------------------------------------- отчёт
    def report(self, coverage: "Coverage | None" = None) -> dict[str, Any]:
        usage = getattr(self.llm, "usage", None)
        return {
            "finished_at": _now_iso(),
            "space_id": self.space_id,
            "requirements_docs": len(self.docs),
            "requirements_chunks": len(self.chunks),
            "checklist_items": len(self.items),
            "generated_cases": len(self.cases),
            "coverage": coverage.as_dict() if coverage else None,
            "exported_cases": len(self.export_bundle.cases) if self.export_bundle else 0,
            "timings_sec": dict(self.timings),
            "llm_usage": usage.as_dict() if usage else None,
            "warnings": list(self.warnings),
            "files": {
                "export": str(self.raw_export_path) if self.raw_export_path else None,
                "checklist_json": str(self.paths.checklist / "checklist.json") if self.items else None,
                "checklist_md": str(self.paths.checklist / "checklist.md") if self.items else None,
                "cases_xlsx": str(self.paths.cases / "cases.xlsx") if self.cases else None,
                "cases_json": str(self.paths.cases / "cases.json") if self.cases else None,
            },
        }


def _priority_rank(priority: str) -> int:
    return {"low": 1, "medium": 2, "high": 3}.get(priority, 2)
