from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from .case_export import CaseExportError
from .config import ConfigError, PROJECT_ROOT, Settings, load_env_file
from . import __version__
from .doqa_client import DoQAError
from .llm import DryRunLLM, LLMClient, LLMError
from .models import ChecklistItem, ModelError
from .pipeline import Coverage, Pipeline, PipelineError, RunPaths
from .prompts import PromptError
from .requirements_io import RequirementsError, load_requirements


def configure_console() -> None:
    """UTF-8 в stdout/stderr: иначе Windows-консоль в cp1251 роняет --help."""
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            continue


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="doqa-ai-cases",
        description=(
            "Экспорт тест-кейсов из DoQA, генерация чек-листа по требованиям "
            "и тест-кейсов по чек-листу с помощью OpenAI."
        ),
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"doqa-ai-cases {__version__}",
        help="Показать версию и выйти",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--env-file", default=".env", help="Путь к .env (по умолчанию .env)")
    common.add_argument("--out", default=None, help="Каталог результатов (по умолчанию runs/<timestamp>)")
    common.add_argument("--dry-run", action="store_true", help="Не вызывать модель, только сохранить промпты")
    common.add_argument("--no-cache", action="store_true", help="Отключить кэш ответов модели")
    common.add_argument("--no-prompts", action="store_true", help="Не сохранять рендеры промптов")
    common.add_argument("--model", default=None, help="Переопределить OPENAI_MODEL")
    common.add_argument("--verbose", "-v", action="store_true", help="Подробный лог")
    common.add_argument(
        "--version",
        action="version",
        version=f"doqa-ai-cases {__version__}",
        help="Показать версию и выйти",
    )

    def add_filters(target: argparse.ArgumentParser) -> None:
        target.add_argument("--format", dest="export_format", default=None, choices=["xlsx", "pdf"])
        target.add_argument("--folder-id", dest="folder_ids", type=int, action="append", default=None)
        target.add_argument("--item-id", dest="item_ids", type=int, action="append", default=None)
        target.add_argument("--tag-id", dest="tag_ids", type=int, action="append", default=None)
        target.add_argument("--status", dest="statuses", action="append", default=None,
                            choices=["ready", "draft", "archived", "broken"])
        target.add_argument("--priority", dest="priorities", action="append", default=None,
                            choices=["low", "medium", "high"])
        target.add_argument("--search", default=None)
        target.add_argument("--sort", default=None,
                            choices=["id", "title", "creatorId", "createdAt", "changedAt", "priority", "status", "tag"])
        target.add_argument("--direction", default=None, choices=["asc", "desc"])
        target.add_argument("--column", dest="columns", action="append", default=None)
        target.add_argument("--filter-json", default=None,
                            help="Дополнительные фильтры метода экспорта в виде JSON-объекта")

    export_parser = subparsers.add_parser("export", parents=[common], help="Только выгрузка кейсов из DoQA")
    export_parser.add_argument("--space-id", type=int, default=None)
    add_filters(export_parser)

    checklist_parser = subparsers.add_parser(
        "checklist", parents=[common], help="Только чек-лист по требованиям"
    )
    checklist_parser.add_argument("--requirements", default=None,
                                  help="Папка с требованиями (по умолчанию REQUIREMENTS_DIR)")

    cases_parser = subparsers.add_parser(
        "cases", parents=[common], help="Только тест-кейсы по требованиям и выгрузке"
    )
    cases_parser.add_argument("--requirements", default=None,
                              help="Папка с требованиями (по умолчанию REQUIREMENTS_DIR)")
    cases_parser.add_argument("--export-file", default=None,
                              help="Ранее скачанный xlsx из DoQA (по умолчанию DOQA_EXPORT_FILE)")
    cases_parser.add_argument("--checklist-file", default=None,
                             help="Готовый checklist.json (иначе чек-лист будет сгенерирован)")

    all_parser = subparsers.add_parser(
        "all", parents=[common], help="Полный цикл: экспорт -> чек-лист -> тест-кейсы"
    )
    all_parser.add_argument("--space-id", type=int, default=None)
    all_parser.add_argument("--requirements", default=None,
                            help="Папка с требованиями (по умолчанию REQUIREMENTS_DIR)")
    all_parser.add_argument("--export-file", default=None,
                            help="Использовать готовый xlsx вместо запроса к DoQA")
    add_filters(all_parser)

    return parser


def _configure_logging(log_path: Path, verbose: bool, level_name: str = "INFO") -> logging.Logger:
    logger = logging.getLogger("doqa_ai_cases")
    level = logging.DEBUG if verbose else getattr(logging, level_name.upper(), logging.INFO)
    logger.setLevel(logging.DEBUG if verbose else level)
    for handler in list(logger.handlers):
        try:
            handler.close()
        except Exception:  # noqa: BLE001
            pass
    logger.handlers.clear()
    logger.propagate = False

    formatter = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", "%H:%M:%S")
    console = logging.StreamHandler(stream=sys.stdout)
    console.setLevel(logging.DEBUG if verbose else level)
    console.setFormatter(formatter)
    logger.addHandler(console)

    try:
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)-7s %(name)s: %(message)s")
        )
        logger.addHandler(file_handler)
    except OSError:
        pass
    return logger


def _build_filters(args: argparse.Namespace, settings: Settings) -> dict[str, Any]:
    filters: dict[str, Any] = {}
    filters["format"] = getattr(args, "export_format", None) or settings.export_format
    if getattr(args, "folder_ids", None):
        filters["folderIds"] = args.folder_ids
    if getattr(args, "item_ids", None):
        filters["itemIds"] = args.item_ids
    if getattr(args, "tag_ids", None):
        filters["tagIds"] = args.tag_ids
    if getattr(args, "statuses", None):
        filters["statuses"] = args.statuses
    if getattr(args, "priorities", None):
        filters["priorities"] = args.priorities
    if getattr(args, "search", None):
        filters["search"] = args.search
    if getattr(args, "sort", None):
        filters["sort"] = args.sort
    if getattr(args, "direction", None):
        filters["direction"] = args.direction
    if getattr(args, "columns", None):
        filters["columns"] = args.columns
    if getattr(args, "filter_json", None):
        try:
            extra = json.loads(args.filter_json)
        except json.JSONDecodeError as exc:
            raise ConfigError(f"--filter-json: некорректный JSON ({exc})") from exc
        if not isinstance(extra, dict):
            raise ConfigError("--filter-json: ожидается JSON-объект")
        filters.update(extra)
    return filters


def _resolve_run_dir(args: argparse.Namespace, settings: Settings, space_id: "int | None") -> Path:
    if args.out:
        return Path(args.out).expanduser().resolve()
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    name = f"{stamp}-space{space_id}" if space_id else stamp
    return (settings.runs_dir / name).resolve()


def _resolve_settings(args: argparse.Namespace) -> Settings:
    settings = Settings.from_env()
    if args.no_cache:
        settings.llm_cache = False
    if args.no_prompts:
        settings.save_prompts = False
    if args.model:
        settings.openai_model = args.model
    return settings


def _requirements_path(args: argparse.Namespace, settings: Settings) -> Path:
    value = getattr(args, "requirements", None) or settings.requirements_dir
    if not value:
        raise ConfigError("Укажите --requirements или REQUIREMENTS_DIR в .env")
    return Path(value).expanduser()


def _load_checklist_file(path: "Path | str") -> list[ChecklistItem]:
    payload = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    raw_items = payload.get("items") if isinstance(payload, dict) else payload
    if not isinstance(raw_items, list) or not raw_items:
        raise PipelineError(f"В файле чек-листа нет пунктов: {path}")
    items: list[ChecklistItem] = []
    for raw in raw_items:
        items.append(ChecklistItem.from_dict(raw))
    for index, item in enumerate(items, start=1):
        if not item.id:
            item.id = f"CHK-{index:03d}"
    return items


def _make_llm(args: argparse.Namespace, settings: Settings, paths: RunPaths):
    if args.dry_run:
        return DryRunLLM()
    settings.require_llm()
    return LLMClient(settings)


def _write_report(paths: RunPaths, pipeline: Pipeline, coverage: "Coverage | None") -> Path:
    report = pipeline.report(coverage)
    report["out_dir"] = str(paths.root)
    path = paths.root / "report.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _print_summary(pipeline: Pipeline, paths: RunPaths, coverage: "Coverage | None") -> None:
    print()
    print("=" * 62)
    if pipeline.raw_export_path:
        print(f"Выгрузка DoQA:      {pipeline.raw_export_path}")
    if pipeline.items:
        print(f"Чек-лист:           {len(pipeline.items)} пунктов")
        print(f"  {paths.checklist / 'checklist.md'}")
    if pipeline.cases:
        print(f"Тест-кейсы:         {len(pipeline.cases)} шт.")
        print(f"  {paths.cases / 'cases.xlsx'}")
    if coverage is not None:
        print(f"Покрытие чек-листа: {coverage.covered_count}/{coverage.total} ({coverage.percent}%)")
    usage = getattr(pipeline.llm, "usage", None)
    if usage is not None:
        print(
            f"Запросы к модели:   {usage.api_calls} (кэш: {usage.cache_hits}), "
            f"токенов: {usage.total_tokens}"
        )
    print(f"Отчёт:              {paths.root / 'report.json'}")
    print("=" * 62)


def _run_export(args: argparse.Namespace, settings: Settings, logger: logging.Logger) -> int:
    settings.require_doqa()
    space_id = args.space_id or settings.doqa_space_id
    if not space_id:
        raise ConfigError("Укажите --space-id или DOQA_SPACE_ID в .env")

    paths = RunPaths(_resolve_run_dir(args, settings, space_id))
    pipeline = Pipeline(settings, DryRunLLM(), paths, logger)
    pipeline.run_export(space_id, _build_filters(args, settings))
    _write_report(paths, pipeline, None)
    _print_summary(pipeline, paths, None)
    return 0


def _run_checklist_only(args: argparse.Namespace, settings: Settings, logger: logging.Logger) -> int:
    requirements = _requirements_path(args, settings)
    paths = RunPaths(_resolve_run_dir(args, settings, settings.doqa_space_id))
    llm = _make_llm(args, settings, paths)
    pipeline = Pipeline(settings, llm, paths, logger)

    docs, chunks = load_requirements(
        requirements, settings.requirements_chunk_chars, logger
    )
    try:
        pipeline.build_checklist(chunks, docs)
    except PipelineError as exc:
        if args.dry_run:
            logger.warning("Dry-run: %s", exc)
            _print_summary(pipeline, paths, None)
            return 0
        raise
    _write_report(paths, pipeline, None)
    _print_summary(pipeline, paths, None)
    return 0


def _run_cases_only(args: argparse.Namespace, settings: Settings, logger: logging.Logger) -> int:
    requirements = _requirements_path(args, settings)
    export_file = args.export_file or settings.doqa_export_file
    if not export_file:
        raise ConfigError("Укажите --export-file или DOQA_EXPORT_FILE в .env")

    paths = RunPaths(_resolve_run_dir(args, settings, settings.doqa_space_id))
    llm = _make_llm(args, settings, paths)
    pipeline = Pipeline(settings, llm, paths, logger)
    pipeline.attach_export_file(export_file)

    docs, chunks = load_requirements(requirements, settings.requirements_chunk_chars, logger)
    if args.checklist_file:
        items = _load_checklist_file(args.checklist_file)
        pipeline.chunks = chunks
        pipeline.docs = docs
        pipeline.items = items
        logger.info("Чек-лист загружен из файла: %d пунктов", len(items))
    else:
        items = pipeline.build_checklist(chunks, docs)

    try:
        cases, coverage = pipeline.generate_cases(items, pipeline.export_bundle.cases if pipeline.export_bundle else [])
    except PipelineError as exc:
        if args.dry_run:
            logger.warning("Dry-run: %s", exc)
            _print_summary(pipeline, paths, None)
            return 0
        raise
    _write_report(paths, pipeline, coverage)
    _print_summary(pipeline, paths, coverage)
    return 0


def _run_all(args: argparse.Namespace, settings: Settings, logger: logging.Logger) -> int:
    requirements = _requirements_path(args, settings)
    export_file = args.export_file or settings.doqa_export_file
    space_id = args.space_id or settings.doqa_space_id
    if not export_file:
        settings.require_doqa()
        if not space_id:
            raise ConfigError("Укажите --space-id или DOQA_SPACE_ID в .env")

    paths = RunPaths(_resolve_run_dir(args, settings, space_id))
    llm = _make_llm(args, settings, paths)
    pipeline = Pipeline(settings, llm, paths, logger)

    if export_file:
        pipeline.attach_export_file(export_file)
    else:
        pipeline.run_export(int(space_id), _build_filters(args, settings))

    docs, chunks = load_requirements(requirements, settings.requirements_chunk_chars, logger)
    try:
        items = pipeline.build_checklist(chunks, docs)
    except PipelineError as exc:
        if args.dry_run:
            logger.warning("Dry-run: %s", exc)
            _write_report(paths, pipeline, None)
            _print_summary(pipeline, paths, None)
            return 0
        raise

    exported = pipeline.export_bundle.cases if pipeline.export_bundle else []
    try:
        cases, coverage = pipeline.generate_cases(items, exported)
    except PipelineError as exc:
        if args.dry_run:
            logger.warning("Dry-run: %s", exc)
            _write_report(paths, pipeline, None)
            _print_summary(pipeline, paths, None)
            return 0
        raise

    _write_report(paths, pipeline, coverage)
    _print_summary(pipeline, paths, coverage)
    return 0


def main(argv: "list[str] | None" = None) -> int:
    configure_console()
    parser = build_parser()
    args = parser.parse_args(argv)

    env_path = Path(args.env_file)
    if not env_path.is_absolute():
        env_path = Path.cwd() / env_path
    if not env_path.is_file() and (PROJECT_ROOT / ".env").is_file():
        env_path = PROJECT_ROOT / ".env"
    try:
        load_env_file(env_path)
        settings = _resolve_settings(args)
        space_id = getattr(args, "space_id", None) or settings.doqa_space_id
        paths = RunPaths(_resolve_run_dir(args, settings, space_id))
        logger = _configure_logging(paths.logs / "run.log", args.verbose, settings.log_level)
        logger.debug("Каталог запуска: %s", paths.root)

        if args.command == "export":
            return _run_export(args, settings, logger)
        if args.command == "checklist":
            return _run_checklist_only(args, settings, logger)
        if args.command == "cases":
            return _run_cases_only(args, settings, logger)
        if args.command == "all":
            return _run_all(args, settings, logger)
        parser.error(f"Неизвестная команда: {args.command}")
        return 2
    except KeyboardInterrupt:
        print("\nПрервано пользователем", file=sys.stderr)
        return 130
    except (ConfigError, DoQAError, CaseExportError, RequirementsError,
            PromptError, LLMError, PipelineError, ModelError) as exc:
        print(f"\nОшибка: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # noqa: BLE001 - верхнеуровневая защита CLI
        print(f"\nНепредвиденная ошибка: {exc!r}", file=sys.stderr)
        if args.verbose:
            raise
        return 2


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
