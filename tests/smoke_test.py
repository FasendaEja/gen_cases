"""Офлайн smoke-тест: DoQA-заглушка, фейковая модель, полный конвейер.

Запуск:  python tests/smoke_test.py
Не требует ключей OpenAI и доступа к DoQA.
"""

from __future__ import annotations

import json
import logging
import os
import re
import sys
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402

from doqa_ai_cases.case_export import load_export, write_cases_xlsx  # noqa: E402
from doqa_ai_cases.config import Settings, load_env_file  # noqa: E402
from doqa_ai_cases import cli as cli_module  # noqa: E402
from doqa_ai_cases.llm import Usage  # noqa: E402
from doqa_ai_cases.models import GeneratedCase  # noqa: E402
from doqa_ai_cases.pipeline import Pipeline, RunPaths  # noqa: E402
from doqa_ai_cases.requirements_io import load_requirements  # noqa: E402


HEADERS = [
    "№",
    "ID",
    "Папка",
    "Название",
    "Автор",
    "Дата создания",
    "Приоритет",
    "Статус",
    "Теги",
    "Описание",
    "Параметры",
    "Предусловия",
    "Постусловия",
    "Вложения",
    "Шаг 1",
    "Ожидаемый результат 1",
    "Шаг 2",
    "Ожидаемый результат 2",
    "Тестовые данные",
    "Статус автоматизации",
    "Автотест",
    "Атрибуты",
]


def make_export_file(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Кейсы"
    sheet.cell(row=1, column=1, value="Выгрузка тест-кейсов DoQA")
    for index, header in enumerate(HEADERS, start=1):
        sheet.cell(row=3, column=index, value=header)

    sheet.cell(row=4, column=1, value=1)
    sheet.cell(row=4, column=2, value=101)
    sheet.cell(row=4, column=3, value="Регресс/Авторизация")
    sheet.cell(row=4, column=4, value="Вход с валидными данными")
    sheet.cell(row=4, column=7, value="high")
    sheet.cell(row=4, column=8, value="ready")
    sheet.cell(row=4, column=10, value="Проверка успешного входа")
    sheet.cell(row=4, column=12, value="Пользователь зарегистрирован")
    sheet.cell(row=4, column=15, value="Открыть страницу входа")
    sheet.cell(row=4, column=16, value="Отображается форма входа")
    sheet.cell(row=4, column=17, value="Ввести валидные email и пароль")
    sheet.cell(row=4, column=18, value="Выполнен переход в личный кабинет")

    sheet.cell(row=5, column=1, value=2)
    sheet.cell(row=5, column=2, value=102)
    sheet.cell(row=5, column=3, value="Регресс/Авторизация")
    sheet.cell(row=5, column=4, value="Вход с неверным паролем")
    sheet.cell(row=5, column=7, value="medium")
    sheet.cell(row=5, column=15, value="Открыть страницу входа")
    sheet.cell(row=5, column=16, value="Отображается форма входа")
    sheet.cell(row=5, column=17, value="Ввести неверный пароль")
    sheet.cell(row=5, column=18, value="Показано сообщение об ошибке")

    workbook.save(str(path))


def make_requirements(folder: Path) -> None:
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "login.md").write_text(
        "# Авторизация\n\n"
        "Вход выполняется по email и паролю. Пароль не короче 8 символов.\n\n"
        "## Ошибки\n\n"
        "При неверном пароле показывается сообщение «Неверный логин или пароль».\n"
        "После 5 неудачных попыток учётная запись блокируется на 15 минут.\n",
        encoding="utf-8",
    )
    sub = folder / "sub"
    sub.mkdir(exist_ok=True)
    (sub / "limits.txt").write_text(
        "Максимальная длина email — 254 символа.\n"
        "Поле пароля обязательно для заполнения.\n",
        encoding="utf-8",
    )
    try:
        from docx import Document

        document = Document()
        document.add_heading("Восстановление пароля", level=2)
        document.add_paragraph("Ссылка на сброс пароля действует 24 часа.")
        document.save(str(folder / "reset.docx"))
    except ImportError:
        pass


class ScriptedLLM:
    """Детерминированная заглушка вместо OpenAI."""

    def __init__(self) -> None:
        self.usage = Usage()
        self.calls: list[str] = []
        self.seen_prompts: list[tuple[str, str]] = []

    def json_call(self, *, system, user, schema, schema_name, prompt_name="request", temperature="default"):
        self.calls.append(prompt_name)
        self.seen_prompts.append((prompt_name, user))
        self.usage.api_calls += 1

        if schema_name == "checklist_consolidate":
            return {"items": []}
        if schema_name == "checklist":
            return {
                "items": [
                    {
                        "area": "Авторизация",
                        "title": "Вход с валидными данными",
                        "type": "positive",
                        "priority": "high",
                        "requirements": ["login.md#Авторизация"],
                        "preconditions": "Пользователь зарегистрирован",
                        "checks": ["Открыть страницу входа", "Ввести валидные данные"],
                        "expected": "Открывается личный кабинет",
                        "notes": "",
                    },
                    {
                        "area": "Авторизация",
                        "title": "Вход с неверным паролем",
                        "type": "negative",
                        "priority": "medium",
                        "requirements": ["login.md#Ошибки"],
                        "preconditions": "Пользователь зарегистрирован",
                        "checks": ["Ввести неверный пароль"],
                        "expected": "Показано сообщение «Неверный логин или пароль»",
                        "notes": "",
                    },
                ]
            }

        ids = sorted(set(re.findall(r"CHK-\d{3}", user)))
        cases = []
        for index, checklist_id in enumerate(ids):
            steps_count = 3 if index == 0 else 2
            steps = []
            for step_index in range(1, steps_count + 1):
                steps.append(
                    {
                        "action": f"Действие {step_index} для {checklist_id}",
                        "expected_result": f"Результат {step_index}",
                        "test_data": "",
                    }
                )
            cases.append(
                {
                    "checklist_id": checklist_id,
                    "title": f"Кейс для {checklist_id}",
                    "description": "Проверка по чек-листу",
                    "area": "Авторизация",
                    "priority": "high" if index == 0 else "medium",
                    "preconditions": "Пользователь зарегистрирован",
                    "postconditions": "",
                    "steps": steps,
                    "test_data": "email: user@example.com",
                    "automation_status": "manual",
                    "source_cases": ["101"],
                }
            )
        return {"cases": cases}


class FakeDoQAHandler(BaseHTTPRequestHandler):
    payload = b""
    requests: list[dict] = []

    def do_POST(self) -> None:  # noqa: N802
        length = int(self.headers.get("Content-Length") or 0)
        body = json.loads(self.rfile.read(length) or b"{}")
        FakeDoQAHandler.requests.append(body)
        if self.path != "/api/cases/export":
            self.send_response(404)
            self.end_headers()
            return
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.send_header("Content-Disposition", 'attachment; filename="cases.xlsx"')
        self.send_header("Content-Length", str(len(FakeDoQAHandler.payload)))
        self.end_headers()
        self.wfile.write(FakeDoQAHandler.payload)

    def log_message(self, format: str, *args) -> None:  # noqa: A002
        return


def start_fake_doqa(payload: bytes):
    FakeDoQAHandler.payload = payload
    FakeDoQAHandler.requests = []
    server = ThreadingHTTPServer(("127.0.0.1", 0), FakeDoQAHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server, thread


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def test_env_parsing(tmp: Path) -> None:
    env_file = tmp / ".env"
    env_file.write_text(
        '# comment\nDOQA_BASE_URL="https://example.test"\nexport OPENAI_MODEL=gpt-test\n',
        encoding="utf-8",
    )
    loaded = load_env_file(env_file, override=True)
    check(loaded["DOQA_BASE_URL"] == "https://example.test", "не разобраны кавычки в .env")
    check(loaded["OPENAI_MODEL"] == "gpt-test", "не разобран export в .env")


def test_settings_from_env(tmp: Path) -> None:
    env_file = tmp / ".env.settings"
    env_file.write_text(
        "DOQA_BASE_URL=https://doqa.test\n"
        "EXPORT_FORMAT=pdf\n"
        "LOG_LEVEL=warning\n"
        "RUNS_DIR=custom-runs\n"
        "CONSOLIDATE_MAX_CHARS=7000\n"
        "REQUIREMENTS_CONTEXT_MAX_CHARS=5000\n"
        "REFERENCE_CASE_MAX_CHARS=900\n",
        encoding="utf-8",
    )
    keys = (
        "DOQA_BASE_URL",
        "EXPORT_FORMAT",
        "LOG_LEVEL",
        "RUNS_DIR",
        "CONSOLIDATE_MAX_CHARS",
        "REQUIREMENTS_CONTEXT_MAX_CHARS",
        "REFERENCE_CASE_MAX_CHARS",
    )
    saved = {key: os.environ.get(key) for key in keys}
    try:
        load_env_file(env_file, override=True)
        settings = Settings.from_env()
        check(settings.export_format == "pdf", "EXPORT_FORMAT не попал в настройки")
        check(settings.log_level == "WARNING", "LOG_LEVEL не нормализован")
        check(settings.runs_dir.name == "custom-runs", "RUNS_DIR не попал в настройки")
        check(settings.consolidate_max_chars == 7000, "CONSOLIDATE_MAX_CHARS не прочитан")
        check(settings.requirements_context_max_chars == 5000, "REQUIREMENTS_CONTEXT_MAX_CHARS не прочитан")
        check(settings.reference_case_max_chars == 900, "REFERENCE_CASE_MAX_CHARS не прочитан")
    finally:
        for key, value in saved.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


def test_no_template_writer(tmp: Path) -> None:
    cases = [
        GeneratedCase.from_dict(
            {
                "checklist_id": "CHK-001",
                "title": "Кейс без шаблона",
                "priority": "high",
                "steps": [
                    {"action": "Шаг 1", "expected_result": "Результат 1"},
                    {"action": "Шаг 2", "expected_result": "Результат 2"},
                ],
            }
        )
    ]
    path = write_cases_xlsx(cases, tmp / "no_template.xlsx")
    workbook = load_workbook(path)
    sheet = workbook.worksheets[0]
    header = [sheet.cell(row=1, column=col).value for col in range(1, sheet.max_column + 1)]
    check(header[0] == "number", "нет колонки number по умолчанию")
    check("step" in header and "expectedResult" in header, "нет колонок step/expectedResult")
    step_col = header.index("step") + 1
    check(sheet.cell(row=2, column=step_col).value == "1. Шаг 1\n2. Шаг 2", "шаги не склеены в одну ячейку")
    workbook.close()


def test_export_parsing(tmp: Path) -> None:
    export_path = tmp / "export.xlsx"
    make_export_file(export_path)
    bundle = load_export(export_path)
    check(len(bundle.cases) == 2, f"ожидалось 2 кейса, получено {len(bundle.cases)}")
    check(bundle.header_row == 3, f"строка заголовка определена как {bundle.header_row}")
    first = bundle.cases[0]
    check(first.id == "101", f"не разобран ID: {first.id!r}")
    check(first.title == "Вход с валидными данными", f"не разобрано название: {first.title!r}")
    check(len(first.steps) == 2, f"ожидалось 2 шага, получено {len(first.steps)}")
    check(first.steps[1].expected_result == "Выполнен переход в личный кабинет",
          "не разобран ожидаемый результат второго шага")
    check(bundle.sheet_name == "Кейсы", "не разобрано имя листа")


def test_full_pipeline(tmp: Path) -> None:
    export_path = tmp / "export_for_server.xlsx"
    make_export_file(export_path)
    payload = export_path.read_bytes()

    requirements_dir = tmp / "requirements"
    make_requirements(requirements_dir)

    server, _thread = start_fake_doqa(payload)
    try:
        host, port = server.server_address[0], server.server_address[1]
        settings = Settings(
            doqa_base_url=f"http://{host}:{port}",
            doqa_token="test-token",
            openai_api_key="test-key",
            openai_model="test-model",
            llm_cache=False,
            save_prompts=True,
            prompts_dir=ROOT / "prompts",
            requirements_chunk_chars=1200,
            checklist_consolidate=True,
        )
        paths = RunPaths(tmp / "run")
        llm = ScriptedLLM()
        logger = logging.getLogger("smoke")
        logger.addHandler(logging.NullHandler())
        pipeline = Pipeline(settings, llm, paths, logger)

        bundle = pipeline.run_export(77, {"statuses": ["ready"], "folderIds": [5]})
        check(bundle is not None and len(bundle.cases) == 2, "экспорт через HTTP не сработал")
        sent = FakeDoQAHandler.requests[-1]
        check(sent.get("spaceId") == 77, f"spaceId не передан: {sent}")
        check(sent.get("statuses") == ["ready"], f"фильтр статусов не передан: {sent}")
        check(sent.get("folderIds") == [5], f"фильтр папок не передан: {sent}")
        check((paths.doqa / "cases.normalized.json").is_file(), "нет нормализованного JSON выгрузки")

        docs, chunks = load_requirements(requirements_dir, 1200, logger)
        check(len(docs) >= 2, f"прочитано мало документов: {len(docs)}")
        check(len(chunks) >= 2, f"получено мало фрагментов: {len(chunks)}")

        items = pipeline.build_checklist(chunks, docs)
        check(len(items) == 2, f"ожидалось 2 пункта чек-листа после дедупликации, получено {len(items)}")
        check(items[0].id == "CHK-001", "идентификаторы чек-листа не присвоены")
        check((paths.checklist / "checklist.md").is_file(), "не сохранён checklist.md")
        checklist_json = json.loads((paths.checklist / "checklist.json").read_text(encoding="utf-8"))
        check(checklist_json["count"] == 2, "в checklist.json неверное количество пунктов")

        cases, coverage = pipeline.generate_cases(items, bundle.cases)
        check(len(cases) == 2, f"ожидалось 2 кейса, получено {len(cases)}")
        check(coverage.percent == 100.0, f"покрытие чек-листа {coverage.percent}%")
        check((paths.cases / "cases.xlsx").is_file(), "не сохранён cases.xlsx")
        check((paths.cases / "coverage.md").is_file(), "не сохранён coverage.md")

        # Проверяем, что формат xlsx повторяет выгрузку и шаги расширены до 3 пар.
        workbook = load_workbook(paths.cases / "cases.xlsx")
        sheet = workbook.worksheets[0]
        header = [sheet.cell(row=1, column=col).value for col in range(1, sheet.max_column + 1)]
        check("Шаг 3" in header, f"колонка третьего шага не добавлена: {header}")
        check("Ожидаемый результат 3" in header, "колонка ожидаемого результата 3 не добавлена")
        check(sheet.cell(row=1, column=4).value == "Название", "заголовки не взяты из выгрузки")
        check(sheet.cell(row=2, column=4).value == "Кейс для CHK-001", "название кейса не записано")
        check(sheet.max_row == 3, f"ожидалось 2 строки данных, получено {sheet.max_row - 1}")
        workbook.close()

        # Обратный разбор сгенерированного файла тем же парсером.
        generated_bundle = load_export(paths.cases / "cases.xlsx")
        check(len(generated_bundle.cases) == 2, "сгенерированный xlsx не читается парсером")
        check(len(generated_bundle.cases[0].steps) == 3, "не разобрались 3 шага из сгенерированного файла")

        # Промпты сохранены и не содержат незаполненных плейсхолдеров.
        prompt_files = sorted(paths.prompts.glob("*.md"))
        check(prompt_files, "промпты не сохранены")
        for prompt_file in prompt_files:
            text = prompt_file.read_text(encoding="utf-8")
            check("{{" not in text, f"в промпте остался плейсхолдер: {prompt_file.name}")

        report = pipeline.report(coverage)
        check(report["generated_cases"] == 2, "отчёт содержит неверное число кейсов")
        check(report["llm_usage"]["api_calls"] >= 2, "не учтены вызовы модели")
    finally:
        server.shutdown()
        server.server_close()


def test_cli_all(tmp: Path) -> None:
    export_path = tmp / "cli_export.xlsx"
    make_export_file(export_path)
    requirements_dir = tmp / "cli_requirements"
    make_requirements(requirements_dir)
    out_dir = tmp / "cli_run"

    server, _thread = start_fake_doqa(export_path.read_bytes())
    original_llm = cli_module.LLMClient
    env_keys = ("DOQA_BASE_URL", "DOQA_TOKEN", "OPENAI_API_KEY", "OPENAI_MODEL")
    saved_env = {key: os.environ.get(key) for key in env_keys}
    try:
        host, port = server.server_address[0], server.server_address[1]
        os.environ.update(
            {
                "DOQA_BASE_URL": f"http://{host}:{port}",
                "DOQA_TOKEN": "test-token",
                "OPENAI_API_KEY": "test-key",
                "OPENAI_MODEL": "test-model",
            }
        )
        cli_module.LLMClient = lambda settings: ScriptedLLM()  # noqa: E731
        code = cli_module.main(
            [
                "all",
                "--space-id",
                "77",
                "--requirements",
                str(requirements_dir),
                "--out",
                str(out_dir),
                "--no-cache",
                "--no-prompts",
            ]
        )
        check(code == 0, f"CLI all вернул код {code}")
        check((out_dir / "checklist" / "checklist.md").is_file(), "CLI не сохранил чек-лист")
        check((out_dir / "cases" / "cases.xlsx").is_file(), "CLI не сохранил кейсы")
        check((out_dir / "report.json").is_file(), "CLI не сохранил report.json")
        report = json.loads((out_dir / "report.json").read_text(encoding="utf-8"))
        check(report["generated_cases"] == 2, f"в отчёте CLI {report['generated_cases']} кейсов")
        check(report["coverage"]["percent"] == 100.0, "CLI: неполное покрытие чек-листа")
    finally:
        cli_module.LLMClient = original_llm
        for key, value in saved_env.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
        cli_logger = logging.getLogger("doqa_ai_cases")
        for handler in list(cli_logger.handlers):
            try:
                handler.close()
            except Exception:  # noqa: BLE001
                pass
            cli_logger.removeHandler(handler)
        server.shutdown()
        server.server_close()


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="doqa_ai_cases_") as tmp_dir:
        tmp = Path(tmp_dir)
        test_env_parsing(tmp)
        test_settings_from_env(tmp)
        test_no_template_writer(tmp)
        test_export_parsing(tmp)
        test_full_pipeline(tmp)
        test_cli_all(tmp)
    print("SMOKE OK: все проверки пройдены")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
