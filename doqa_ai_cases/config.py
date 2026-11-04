from __future__ import annotations

import base64
import json
import os
import ssl
from dataclasses import dataclass, field
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent


class ConfigError(RuntimeError):
    """Некорректная или неполная конфигурация (env-переменные)."""


def load_env_file(path: "Path | str | None", override: bool = False) -> dict[str, str]:
    """Читает .env в духе python-dotenv: KEY=VALUE, # комментарии, кавычки."""
    loaded: dict[str, str] = {}
    if not path:
        return loaded
    env_path = Path(path)
    if not env_path.is_file():
        return loaded

    for raw_line in env_path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export "):].strip()
        key, sep, value = line.partition("=")
        if not sep:
            continue
        key = key.strip()
        value = value.strip()
        if not key:
            continue
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]
        loaded[key] = value
        if override or key not in os.environ:
            os.environ[key] = value
    return loaded


def _env_str(name: str, default: str = "") -> str:
    value = os.environ.get(name)
    return default if value is None else str(value).strip()


def _env_int(name: str, default: int) -> int:
    raw = _env_str(name)
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError as exc:
        raise ConfigError(f"{name}: ожидается целое число, получено {raw!r}") from exc


def _env_float(name: str, default: float) -> float:
    raw = _env_str(name)
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError as exc:
        raise ConfigError(f"{name}: ожидается число, получено {raw!r}") from exc


def _env_bool(name: str, default: bool) -> bool:
    raw = _env_str(name)
    if not raw:
        return default
    normalized = raw.lower()
    if normalized in {"1", "true", "yes", "y", "on", "да"}:
        return True
    if normalized in {"0", "false", "no", "n", "off", "нет"}:
        return False
    raise ConfigError(f"{name}: ожидается true/false, получено {raw!r}")


def _env_path(name: str, default: Path) -> Path:
    raw = _env_str(name)
    if not raw:
        return default
    path = Path(raw).expanduser()
    return path if path.is_absolute() else (PROJECT_ROOT / path)


def _env_optional_path(name: str) -> "Path | None":
    raw = _env_str(name)
    if not raw:
        return None
    path = Path(raw).expanduser()
    return path if path.is_absolute() else (PROJECT_ROOT / path)


def _env_json_object(name: str) -> dict[str, str]:
    raw = _env_str(name)
    if not raw:
        return {}
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ConfigError(f"{name}: ожидается JSON-объект, получено {raw!r}") from exc
    if not isinstance(data, dict):
        raise ConfigError(f"{name}: ожидается JSON-объект")
    return {str(k): str(v) for k, v in data.items()}


@dataclass
class Settings:
    """Все настройки запуска, собранные из переменных окружения."""

    # --- DoQA ---
    doqa_base_url: str = ""
    doqa_token: str = ""
    doqa_username: str = ""
    doqa_password: str = ""
    doqa_cookie: str = ""
    doqa_space_id: "int | None" = None
    doqa_export_path: str = "/api/cases/export"
    doqa_auth_header: str = "Authorization"
    doqa_auth_scheme: str = "Bearer"
    doqa_extra_headers: dict[str, str] = field(default_factory=dict)
    doqa_timeout: float = 180.0
    doqa_verify_ssl: bool = True
    doqa_export_file: "Path | None" = None

    # --- OpenAI ---
    openai_api_key: str = ""
    openai_base_url: str = "https://api.openai.com/v1"
    openai_model: str = "gpt-4o-mini"
    openai_temperature: "float | None" = 0.2
    openai_max_tokens: "int | None" = None
    openai_timeout: float = 240.0
    openai_max_retries: int = 4
    openai_json_mode: str = "auto"

    # --- Конвейер ---
    requirements_dir: "Path | None" = None
    requirements_chunk_chars: int = 12000
    checklist_consolidate: bool = True
    consolidate_max_chars: int = 60_000
    requirements_context_max_chars: int = 18_000
    reference_case_max_chars: int = 2_200
    similar_cases_k: int = 4
    max_reference_cases: int = 8
    case_batch_size: int = 3
    max_retry_items: int = 20
    llm_cache: bool = True
    llm_cache_dir: Path = field(default_factory=lambda: PROJECT_ROOT / ".cache" / "llm")
    save_prompts: bool = True
    prompts_dir: Path = field(default_factory=lambda: PROJECT_ROOT / "prompts")
    runs_dir: Path = field(default_factory=lambda: PROJECT_ROOT / "runs")
    export_format: str = "xlsx"
    log_level: str = "INFO"

    # ------------------------------------------------------------------ сборка
    @classmethod
    def from_env(cls) -> "Settings":
        temperature: "float | None" = None
        raw_temperature = _env_str("OPENAI_TEMPERATURE")
        if raw_temperature:
            try:
                temperature = float(raw_temperature)
            except ValueError as exc:
                raise ConfigError(
                    f"OPENAI_TEMPERATURE: ожидается число, получено {raw_temperature!r}"
                ) from exc

        raw_space_id = _env_str("DOQA_SPACE_ID")
        space_id: "int | None" = None
        if raw_space_id:
            try:
                space_id = int(raw_space_id)
            except ValueError as exc:
                raise ConfigError(
                    f"DOQA_SPACE_ID: ожидается целое число, получено {raw_space_id!r}"
                ) from exc

        json_mode = _env_str("OPENAI_JSON_MODE", "auto").lower() or "auto"
        if json_mode not in {"auto", "json_schema", "json_object", "text"}:
            raise ConfigError(
                "OPENAI_JSON_MODE: допустимо auto | json_schema | json_object | text, "
                f"получено {json_mode!r}"
            )

        export_format = _env_str("EXPORT_FORMAT", "xlsx").lower() or "xlsx"
        if export_format not in {"xlsx", "pdf"}:
            raise ConfigError(
                f"EXPORT_FORMAT: допустимо xlsx | pdf, получено {export_format!r}"
            )

        log_level = _env_str("LOG_LEVEL", "INFO").upper() or "INFO"
        if log_level not in {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}:
            raise ConfigError(
                "LOG_LEVEL: допустимо DEBUG | INFO | WARNING | ERROR | CRITICAL, "
                f"получено {log_level!r}"
            )

        return cls(
            doqa_base_url=_env_str("DOQA_BASE_URL").rstrip("/"),
            doqa_token=_env_str("DOQA_TOKEN"),
            doqa_username=_env_str("DOQA_USERNAME"),
            doqa_password=_env_str("DOQA_PASSWORD"),
            doqa_cookie=_env_str("DOQA_COOKIE"),
            doqa_space_id=space_id,
            doqa_export_path=_env_str("DOQA_EXPORT_PATH", "/api/cases/export"),
            doqa_auth_header=_env_str("DOQA_AUTH_HEADER", "Authorization"),
            doqa_auth_scheme=_env_str("DOQA_AUTH_SCHEME", "Bearer"),
            doqa_extra_headers=_env_json_object("DOQA_EXTRA_HEADERS"),
            doqa_timeout=_env_float("DOQA_TIMEOUT", 180.0),
            doqa_verify_ssl=_env_bool("DOQA_VERIFY_SSL", True),
            doqa_export_file=_env_optional_path("DOQA_EXPORT_FILE"),
            openai_api_key=_env_str("OPENAI_API_KEY"),
            openai_base_url=_env_str("OPENAI_BASE_URL", "https://api.openai.com/v1"),
            openai_model=_env_str("OPENAI_MODEL", "gpt-4o-mini"),
            openai_temperature=temperature,
            openai_max_tokens=_env_int("OPENAI_MAX_TOKENS", 0) or None,
            openai_timeout=_env_float("OPENAI_TIMEOUT", 240.0),
            openai_max_retries=_env_int("OPENAI_MAX_RETRIES", 4),
            openai_json_mode=json_mode,
            requirements_dir=_env_optional_path("REQUIREMENTS_DIR"),
            requirements_chunk_chars=max(2000, _env_int("REQUIREMENTS_CHUNK_CHARS", 12000)),
            checklist_consolidate=_env_bool("CHECKLIST_CONSOLIDATE", True),
            consolidate_max_chars=max(5_000, _env_int("CONSOLIDATE_MAX_CHARS", 60_000)),
            requirements_context_max_chars=max(
                2_000, _env_int("REQUIREMENTS_CONTEXT_MAX_CHARS", 18_000)
            ),
            reference_case_max_chars=max(400, _env_int("REFERENCE_CASE_MAX_CHARS", 2_200)),
            similar_cases_k=max(1, _env_int("SIMILAR_CASES_K", 4)),
            max_reference_cases=max(1, _env_int("MAX_REFERENCE_CASES", 8)),
            case_batch_size=max(1, _env_int("CASE_BATCH_SIZE", 3)),
            max_retry_items=max(0, _env_int("MAX_RETRY_ITEMS", 20)),
            llm_cache=_env_bool("LLM_CACHE", True),
            llm_cache_dir=_env_path("LLM_CACHE_DIR", PROJECT_ROOT / ".cache" / "llm"),
            save_prompts=_env_bool("SAVE_PROMPTS", True),
            prompts_dir=_env_path("PROMPTS_DIR", PROJECT_ROOT / "prompts"),
            runs_dir=_env_path("RUNS_DIR", PROJECT_ROOT / "runs"),
            export_format=export_format,
            log_level=log_level,
        )

    # --------------------------------------------------------------- проверки
    def require_doqa(self) -> None:
        missing = []
        if not self.doqa_base_url:
            missing.append("DOQA_BASE_URL")
        has_auth = bool(
            self.doqa_token or (self.doqa_username and self.doqa_password) or self.doqa_cookie
        )
        if not has_auth:
            missing.append("DOQA_TOKEN (или DOQA_USERNAME+DOQA_PASSWORD, или DOQA_COOKIE)")
        if missing:
            raise ConfigError(
                "Не настроено подключение к DoQA. Заполните .env: " + "; ".join(missing)
            )

    def require_llm(self) -> None:
        missing = []
        if not self.openai_api_key:
            missing.append("OPENAI_API_KEY")
        if not self.openai_model:
            missing.append("OPENAI_MODEL")
        if missing:
            raise ConfigError(
                "Не настроен доступ к модели. Заполните .env: " + "; ".join(missing)
            )

    # ---------------------------------------------------------------- helpers
    def doqa_headers(self) -> dict[str, str]:
        headers = {
            "Accept": "application/json, application/octet-stream, */*",
            "Content-Type": "application/json",
            "User-Agent": "doqa-ai-cases/1.0",
        }
        headers.update(self.doqa_extra_headers)
        if self.doqa_token:
            value = f"{self.doqa_auth_scheme} {self.doqa_token}".strip()
            headers[self.doqa_auth_header] = value
        if self.doqa_cookie:
            headers["Cookie"] = self.doqa_cookie
        if self.doqa_username and self.doqa_password:
            credentials = f"{self.doqa_username}:{self.doqa_password}".encode("utf-8")
            headers["Authorization"] = "Basic " + base64.b64encode(credentials).decode("ascii")
        return headers

    def ssl_context(self) -> "ssl.SSLContext":
        context = ssl.create_default_context()
        if not self.doqa_verify_ssl:
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE
        return context
