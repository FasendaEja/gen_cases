from __future__ import annotations

import json
import logging
import re
import ssl
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import unquote
from urllib.request import Request, urlopen

from .config import Settings


class DoQAError(RuntimeError):
    """Ошибка при обращении к DoQA."""

    def __init__(self, message: str, status: "int | None" = None, body: str = "") -> None:
        super().__init__(message)
        self.status = status
        self.body = body


@dataclass
class ExportResult:
    content: bytes
    filename: str
    content_type: str
    status: int

    @property
    def extension(self) -> str:
        suffix = Path(self.filename).suffix.lower()
        if suffix in {".xlsx", ".pdf", ".zip"}:
            return suffix
        content_type = self.content_type.lower()
        if "spreadsheetml" in content_type or "xlsx" in content_type:
            return ".xlsx"
        if "pdf" in content_type:
            return ".pdf"
        return ".xlsx"


_FILENAME_RE = re.compile(r"filename\*?=(?:UTF-8''|[\"']?)([^\"';]+)", re.IGNORECASE)


def _decode_body(raw: bytes) -> str:
    for encoding in ("utf-8", "cp1251"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _extract_error_message(body: str) -> str:
    try:
        payload = json.loads(body)
    except (ValueError, TypeError):
        return body.strip()[:800]
    if isinstance(payload, dict):
        for key in ("message", "detail", "error", "title"):
            value = payload.get(key)
            if isinstance(value, str) and value.strip():
                return value.strip()[:800]
            if isinstance(value, dict):
                nested = value.get("message")
                if isinstance(nested, str) and nested.strip():
                    return nested.strip()[:800]
    return body.strip()[:800]


class DoQAClient:
    """HTTP-клиент для POST /api/cases/export."""

    def __init__(self, settings: Settings, logger: "logging.Logger | None" = None) -> None:
        self.settings = settings
        self.log = logger or logging.getLogger(__name__)

    # ------------------------------------------------------------------ URL
    def _url(self) -> str:
        base = self.settings.doqa_base_url.rstrip("/")
        path = self.settings.doqa_export_path.strip()
        if not path.startswith("/"):
            path = "/" + path
        if base.lower().endswith("/api") and path.lower().startswith("/api/"):
            path = path[len("/api"):]
        return base + path

    # --------------------------------------------------------------- экспорт
    def export_cases(self, body: dict[str, Any]) -> ExportResult:
        if "spaceId" not in body:
            raise DoQAError("В теле запроса отсутствует обязательный параметр spaceId")

        url = self._url()
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        request = Request(url, data=data, headers=self.settings.doqa_headers(), method="POST")
        context = self.settings.ssl_context()

        self.log.info("Запрос экспорта DoQA: POST %s (spaceId=%s, format=%s)",
                      url, body.get("spaceId"), body.get("format", "xlsx"))

        try:
            response = urlopen(request, timeout=self.settings.doqa_timeout, context=context)
        except HTTPError as exc:
            raw = exc.read() if exc.fp else b""
            message = _extract_error_message(_decode_body(raw))
            raise DoQAError(
                f"DoQA вернул HTTP {exc.code} при экспорте кейсов: {message}",
                status=exc.code,
                body=_decode_body(raw),
            ) from exc
        except URLError as exc:
            raise DoQAError(
                f"Не удалось подключиться к DoQA по адресу {url}: {exc.reason}"
            ) from exc
        except (TimeoutError, ssl.SSLError) as exc:
            raise DoQAError(f"Ошибка сети/SSL при запросе к DoQA ({url}): {exc}") from exc

        with response:
            content = response.read()
            status = getattr(response, "status", 200)
            content_type = response.headers.get("Content-Type", "") or ""
            disposition = response.headers.get("Content-Disposition", "") or ""

        filename = "cases_export"
        match = _FILENAME_RE.search(disposition)
        if match:
            filename = unquote(match.group(1).strip()) or filename

        if "json" in content_type.lower() or content.strip().startswith(b"{"):
            message = _extract_error_message(_decode_body(content))
            raise DoQAError(
                f"DoQA вернул JSON вместо файла экспорта: {message}", status=status,
                body=_decode_body(content),
            )

        if not content:
            raise DoQAError(f"DoQA вернул пустой ответ (HTTP {status})")

        self.log.info("Получен файл экспорта: %s (%d байт, %s)", filename, len(content), content_type)
        return ExportResult(content=content, filename=filename,
                            content_type=content_type, status=status)
