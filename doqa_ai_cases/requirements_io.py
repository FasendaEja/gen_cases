from __future__ import annotations

import csv
import hashlib
import html
import io
import json
import logging
import re
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path


class RequirementsError(RuntimeError):
    """Не удалось прочитать требования."""


TEXT_SUFFIXES = {".md", ".markdown", ".txt", ".rst", ".adoc", ".log"}
DOCX_SUFFIXES = {".docx"}
PDF_SUFFIXES = {".pdf"}
XLSX_SUFFIXES = {".xlsx", ".xlsm"}
HTML_SUFFIXES = {".html", ".htm"}
CSV_SUFFIXES = {".csv", ".tsv"}
JSON_SUFFIXES = {".json"}

SUPPORTED_SUFFIXES = (
    TEXT_SUFFIXES
    | DOCX_SUFFIXES
    | PDF_SUFFIXES
    | XLSX_SUFFIXES
    | HTML_SUFFIXES
    | CSV_SUFFIXES
    | JSON_SUFFIXES
)

SKIP_DIRS = {
    ".git",
    ".svn",
    ".hg",
    ".idea",
    ".vscode",
    "__pycache__",
    "node_modules",
    ".venv",
    "venv",
    "env",
    "dist",
    "build",
    ".mypy_cache",
    ".pytest_cache",
}


@dataclass
class RequirementDoc:
    path: Path
    rel_path: str
    text: str


@dataclass
class RequirementChunk:
    id: str
    doc_rel_path: str
    heading: str
    text: str
    index: int
    total: int


# --------------------------------------------------------------------------- #
#  Чтение файлов
# --------------------------------------------------------------------------- #


def _read_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1251", "utf-16"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


class _HTMLTextExtractor(HTMLParser):
    _BLOCK_TAGS = {
        "p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6",
        "table", "ul", "ol", "section", "article", "header", "footer",
    }
    _SKIP_TAGS = {"script", "style", "noscript"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self._skip_depth = 0

    def handle_starttag(self, tag: str, attrs) -> None:  # noqa: ANN001
        if tag in self._SKIP_TAGS:
            self._skip_depth += 1
        elif tag in self._BLOCK_TAGS:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in self._SKIP_TAGS:
            self._skip_depth = max(0, self._skip_depth - 1)
        elif tag in self._BLOCK_TAGS:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if not self._skip_depth:
            self.parts.append(data)


def _read_html(path: Path) -> str:
    parser = _HTMLTextExtractor()
    parser.feed(_read_text(path))
    parser.close()
    text = "".join(parser.parts)
    text = html.unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n\s*\n+", "\n\n", text)
    return text.strip()


def _read_docx(path: Path) -> str:
    try:
        from docx import Document  # type: ignore
    except ImportError as exc:  # pragma: no cover
        raise RequirementsError(
            "Для чтения .docx установите python-docx: pip install python-docx"
        ) from exc

    document = Document(str(path))
    parts: list[str] = []

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        style_name = ""
        try:
            style_name = (paragraph.style.name or "").lower() if paragraph.style else ""
        except AttributeError:
            style_name = ""
        if style_name.startswith(("heading", "заголовок", "title")) or style_name == "title":
            parts.append(f"## {text}")
        else:
            parts.append(text)

    for table in document.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            if any(cells):
                parts.append(" | ".join(cells))

    return "\n\n".join(parts)


def _read_pdf(path: Path) -> str:
    try:
        from pypdf import PdfReader  # type: ignore
    except ImportError as exc:
        raise RequirementsError(
            "Для чтения .pdf установите pypdf: pip install pypdf (или удалите PDF из папки)"
        ) from exc

    reader = PdfReader(str(path))
    parts: list[str] = []
    for page in reader.pages:
        try:
            text = page.extract_text() or ""
        except Exception:  # pragma: no cover - повреждённые страницы
            text = ""
        if text.strip():
            parts.append(text.strip())
    return "\n\n".join(parts)


def _read_xlsx(path: Path) -> str:
    try:
        from openpyxl import load_workbook  # type: ignore
    except ImportError as exc:  # pragma: no cover
        raise RequirementsError("Для чтения .xlsx установите openpyxl") from exc

    workbook = load_workbook(filename=str(path), read_only=True, data_only=True)
    parts: list[str] = []
    try:
        for sheet in workbook.worksheets:
            rows: list[str] = []
            for row in sheet.iter_rows(values_only=True):
                cells = [str(value).strip() for value in row if value not in (None, "")]
                if cells:
                    rows.append(" | ".join(cells))
            if rows:
                parts.append(f"## {sheet.title}\n" + "\n".join(rows))
    finally:
        workbook.close()
    return "\n\n".join(parts)


def _read_csv(path: Path) -> str:
    raw = _read_text(path)
    delimiter = "\t" if path.suffix.lower() == ".tsv" else ","
    reader = csv.reader(io.StringIO(raw), delimiter=delimiter)
    rows = [" | ".join(cell.strip() for cell in row) for row in reader if any(cell.strip() for cell in row)]
    return "\n".join(rows)


def _read_json(path: Path) -> str:
    raw = _read_text(path)
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError:
        return raw
    return json.dumps(payload, ensure_ascii=False, indent=2)


def read_document(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in TEXT_SUFFIXES:
        return _read_text(path)
    if suffix in DOCX_SUFFIXES:
        return _read_docx(path)
    if suffix in PDF_SUFFIXES:
        return _read_pdf(path)
    if suffix in XLSX_SUFFIXES:
        return _read_xlsx(path)
    if suffix in HTML_SUFFIXES:
        return _read_html(path)
    if suffix in CSV_SUFFIXES:
        return _read_csv(path)
    if suffix in JSON_SUFFIXES:
        return _read_json(path)
    raise RequirementsError(f"Формат {suffix} не поддерживается: {path}")


def iter_requirement_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        relative_parts = path.relative_to(root).parts[:-1]
        if any(part.lower() in SKIP_DIRS or part.startswith(".") for part in relative_parts):
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in SUPPORTED_SUFFIXES:
            continue
        files.append(path)
    return files


# --------------------------------------------------------------------------- #
#  Разбиение на фрагменты
# --------------------------------------------------------------------------- #

_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")


def _split_paragraphs(text: str) -> list[str]:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    normalized = normalized.replace("\x0b", "\n")
    blocks = re.split(r"\n\s*\n", normalized)
    return [block.strip() for block in blocks if block.strip()]


def _split_long_block(block: str, limit: int) -> list[str]:
    pieces: list[str] = []
    remaining = block
    while len(remaining) > limit:
        cut = remaining.rfind("\n", 0, limit)
        if cut < limit // 2:
            cut = remaining.rfind(" ", 0, limit)
        if cut < limit // 2:
            cut = limit
        piece = remaining[:cut].strip()
        if piece:
            pieces.append(piece)
        remaining = remaining[cut:].strip()
    if remaining:
        pieces.append(remaining)
    return pieces


def chunk_document(
    text: str,
    chunk_chars: int,
    doc_index: int,
    rel_path: str,
) -> list[RequirementChunk]:
    paragraphs = _split_paragraphs(text)
    if not paragraphs:
        return []

    chunk_chars = max(1000, int(chunk_chars))
    raw_chunks: list[tuple[str, str]] = []
    heading = ""
    buffer: list[str] = []
    buffer_len = 0

    def flush() -> None:
        nonlocal buffer, buffer_len
        if not buffer:
            return
        body = "\n\n".join(buffer).strip()
        if body:
            full_text = f"{heading}\n\n{body}" if heading else body
            raw_chunks.append((heading or "Общее", full_text))
        buffer = []
        buffer_len = 0

    for paragraph in paragraphs:
        match = _HEADING_RE.match(paragraph)
        if match and len(paragraph.splitlines()) == 1:
            flush()
            heading = match.group(2).strip() or heading
            continue

        if buffer and buffer_len + len(paragraph) + 2 > chunk_chars:
            flush()

        if len(paragraph) > chunk_chars:
            for piece in _split_long_block(paragraph, chunk_chars):
                if buffer and buffer_len + len(piece) + 2 > chunk_chars:
                    flush()
                buffer.append(piece)
                buffer_len += len(piece) + 2
            continue

        buffer.append(paragraph)
        buffer_len += len(paragraph) + 2

    flush()

    total = len(raw_chunks)
    return [
        RequirementChunk(
            id=f"REQ-{doc_index:03d}-{index:02d}",
            doc_rel_path=rel_path,
            heading=chunk_heading,
            text=chunk_text,
            index=index,
            total=total,
        )
        for index, (chunk_heading, chunk_text) in enumerate(raw_chunks, start=1)
    ]


# --------------------------------------------------------------------------- #
#  Публичный API
# --------------------------------------------------------------------------- #


def load_requirements(
    root: "Path | str",
    chunk_chars: int = 12000,
    logger: "logging.Logger | None" = None,
) -> tuple[list[RequirementDoc], list[RequirementChunk]]:
    log = logger or logging.getLogger(__name__)
    root_path = Path(root).expanduser()
    if not root_path.is_dir():
        raise RequirementsError(f"Папка с требованиями не найдена: {root_path}")

    files = iter_requirement_files(root_path)
    if not files:
        raise RequirementsError(
            f"В папке {root_path} не найдено файлов поддерживаемых форматов: "
            + ", ".join(sorted(SUPPORTED_SUFFIXES))
        )

    docs: list[RequirementDoc] = []
    seen_hashes: set[str] = set()
    for path in files:
        rel_path = path.relative_to(root_path).as_posix()
        try:
            text = read_document(path)
        except RequirementsError as exc:
            log.warning("Пропущен файл %s: %s", rel_path, exc)
            continue
        except Exception as exc:  # noqa: BLE001 - один битый файл не должен ронять запуск
            log.warning("Не удалось прочитать %s: %s", rel_path, exc)
            continue
        text = text.strip()
        if not text:
            log.warning("Пропущен пустой файл: %s", rel_path)
            continue
        digest = hashlib.sha1(text.encode("utf-8", errors="ignore")).hexdigest()
        if digest in seen_hashes:
            log.warning("Пропущен дубликат по содержимому: %s", rel_path)
            continue
        seen_hashes.add(digest)
        docs.append(RequirementDoc(path=path, rel_path=rel_path, text=text))

    if not docs:
        raise RequirementsError(f"Не удалось прочитать ни одного файла требований в {root_path}")

    chunks: list[RequirementChunk] = []
    for doc_index, doc in enumerate(docs, start=1):
        chunks.extend(chunk_document(doc.text, chunk_chars, doc_index, doc.rel_path))

    if not chunks:
        raise RequirementsError("Требования прочитаны, но не удалось разбить их на фрагменты")

    log.info(
        "Требования: %d документ(ов), %d фрагмент(ов), %d символов",
        len(docs),
        len(chunks),
        sum(len(chunk.text) for chunk in chunks),
    )
    return docs, chunks
