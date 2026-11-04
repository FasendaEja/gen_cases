from __future__ import annotations

import re
from pathlib import Path


class PromptError(RuntimeError):
    """Проблема с файлами промптов."""


_PLACEHOLDER_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")


class PromptLibrary:
    """Шаблоны промптов из prompts/: <name>.system.md и <name>.user.md.

    Подстановка — через {{placeholder}}. Это позволяет держать в промптах
    обычные фигурные скобки JSON без экранирования.
    """

    def __init__(self, directory: "Path | str") -> None:
        self.directory = Path(directory)
        self._cache: dict[str, str] = {}

    def _read(self, filename: str) -> str:
        if filename in self._cache:
            return self._cache[filename]
        path = self.directory / filename
        if not path.is_file():
            raise PromptError(f"Не найден файл промпта: {path}")
        text = path.read_text(encoding="utf-8-sig").strip()
        self._cache[filename] = text
        return text

    def system(self, name: str) -> str:
        return self._read(f"{name}.system.md")

    def user_template(self, name: str) -> str:
        return self._read(f"{name}.user.md")

    def render(self, name: str, **values: object) -> str:
        template = self.user_template(name)

        def substitute(match: "re.Match[str]") -> str:
            key = match.group(1)
            if key not in values:
                raise PromptError(
                    f"В промпте {name}.user.md есть плейсхолдер {{{{{key}}}}}, "
                    "но значение не передано"
                )
            return str(values[key])

        return _PLACEHOLDER_RE.sub(substitute, template)

    def available(self) -> list[str]:
        names = set()
        if self.directory.is_dir():
            for path in self.directory.glob("*.user.md"):
                names.add(path.name[: -len(".user.md")])
        return sorted(names)
