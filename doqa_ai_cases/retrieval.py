from __future__ import annotations

import math
import re
from collections import Counter
from dataclasses import dataclass

from .models import ExportedCase
from .requirements_io import RequirementChunk


_WORD_RE = re.compile(r"[0-9a-zа-я]+", re.IGNORECASE)

_STOPWORDS = set(
    """
    и в во не что он на я с со как а то все она так его но да ты к у же вы за бы
    по только ее мне было вот от меня еще нет о из ему теперь когда даже ну вдруг
    ли если уже или ни быть был него до вас нибудь опять уж вам ведь там потом
    себя ничего ей может они тут где есть надо ней для мы тебя их чем была сам
    чтоб без будто чего раз тоже себе под будет ж тогда кто этот того потому
    этого какой совсем ним здесь этом один почти мой тем чтобы нее сейчас были
    куда зачем всех никогда можно при наконец два об другой хоть после над больше
    тот через эти нас про всего них какая много разве три эту моя впрочем хорошо
    свою этой перед иногда лучше чуть том нельзя такой им более всегда конечно всю
    между the a an and or of to in on for with is are was were be been it this
    that these those as at by from not no but if then than so such can will would
    should could may might do does did have has had i you he she we they them his
    her its our your their
    """.split()
)


def tokenize(text: str) -> list[str]:
    tokens = []
    for raw in _WORD_RE.findall((text or "").replace("ё", "е").lower()):
        if len(raw) < 3 or raw in _STOPWORDS:
            continue
        tokens.append(_stem(raw))
    return tokens


def _stem(token: str) -> str:
    """Грубая основа для русского/английского: срезаем окончания длинных слов."""
    if len(token) > 6:
        return token[:5]
    if len(token) > 4:
        return token[:4]
    return token


class BM25Index:
    """Компактный BM25 без внешних зависимостей."""

    def __init__(self, texts: list[str], k1: float = 1.5, b: float = 0.75) -> None:
        self.k1 = k1
        self.b = b
        self._tokens = [tokenize(text) for text in texts]
        self._counts = [Counter(tokens) for tokens in self._tokens]
        self._lengths = [len(tokens) for tokens in self._tokens]
        self._size = len(texts)
        self._avgdl = (sum(self._lengths) / self._size) if self._size else 0.0
        document_frequency: Counter[str] = Counter()
        for tokens in self._tokens:
            document_frequency.update(set(tokens))
        self._idf = {
            term: math.log(1 + (self._size - freq + 0.5) / (freq + 0.5))
            for term, freq in document_frequency.items()
        }

    def search(self, query: str, limit: int = 5, min_score: float = 0.0) -> list[tuple[int, float]]:
        if not self._size or limit <= 0:
            return []
        query_terms = tokenize(query)
        if not query_terms:
            return []
        scores: list[tuple[int, float]] = []
        for index, counts in enumerate(self._counts):
            length = self._lengths[index] or 1
            score = 0.0
            for term in set(query_terms):
                frequency = counts.get(term)
                if not frequency:
                    continue
                idf = self._idf.get(term, 0.0)
                denominator = frequency + self.k1 * (1 - self.b + self.b * length / (self._avgdl or 1))
                score += idf * frequency * (self.k1 + 1) / denominator
            if score > min_score:
                scores.append((index, score))
        scores.sort(key=lambda item: item[1], reverse=True)
        return scores[:limit]


class CaseIndex:
    """Поиск похожих тест-кейсов среди выгрузки DoQA."""

    def __init__(self, cases: list[ExportedCase]) -> None:
        self.cases = cases
        self._index = BM25Index([case.search_text for case in cases])

    def search(self, query: str, limit: int = 5) -> list[tuple[ExportedCase, float]]:
        hits = self._index.search(query, limit=limit)
        return [(self.cases[index], score) for index, score in hits]


class TextIndex:
    """Поиск по фрагментам требований."""

    def __init__(self, chunks: list[RequirementChunk]) -> None:
        self.chunks = chunks
        self._index = BM25Index([f"{chunk.heading}\n{chunk.text}" for chunk in chunks])

    def search(self, query: str, limit: int = 3) -> list[tuple[RequirementChunk, float]]:
        hits = self._index.search(query, limit=limit)
        return [(self.chunks[index], score) for index, score in hits]


@dataclass
class SearchHit:
    case: ExportedCase
    score: float
