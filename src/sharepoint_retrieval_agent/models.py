"""Domain models for retrieval and grounded answers."""

from __future__ import annotations

import re
from dataclasses import dataclass

_CITATION_PATTERN = re.compile(r"\[(\d+)]")


@dataclass(frozen=True, slots=True)
class RetrievalExtract:
    """A text extract returned by the Microsoft 365 Copilot Retrieval API."""

    text: str
    relevance_score: float | None = None


@dataclass(frozen=True, slots=True)
class RetrievalHit:
    """A validated SharePoint document and its relevant text extracts."""

    web_url: str
    title: str
    extracts: tuple[RetrievalExtract, ...]
    author: str | None = None


@dataclass(frozen=True, slots=True)
class GroundedAnswer:
    """A synthesized answer and the deterministic source list behind it."""

    text: str
    sources: tuple[RetrievalHit, ...] = ()

    @property
    def cited_source_indexes(self) -> tuple[int, ...]:
        """Return valid source indexes cited by the generated answer."""

        maximum = len(self.sources)
        indexes = {
            int(match.group(1))
            for match in _CITATION_PATTERN.finditer(self.text)
            if 1 <= int(match.group(1)) <= maximum
        }
        return tuple(sorted(indexes))

    def render_markdown(self) -> str:
        """Render the answer with a source list controlled by application code."""

        maximum = len(self.sources)
        body = _CITATION_PATTERN.sub(
            lambda match: match.group(0) if 1 <= int(match.group(1)) <= maximum else "",
            self.text,
        ).strip()
        if not self.sources:
            return body

        cited = set(self.cited_source_indexes)
        selected = [
            (index, source)
            for index, source in enumerate(self.sources, start=1)
            if not cited or index in cited
        ]
        lines = [body, "", "### Sources"]
        for index, source in selected:
            title = source.title.replace("[", "\\[").replace("]", "\\]")
            author = f" — {source.author}" if source.author else ""
            lines.append(f"- [{index}] [{title}]({source.web_url}){author}")
        return "\n".join(lines)
