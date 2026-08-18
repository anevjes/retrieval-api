"""Deterministic retrieve-then-synthesize SharePoint answer agent."""

from __future__ import annotations

from .llm import TextGenerator
from .models import GroundedAnswer
from .retrieval import CopilotRetrievalClient, normalize_query
from .scope import SharePointScope
from .synthesis import build_grounding_prompt

_NO_RESULTS_MESSAGE = (
    "I couldn't find relevant information in the permitted SharePoint content for this question."
)


class SharePointAnswerAgent:
    """Run SharePoint retrieval first, then synthesize from validated extracts only."""

    def __init__(
        self,
        *,
        retriever: CopilotRetrievalClient,
        generator: TextGenerator,
        maximum_results: int = 25,
        maximum_context_characters: int = 120_000,
    ) -> None:
        self._retriever = retriever
        self._generator = generator
        self._maximum_results = maximum_results
        self._maximum_context_characters = maximum_context_characters

    async def answer(self, question: str, *, scope: SharePointScope) -> GroundedAnswer:
        """Answer a question using only permission-trimmed SharePoint extracts."""

        normalized_question = normalize_query(question)
        retrieval = await self._retriever.retrieve(
            normalized_question,
            scope=scope,
            maximum_results=self._maximum_results,
        )
        if not retrieval.hits:
            return GroundedAnswer(
                text=_NO_RESULTS_MESSAGE,
                discarded_out_of_scope=retrieval.discarded_out_of_scope,
            )

        prompt, sources, truncated = build_grounding_prompt(
            normalized_question,
            retrieval.hits,
            maximum_context_characters=self._maximum_context_characters,
        )
        if not sources:
            return GroundedAnswer(
                text=_NO_RESULTS_MESSAGE,
                discarded_out_of_scope=retrieval.discarded_out_of_scope,
                context_was_truncated=truncated,
            )

        text = await self._generator.generate(prompt)
        return GroundedAnswer(
            text=text,
            sources=sources,
            discarded_out_of_scope=retrieval.discarded_out_of_scope,
            context_was_truncated=truncated,
        )
