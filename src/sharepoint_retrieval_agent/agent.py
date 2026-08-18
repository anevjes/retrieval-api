"""Deterministic retrieve-then-synthesize SharePoint answer agent."""

from __future__ import annotations

import logging

from .llm import AnswerSynthesizer
from .models import GroundedAnswer
from .retrieval import CopilotRetrievalClient
from .scope import SharePointScope
from .synthesis import build_grounding_prompt

_NO_RESULTS_MESSAGE = (
    "I couldn't find relevant information in the permitted SharePoint content for this question."
)

logger = logging.getLogger(__name__)


class SharePointAnswerAgent:
    """Retrieve SharePoint documents, then synthesize only from validated extracts."""

    def __init__(
        self,
        *,
        retriever: CopilotRetrievalClient,
        synthesizer: AnswerSynthesizer,
        maximum_results: int = 25,
    ) -> None:
        self._retriever = retriever
        self._synthesizer = synthesizer
        self._maximum_results = maximum_results

    async def answer(self, question: str, *, scope: SharePointScope) -> GroundedAnswer:
        """Answer a question using only permission-trimmed SharePoint extracts."""

        documents = await self._retriever.retrieve(
            question,
            scope=scope,
            maximum_results=self._maximum_results,
        )
        if not documents:
            logger.info("[3/4] Grounding: no validated SharePoint sources; skipping synthesis.")
            return GroundedAnswer(text=_NO_RESULTS_MESSAGE)

        logger.info(
            "[3/4] Grounding: serializing %s validated SharePoint source(s) as untrusted data.",
            len(documents),
        )
        prompt, sources = build_grounding_prompt(question, documents)
        text = await self._synthesizer.generate(prompt)
        return GroundedAnswer(text=text, sources=sources)
