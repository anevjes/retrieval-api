from __future__ import annotations

from collections.abc import Awaitable

import pytest

from sharepoint_retrieval_agent.agent import SharePointAnswerAgent
from sharepoint_retrieval_agent.models import (
    RetrievalExtract,
    RetrievalHit,
)
from sharepoint_retrieval_agent.scope import SharePointScope


class FakeRetriever:
    def __init__(self, documents: tuple[RetrievalHit, ...]) -> None:
        self.documents = documents
        self.questions: list[str] = []

    async def retrieve(
        self,
        question: str,
        *,
        scope: SharePointScope,
        maximum_results: int,
    ) -> tuple[RetrievalHit, ...]:
        del scope, maximum_results
        self.questions.append(question)
        return self.documents


class FakeGenerator:
    def __init__(self, answer: str = "Grounded answer [1].") -> None:
        self.answer = answer
        self.prompts: list[str] = []

    def generate(self, prompt: str) -> Awaitable[str]:
        async def response() -> str:
            self.prompts.append(prompt)
            return self.answer

        return response()


@pytest.mark.asyncio
async def test_no_hits_skips_the_llm() -> None:
    retriever = FakeRetriever(())
    synthesizer = FakeGenerator()
    agent = SharePointAnswerAgent(
        retriever=retriever,  # type: ignore[arg-type]
        synthesizer=synthesizer,  # type: ignore[arg-type]
    )

    answer = await agent.answer(
        "Where is the policy?",
        scope=SharePointScope.all_accessible_sites(),
    )

    assert "couldn't find relevant information" in answer.text
    assert synthesizer.prompts == []


@pytest.mark.asyncio
async def test_agent_stitches_retrieval_hits_into_a_cited_prompt() -> None:
    hit = RetrievalHit(
        web_url="https://contoso.sharepoint.com/sites/HR/policy.docx",
        title="Policy",
        extracts=(RetrievalExtract("Employees receive 20 days.", 0.9),),
    )
    retriever = FakeRetriever((hit,))
    synthesizer = FakeGenerator("Employees receive 20 days [1].")
    agent = SharePointAnswerAgent(
        retriever=retriever,  # type: ignore[arg-type]
        synthesizer=synthesizer,  # type: ignore[arg-type]
    )

    answer = await agent.answer(
        "How much leave do employees receive?",
        scope=SharePointScope.all_accessible_sites(),
    )

    assert answer.sources[0].web_url == hit.web_url
    assert "Employees receive 20 days." in synthesizer.prompts[0]
    assert "[1]" in answer.render_markdown()
