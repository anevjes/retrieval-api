from __future__ import annotations

from collections.abc import Awaitable

import pytest

from sharepoint_retrieval_agent.agent import SharePointAnswerAgent
from sharepoint_retrieval_agent.models import (
    RetrievalExtract,
    RetrievalHit,
    RetrievalResult,
)
from sharepoint_retrieval_agent.scope import SharePointScope


class FakeRetriever:
    def __init__(self, result: RetrievalResult) -> None:
        self.result = result
        self.questions: list[str] = []

    async def retrieve(
        self,
        question: str,
        *,
        scope: SharePointScope,
        maximum_results: int,
    ) -> RetrievalResult:
        del scope, maximum_results
        self.questions.append(question)
        return self.result


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
    retriever = FakeRetriever(RetrievalResult(()))
    generator = FakeGenerator()
    agent = SharePointAnswerAgent(retriever=retriever, generator=generator)  # type: ignore[arg-type]

    answer = await agent.answer(
        "Where is the policy?",
        scope=SharePointScope.all_accessible_sites(),
    )

    assert "couldn't find relevant information" in answer.text
    assert generator.prompts == []


@pytest.mark.asyncio
async def test_agent_stitches_retrieval_hits_into_a_cited_prompt() -> None:
    hit = RetrievalHit(
        web_url="https://contoso.sharepoint.com/sites/HR/policy.docx",
        metadata={"title": "Policy"},
        extracts=(RetrievalExtract("Employees receive 20 days.", 0.9),),
    )
    retriever = FakeRetriever(RetrievalResult((hit,)))
    generator = FakeGenerator("Employees receive 20 days [1].")
    agent = SharePointAnswerAgent(retriever=retriever, generator=generator)  # type: ignore[arg-type]

    answer = await agent.answer(
        "How much leave do employees receive?",
        scope=SharePointScope.all_accessible_sites(),
    )

    assert answer.sources[0].web_url == hit.web_url
    assert "Employees receive 20 days." in generator.prompts[0]
    assert "[1]" in answer.render_markdown()
