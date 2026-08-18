from __future__ import annotations

import json
from collections.abc import Awaitable

import httpx
import pytest

from sharepoint_retrieval_agent.retrieval import (
    CopilotRetrievalClient,
    CopilotRetrievalError,
)
from sharepoint_retrieval_agent.scope import SharePointScope


class FakeTokenProvider:
    def get_token(self) -> Awaitable[str]:
        async def token() -> str:
            return "delegated-token"

        return token()


@pytest.mark.asyncio
async def test_request_is_sharepoint_only_and_results_are_post_filtered() -> None:
    captured: dict[str, object] = {}

    def handler(request: httpx.Request) -> httpx.Response:
        captured.update(json.loads(request.content))
        assert request.headers["Authorization"] == "Bearer delegated-token"
        return httpx.Response(
            200,
            headers={"request-id": "request-123"},
            json={
                "retrievalHits": [
                    {
                        "webUrl": "https://contoso.sharepoint.com/sites/HR/policy.docx",
                        "resourceType": "listItem",
                        "resourceMetadata": {"title": "Policy", "author": "Adele"},
                        "extracts": [{"text": "Approved policy text", "relevanceScore": 0.9}],
                    },
                    {
                        "webUrl": "https://contoso.sharepoint.com/sites/Finance/budget.docx",
                        "extracts": [{"text": "Out of scope", "relevanceScore": 0.99}],
                    },
                    {
                        "webUrl": "https://contoso-my.sharepoint.com/personal/alex/file.docx",
                        "extracts": [{"text": "OneDrive", "relevanceScore": 1.0}],
                    },
                    {
                        "webUrl": "https://contoso.sharepoint.com/sites/HR/spoofed.docx",
                        "resourceType": "externalItem",
                        "extracts": [{"text": "Connector content", "relevanceScore": 1.0}],
                    },
                ]
            },
        )

    http_client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    client = CopilotRetrievalClient(
        FakeTokenProvider(),
        http_client=http_client,
        maximum_attempts=1,
    )
    scope = SharePointScope.selected_sites(["https://contoso.sharepoint.com/sites/HR/"])

    result = await client.retrieve("What is the policy?", scope=scope, maximum_results=25)
    await http_client.aclose()

    assert captured["dataSource"] == "sharePoint"
    assert captured["filterExpression"] == 'Path:"https://contoso.sharepoint.com/sites/HR/"'
    assert "oneDriveBusiness" not in json.dumps(captured)
    assert "externalItem" not in json.dumps(captured)
    assert len(result.hits) == 1
    assert result.hits[0].metadata["title"] == "Policy"
    assert result.discarded_out_of_scope == 3
    assert result.request_id == "request-123"


@pytest.mark.asyncio
async def test_all_sites_omits_filter_expression() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content)
        assert "filterExpression" not in body
        return httpx.Response(200, json={"retrievalHits": []})

    http_client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    client = CopilotRetrievalClient(FakeTokenProvider(), http_client=http_client)

    result = await client.retrieve(
        "Find the travel policy.",
        scope=SharePointScope.all_accessible_sites(),
    )
    await http_client.aclose()

    assert result.hits == ()


@pytest.mark.asyncio
async def test_graph_error_is_sanitized_and_preserves_request_id() -> None:
    def handler(_: httpx.Request) -> httpx.Response:
        return httpx.Response(
            403,
            headers={"request-id": "request-denied"},
            json={"error": {"message": "Insufficient delegated permissions."}},
        )

    http_client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    client = CopilotRetrievalClient(
        FakeTokenProvider(),
        http_client=http_client,
        maximum_attempts=1,
    )

    with pytest.raises(CopilotRetrievalError) as caught:
        await client.retrieve(
            "Find the policy.",
            scope=SharePointScope.all_accessible_sites(),
        )
    await http_client.aclose()

    assert caught.value.status_code == 403
    assert caught.value.request_id == "request-denied"
    assert "Insufficient delegated permissions" in str(caught.value)


def test_query_length_is_enforced_before_network_call() -> None:
    from sharepoint_retrieval_agent.retrieval import normalize_query

    with pytest.raises(ValueError, match="1,500"):
        normalize_query("x" * 1_501)
