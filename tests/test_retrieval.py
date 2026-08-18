from __future__ import annotations

import json
import logging

import httpx
import pytest

from sharepoint_retrieval_agent.retrieval import (
    CopilotRetrievalClient,
    CopilotRetrievalError,
)
from sharepoint_retrieval_agent.scope import SharePointScope


async def get_access_token() -> str:
    return "delegated-token"


@pytest.mark.asyncio
async def test_request_is_sharepoint_only_and_results_are_post_filtered(caplog) -> None:
    captured: dict[str, object] = {}
    request_count = 0

    def handler(request: httpx.Request) -> httpx.Response:
        nonlocal request_count
        request_count += 1
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
    client = CopilotRetrievalClient(get_access_token, http_client=http_client)
    scope = SharePointScope.selected_sites(["https://contoso.sharepoint.com/sites/HR/"])

    with caplog.at_level(logging.INFO, logger="sharepoint_retrieval_agent.retrieval"):
        result = await client.retrieve("What is the policy?", scope=scope, maximum_results=25)
    await http_client.aclose()

    assert captured["dataSource"] == "sharePoint"
    assert captured["filterExpression"] == 'Path:"https://contoso.sharepoint.com/sites/HR/"'
    assert "oneDriveBusiness" not in json.dumps(captured)
    assert "externalItem" not in json.dumps(captured)
    assert request_count == 1
    assert len(result) == 1
    assert result[0].title == "Policy"
    assert result[0].author == "Adele"
    assert "[2/4] Retrieval: POST" in caplog.text
    assert 'filterExpression=Path:"https://contoso.sharepoint.com/sites/HR/"' in caplog.text
    assert "returned=4; accepted=1; discarded=3" in caplog.text
    assert "delegated-token" not in caplog.text
    assert "Approved policy text" not in caplog.text


@pytest.mark.asyncio
async def test_all_sites_omits_filter_expression() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        body = json.loads(request.content)
        assert "filterExpression" not in body
        return httpx.Response(200, json={"retrievalHits": []})

    http_client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    client = CopilotRetrievalClient(get_access_token, http_client=http_client)

    result = await client.retrieve(
        "Find the travel policy.",
        scope=SharePointScope.all_accessible_sites(),
    )
    await http_client.aclose()

    assert result == ()


@pytest.mark.asyncio
async def test_empty_results_explain_indexing_and_query_diagnostics(caplog) -> None:
    def handler(_: httpx.Request) -> httpx.Response:
        return httpx.Response(200, json={"retrievalHits": []})

    http_client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    client = CopilotRetrievalClient(get_access_token, http_client=http_client)

    with caplog.at_level(logging.INFO, logger="sharepoint_retrieval_agent.retrieval"):
        result = await client.retrieve(
            "Western Australia",
            scope=SharePointScope.selected_sites(
                ["https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/"]
            ),
        )
    await http_client.aclose()

    assert result == ()
    assert "query is short" in caplog.text
    assert "indexed daily" in caplog.text
    assert "allowed to appear in search" in caplog.text


@pytest.mark.asyncio
async def test_graph_error_is_sanitized_and_preserves_request_id() -> None:
    def handler(_: httpx.Request) -> httpx.Response:
        return httpx.Response(
            403,
            headers={"request-id": "request-denied"},
            json={"error": {"message": "Insufficient delegated permissions."}},
        )

    http_client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    client = CopilotRetrievalClient(get_access_token, http_client=http_client)

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
