"""Client for the stable Microsoft 365 Copilot Retrieval API."""

from __future__ import annotations

import logging
from collections.abc import Mapping
from pathlib import PurePosixPath
from urllib.parse import unquote, urlsplit

import httpx

from .auth import MsalDeviceCodeTokenProvider
from .models import RetrievalExtract, RetrievalHit
from .scope import SharePointScope

logger = logging.getLogger(__name__)

RETRIEVAL_ENDPOINT = "https://graph.microsoft.com/v1.0/copilot/retrieval"


class CopilotRetrievalError(RuntimeError):
    """Raised when Microsoft Graph cannot complete a retrieval request."""

    def __init__(
        self,
        message: str,
        *,
        status_code: int | None = None,
        request_id: str | None = None,
    ) -> None:
        super().__init__(message)
        self.status_code = status_code
        self.request_id = request_id


def normalize_query(question: str) -> str:
    """Normalize and validate a Retrieval API natural-language query."""

    normalized = " ".join(question.split())
    if not normalized:
        raise ValueError("The question cannot be empty.")
    if len(normalized) > 1_500:
        raise ValueError("The Retrieval API query cannot exceed 1,500 characters.")
    return normalized


def _error_message(response: httpx.Response) -> str:
    try:
        payload = response.json()
        error = payload.get("error", {}) if isinstance(payload, dict) else {}
        message = error.get("message") if isinstance(error, dict) else None
        if isinstance(message, str) and message.strip():
            return message.strip()
    except ValueError:
        pass
    return f"Microsoft Graph returned HTTP {response.status_code}."


def _score(value: object) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, int | float):
        return float(value)
    return None


def _parse_hit(raw_hit: object) -> RetrievalHit | None:
    if not isinstance(raw_hit, Mapping):
        return None
    web_url = raw_hit.get("webUrl")
    if not isinstance(web_url, str) or not web_url.strip():
        return None

    extracts: list[RetrievalExtract] = []
    raw_extracts = raw_hit.get("extracts", [])
    if isinstance(raw_extracts, list):
        for raw_extract in raw_extracts:
            if not isinstance(raw_extract, Mapping):
                continue
            text = raw_extract.get("text")
            if isinstance(text, str) and text.strip():
                extracts.append(
                    RetrievalExtract(
                        text=text.strip(),
                        relevance_score=_score(raw_extract.get("relevanceScore")),
                    )
                )

    resource_type = raw_hit.get("resourceType")
    if isinstance(resource_type, str) and resource_type.casefold() != "listitem":
        return None

    metadata = raw_hit.get("resourceMetadata")
    metadata = metadata if isinstance(metadata, Mapping) else {}
    title = metadata.get("title")
    if not isinstance(title, str) or not title.strip():
        title = PurePosixPath(unquote(urlsplit(web_url).path)).name or web_url
    author = metadata.get("author")
    return RetrievalHit(
        web_url=web_url.strip(),
        title=title.strip(),
        extracts=tuple(extracts),
        author=author.strip() if isinstance(author, str) and author.strip() else None,
    )


class CopilotRetrievalClient:
    """Retrieve SharePoint extracts while preventing other M365 data sources."""

    def __init__(
        self,
        token_provider: MsalDeviceCodeTokenProvider,
        *,
        http_client: httpx.AsyncClient | None = None,
    ) -> None:
        self._token_provider = token_provider
        self._owns_http_client = http_client is None
        self._http_client = http_client or httpx.AsyncClient(
            timeout=httpx.Timeout(30.0, connect=10.0),
            headers={"User-Agent": "sharepoint-retrieval-agent/0.1"},
        )

    async def close(self) -> None:
        if self._owns_http_client:
            await self._http_client.aclose()

    async def retrieve(
        self,
        question: str,
        *,
        scope: SharePointScope,
        maximum_results: int = 25,
    ) -> tuple[RetrievalHit, ...]:
        """Call the Retrieval API once and return validated SharePoint documents."""

        if not 1 <= maximum_results <= 25:
            raise ValueError("maximum_results must be between 1 and 25.")

        payload: dict[str, object] = {
            "queryString": normalize_query(question),
            # Intentionally fixed: OneDrive, connectors, mail, and Teams are never requested.
            "dataSource": "sharePoint",
            "resourceMetadata": ["title", "author"],
            "maximumNumberOfResults": maximum_results,
        }
        if scope.filter_expression is not None:
            payload["filterExpression"] = scope.filter_expression

        token = await self._token_provider.get_token()
        try:
            response = await self._http_client.post(
                RETRIEVAL_ENDPOINT,
                headers={"Authorization": f"Bearer {token}"},
                json=payload,
            )
        except httpx.RequestError as error:
            raise CopilotRetrievalError(f"Could not reach Microsoft Graph: {error}") from error

        request_id = response.headers.get("request-id") or response.headers.get("client-request-id")
        if not response.is_success:
            raise CopilotRetrievalError(
                _error_message(response),
                status_code=response.status_code,
                request_id=request_id,
            )

        try:
            body = response.json()
        except ValueError as error:
            raise CopilotRetrievalError(
                "Microsoft Graph returned invalid JSON.",
                status_code=response.status_code,
                request_id=request_id,
            ) from error
        if not isinstance(body, Mapping) or not isinstance(body.get("retrievalHits", []), list):
            raise CopilotRetrievalError(
                "Microsoft Graph returned an unexpected Retrieval API response.",
                status_code=response.status_code,
                request_id=request_id,
            )

        documents: list[RetrievalHit] = []
        discarded = 0
        for raw_hit in body.get("retrievalHits", []):
            hit = _parse_hit(raw_hit)
            # The KQL filter runs in Microsoft Graph. This URL check is defense in depth because an
            # invalid KQL expression can run unscoped. Unvalidated content never reaches the LLM.
            if hit is None or not hit.extracts or not scope.allows(hit.web_url):
                discarded += 1
                continue
            documents.append(hit)

        if discarded:
            logger.warning(
                "Discarded %s result(s) that were not valid for the SharePoint scope.",
                discarded,
            )
        return tuple(documents)
