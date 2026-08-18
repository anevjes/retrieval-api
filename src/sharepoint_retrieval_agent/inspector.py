"""Optional tool-calling variant for Foundry Toolkit Agent Inspector."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass

from agent_framework import Agent, tool

from .auth import MsalDeviceCodeTokenProvider
from .config import GraphSettings, LLMSettings, RetrievalSettings
from .llm import SYNTHESIS_INSTRUCTIONS, AzureCredential, build_chat_client
from .retrieval import CopilotRetrievalClient
from .scope import SharePointScope
from .synthesis import order_sources, sources_as_json


@dataclass(slots=True)
class InspectorResources:
    """Keep resources alive for the lifetime of the local Inspector server."""

    graph_auth: MsalDeviceCodeTokenProvider
    retriever: CopilotRetrievalClient
    model_credential: AzureCredential | None


def build_inspector_agent(
    *,
    scope: SharePointScope,
    graph_settings: GraphSettings,
    llm_settings: LLMSettings,
    retrieval_settings: RetrievalSettings,
) -> tuple[Agent, InspectorResources]:
    """Build a tool-calling agent for interactive local inspection."""

    graph_auth = MsalDeviceCodeTokenProvider(
        tenant_id=graph_settings.tenant_id,
        client_id=graph_settings.client_id,
    )
    retriever = CopilotRetrievalClient(graph_auth.get_token)
    client, model_credential = build_chat_client(llm_settings)

    @tool
    async def retrieve_sharepoint_documents(query: str) -> str:
        """Retrieve relevant documents only from the configured SharePoint site scope."""

        documents = await retriever.retrieve(
            query,
            scope=scope,
            maximum_results=retrieval_settings.maximum_results,
        )
        if not documents:
            return json.dumps({"status": "no_results", "sources": []})

        sources = order_sources(documents)
        return json.dumps(
            {
                "status": "ok",
                "sources": json.loads(sources_as_json(sources)),
            },
            ensure_ascii=False,
        )

    instructions = f"""{SYNTHESIS_INSTRUCTIONS}
For every user question, call retrieve_sharepoint_documents exactly once using a focused,
single-sentence query. The tool is the only knowledge source available. If it returns no_results,
say that no relevant information was found in permitted SharePoint content. Never answer before
calling it. Do not reproduce raw JSON or fabricate citations.
"""
    agent = Agent(
        client=client,
        name="SharePointRetrievalAgent",
        description="Answers questions from SharePoint through the Copilot Retrieval API.",
        instructions=instructions,
        tools=[retrieve_sharepoint_documents],
        default_options={"store": False},
    )
    return agent, InspectorResources(graph_auth, retriever, model_credential)


def run_inspector_server(
    *,
    port: int,
    scope: SharePointScope,
    graph_settings: GraphSettings,
    llm_settings: LLMSettings,
    retrieval_settings: RetrievalSettings,
) -> None:
    """Serve the agent on localhost for F5 and Foundry Toolkit Agent Inspector."""

    try:
        from agent_framework_foundry_hosting import (
            ResponsesHostServer,
        )
    except ImportError as error:  # pragma: no cover - depends on optional dev package
        raise RuntimeError("Install requirements.txt to use Inspector mode.") from error

    agent, resources = build_inspector_agent(
        scope=scope,
        graph_settings=graph_settings,
        llm_settings=llm_settings,
        retrieval_settings=retrieval_settings,
    )
    # Keep resources referenced while the blocking development server runs.
    _ACTIVE_RESOURCES.append(resources)
    os.environ["PORT"] = str(port)
    server = ResponsesHostServer(agent)
    server.run(host="127.0.0.1", port=port)


_ACTIVE_RESOURCES: list[InspectorResources] = []
