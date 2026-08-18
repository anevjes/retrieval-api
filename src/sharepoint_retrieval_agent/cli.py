"""Command-line and local Inspector entry points."""

from __future__ import annotations

import argparse
import asyncio
import logging
import os
import sys
from collections.abc import Sequence

from dotenv import load_dotenv

from .agent import SharePointAnswerAgent
from .auth import MsalDeviceCodeTokenProvider
from .config import (
    ConfigurationError,
    GraphSettings,
    LLMSettings,
    RetrievalSettings,
    configured_site_urls,
    parse_bool,
)
from .inspector import run_inspector_server
from .llm import AnswerSynthesizer
from .retrieval import CopilotRetrievalClient, CopilotRetrievalError
from .scope import ScopeValidationError, SharePointScope

logger = logging.getLogger(__name__)

_AUTH_SDK_LOGGERS = (
    "azure.identity",
    "azure.core.pipeline.policies.http_logging_policy",
    "msal",
)


def _add_logging_arguments(
    parser: argparse.ArgumentParser,
    *,
    suppress_defaults: bool = False,
) -> None:
    default = argparse.SUPPRESS if suppress_defaults else False
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        default=default,
        help="Show the high-level authentication, retrieval, validation, and synthesis flow.",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        default=default,
        help="Show verbose flow logs and exception tracebacks.",
    )


def _add_scope_arguments(parser: argparse.ArgumentParser) -> None:
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--site",
        action="append",
        metavar="URL",
        help="Restrict retrieval to this SharePoint site path; repeat for multiple sites.",
    )
    group.add_argument(
        "--all-sites",
        action="store_true",
        help="Search all SharePoint sites the signed-in user can access.",
    )


def _add_retrieval_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--max-results", type=int, help="Retrieval result limit (1-25).")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sharepoint-agent",
        description="Answer questions from SharePoint via the Microsoft 365 Copilot Retrieval API.",
    )
    _add_logging_arguments(parser)
    subparsers = parser.add_subparsers(dest="command", required=True)

    ask = subparsers.add_parser("ask", help="Ask one question and exit.")
    _add_logging_arguments(ask, suppress_defaults=True)
    ask.add_argument("question", help="The natural-language question to answer.")
    _add_scope_arguments(ask)
    _add_retrieval_arguments(ask)

    chat = subparsers.add_parser("chat", help="Start an interactive question loop.")
    _add_logging_arguments(chat, suppress_defaults=True)
    _add_scope_arguments(chat)
    _add_retrieval_arguments(chat)

    inspector = subparsers.add_parser(
        "inspector",
        help="Host the tool-calling agent for Foundry Toolkit Agent Inspector.",
    )
    _add_logging_arguments(inspector, suppress_defaults=True)
    inspector.add_argument(
        "--port",
        type=int,
        default=int(os.getenv("INSPECTOR_PORT", "8088")),
        help="Local Inspector server port.",
    )
    _add_scope_arguments(inspector)
    _add_retrieval_arguments(inspector)
    return parser


def _scope_from_arguments(arguments: argparse.Namespace) -> SharePointScope:
    if arguments.all_sites:
        return SharePointScope.all_accessible_sites()
    if arguments.site:
        return SharePointScope.selected_sites(arguments.site)

    site_urls = configured_site_urls()
    if site_urls:
        return SharePointScope.selected_sites(site_urls)
    if parse_bool(os.getenv("ALLOW_ALL_SITES"), default=False):
        return SharePointScope.all_accessible_sites()
    raise ConfigurationError(
        "Choose --site, choose --all-sites, set SHAREPOINT_SITE_URLS, or explicitly set "
        "ALLOW_ALL_SITES=true."
    )


def _retrieval_settings(arguments: argparse.Namespace) -> RetrievalSettings:
    configured = RetrievalSettings.from_environment()
    maximum_results = arguments.max_results or configured.maximum_results
    if not 1 <= maximum_results <= 25:
        raise ConfigurationError("--max-results must be between 1 and 25.")
    return RetrievalSettings(maximum_results)


async def _run_questions(
    arguments: argparse.Namespace,
    *,
    scope: SharePointScope,
    graph_settings: GraphSettings,
    llm_settings: LLMSettings,
    retrieval_settings: RetrievalSettings,
) -> None:
    graph_auth = MsalDeviceCodeTokenProvider(
        tenant_id=graph_settings.tenant_id,
        client_id=graph_settings.client_id,
    )
    retriever = CopilotRetrievalClient(graph_auth.get_token)
    synthesizer = AnswerSynthesizer(llm_settings)
    agent = SharePointAnswerAgent(
        retriever=retriever,
        synthesizer=synthesizer,
        maximum_results=retrieval_settings.maximum_results,
    )

    async def ask(question: str) -> None:
        answer = await agent.answer(question, scope=scope)
        print(answer.render_markdown())

    try:
        if arguments.command == "ask":
            await ask(arguments.question)
            return

        print("SharePoint agent ready. Type 'exit' to quit.")
        while True:
            try:
                question = (await asyncio.to_thread(input, "\nYou: ")).strip()
            except EOFError:
                break
            if question.casefold() in {"exit", "quit", ":q"}:
                break
            if question:
                await ask(question)
    finally:
        await retriever.close()
        await synthesizer.close()


def _configure_logging(*, verbose: bool, debug: bool) -> None:
    app_level = logging.DEBUG if debug else logging.INFO if verbose else logging.WARNING
    logging.basicConfig(level=logging.WARNING, format="%(levelname)s %(name)s: %(message)s")
    logging.getLogger("sharepoint_retrieval_agent").setLevel(app_level)

    # Hide expected MSAL device-code polling responses unless diagnostics are requested.
    auth_level_name = os.getenv("AUTH_SDK_LOG_LEVEL", "WARNING").upper()
    auth_level = getattr(logging, auth_level_name, logging.WARNING)
    for logger_name in _AUTH_SDK_LOGGERS:
        logging.getLogger(logger_name).setLevel(auth_level)


def _translated_argv(argv: Sequence[str]) -> list[str]:
    values = list(argv)
    # Compatibility with the agentdev debugging template.
    if values and values[0] == "--server":
        return ["inspector", *values[1:]]
    return values


def run(argv: Sequence[str] | None = None) -> int:
    load_dotenv()
    parser = build_parser()
    arguments = parser.parse_args(_translated_argv(sys.argv[1:] if argv is None else argv))
    _configure_logging(verbose=arguments.verbose, debug=arguments.debug)

    try:
        scope = _scope_from_arguments(arguments)
        graph_settings = GraphSettings.from_environment()
        llm_settings = LLMSettings.from_environment()
        retrieval_settings = _retrieval_settings(arguments)

        scope_name = (
            "all accessible SharePoint sites"
            if scope.includes_all_accessible_sites
            else f"{len(scope.site_urls)} selected SharePoint site(s)"
        )
        logger.info(
            "CLI flow: command=%s; scope=%s; maximum_results=%s; llm_provider=%s; model=%s.",
            arguments.command,
            scope_name,
            retrieval_settings.maximum_results,
            llm_settings.provider,
            llm_settings.model,
        )

        if arguments.command == "inspector":
            run_inspector_server(
                port=arguments.port,
                scope=scope,
                graph_settings=graph_settings,
                llm_settings=llm_settings,
                retrieval_settings=retrieval_settings,
            )
        else:
            asyncio.run(
                _run_questions(
                    arguments,
                    scope=scope,
                    graph_settings=graph_settings,
                    llm_settings=llm_settings,
                    retrieval_settings=retrieval_settings,
                )
            )
        return 0
    except KeyboardInterrupt:
        print("\nCancelled.", file=sys.stderr)
        return 130
    except (
        ConfigurationError,
        CopilotRetrievalError,
        ScopeValidationError,
        ValueError,
        RuntimeError,
    ) as error:
        if arguments.debug:
            logger.exception("Agent failed")
        else:
            print(f"Error: {error}", file=sys.stderr)
        return 2


def main() -> None:
    raise SystemExit(run())
