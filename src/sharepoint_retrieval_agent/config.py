"""Environment-based configuration with explicit validation."""

from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass, field
from typing import Literal
from urllib.parse import urlsplit


class ConfigurationError(ValueError):
    """Raised when required configuration is missing or unsafe."""


def _value(environment: Mapping[str, str], name: str) -> str | None:
    value = environment.get(name, "").strip()
    return value or None


def _required(environment: Mapping[str, str], name: str) -> str:
    value = _value(environment, name)
    if value is None:
        raise ConfigurationError(f"Set {name} in the environment or .env file.")
    return value


def parse_bool(value: str | None, *, default: bool = False) -> bool:
    if value is None or not value.strip():
        return default
    normalized = value.strip().casefold()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    raise ConfigurationError(f"Expected a boolean value, received {value!r}.")


@dataclass(frozen=True, slots=True)
class GraphSettings:
    """Microsoft Graph delegated authentication settings."""

    tenant_id: str
    client_id: str

    @classmethod
    def from_environment(cls, environment: Mapping[str, str] | None = None) -> GraphSettings:
        values = os.environ if environment is None else environment
        return cls(
            tenant_id=_required(values, "GRAPH_TENANT_ID"),
            client_id=_required(values, "GRAPH_CLIENT_ID"),
        )


@dataclass(frozen=True, slots=True)
class LLMSettings:
    """Configuration for the Agent Framework synthesis model."""

    provider: Literal["azure_openai", "openai"]
    model: str
    azure_base_url: str | None = None
    azure_auth_mode: Literal["azure_cli", "managed_identity", "api_key"] = "azure_cli"
    api_key: str | None = field(default=None, repr=False)
    managed_identity_client_id: str | None = None

    @classmethod
    def from_environment(cls, environment: Mapping[str, str] | None = None) -> LLMSettings:
        values = os.environ if environment is None else environment
        provider = (_value(values, "LLM_PROVIDER") or "azure_openai").casefold()

        if provider == "azure_openai":
            base_url = _required(values, "AZURE_OPENAI_BASE_URL")
            parsed = urlsplit(base_url)
            if parsed.scheme.casefold() != "https" or not parsed.netloc:
                raise ConfigurationError("AZURE_OPENAI_BASE_URL must be an HTTPS URL.")
            if not parsed.path.rstrip("/").endswith("/openai/v1"):
                raise ConfigurationError(
                    "AZURE_OPENAI_BASE_URL must end with /openai/v1/ for the Azure OpenAI v1 API."
                )

            auth_mode = (_value(values, "AZURE_OPENAI_AUTH_MODE") or "azure_cli").casefold()
            if auth_mode not in {"azure_cli", "managed_identity", "api_key"}:
                raise ConfigurationError(
                    "AZURE_OPENAI_AUTH_MODE must be azure_cli, managed_identity, or api_key."
                )
            api_key = _value(values, "AZURE_OPENAI_API_KEY")
            if auth_mode == "api_key" and api_key is None:
                raise ConfigurationError(
                    "Set AZURE_OPENAI_API_KEY when AZURE_OPENAI_AUTH_MODE is api_key."
                )
            return cls(
                provider="azure_openai",
                model=_required(values, "AZURE_OPENAI_MODEL"),
                azure_base_url=base_url.rstrip("/") + "/",
                azure_auth_mode=auth_mode,  # type: ignore[arg-type]
                api_key=api_key,
                managed_identity_client_id=_value(values, "AZURE_CLIENT_ID"),
            )

        if provider == "openai":
            return cls(
                provider="openai",
                model=_value(values, "OPENAI_MODEL") or "gpt-5.4-mini",
                api_key=_required(values, "OPENAI_API_KEY"),
            )

        raise ConfigurationError("LLM_PROVIDER must be azure_openai or openai.")


@dataclass(frozen=True, slots=True)
class RuntimeSettings:
    """Retrieval and context-size settings."""

    maximum_results: int = 25
    maximum_context_characters: int = 120_000

    @classmethod
    def from_environment(cls, environment: Mapping[str, str] | None = None) -> RuntimeSettings:
        values = os.environ if environment is None else environment
        try:
            maximum_results = int(_value(values, "RETRIEVAL_MAX_RESULTS") or "25")
            maximum_context = int(_value(values, "MAX_CONTEXT_CHARACTERS") or "120000")
        except ValueError as error:
            raise ConfigurationError("Retrieval limits must be integers.") from error
        if not 1 <= maximum_results <= 25:
            raise ConfigurationError("RETRIEVAL_MAX_RESULTS must be between 1 and 25.")
        if maximum_context < 4_000:
            raise ConfigurationError("MAX_CONTEXT_CHARACTERS must be at least 4000.")
        return cls(maximum_results, maximum_context)


def configured_site_urls(environment: Mapping[str, str] | None = None) -> tuple[str, ...]:
    """Read semicolon-separated default SharePoint site paths."""

    values = os.environ if environment is None else environment
    raw = _value(values, "SHAREPOINT_SITE_URLS")
    if raw is None:
        return ()
    return tuple(item.strip() for item in raw.split(";") if item.strip())
