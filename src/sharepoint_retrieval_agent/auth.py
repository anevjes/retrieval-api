"""MSAL-based delegated Microsoft Graph authentication."""

from __future__ import annotations

import asyncio
import logging
import sys
from collections.abc import Mapping
from datetime import UTC, datetime, timedelta
from typing import Any

import msal

GRAPH_DELEGATED_SCOPES = ("Files.Read.All", "Sites.Read.All")

logger = logging.getLogger(__name__)


class MsalAuthenticationError(RuntimeError):
    """Raised when MSAL cannot acquire a delegated Microsoft Graph token."""


_PUBLIC_CLIENT_CONFIGURATION_ERROR = "AADSTS7000218"


def _show_device_code(flow: Mapping[str, Any]) -> None:
    verification_uri = str(
        flow.get("verification_uri") or "https://microsoft.com/devicelogin"
    )
    user_code = str(flow.get("user_code") or "<missing>")
    expiration = "before the code expires"
    expires_in = flow.get("expires_in")
    if isinstance(expires_in, int) and not isinstance(expires_in, bool):
        expires_on = datetime.now(UTC) + timedelta(seconds=expires_in)
        expiration = f"by {expires_on.astimezone():%Y-%m-%d %H:%M:%S %Z}"

    print(
        "\nMicrosoft Graph sign-in is required:\n"
        f"  1. Open {verification_uri}\n"
        f"  2. Enter code {user_code}\n"
        f"Complete sign-in {expiration}.\n"
        "Waiting for browser sign-in...",
        file=sys.stderr,
        flush=True,
    )


def _access_token(result: Mapping[str, Any] | None) -> str | None:
    if result is None:
        return None
    value = result.get("access_token")
    return value if isinstance(value, str) and value else None


def _authentication_error(action: str, result: Mapping[str, Any]) -> MsalAuthenticationError:
    error = result.get("error")
    description = result.get("error_description")
    correlation_id = result.get("correlation_id")

    if (
        error == "invalid_client"
        and isinstance(description, str)
        and _PUBLIC_CLIENT_CONFIGURATION_ERROR in description
    ):
        message = (
            f"{action}: {_PUBLIC_CLIENT_CONFIGURATION_ERROR}. The GRAPH_CLIENT_ID app registration "
            "is not enabled as a public client. In Microsoft Entra admin center, open App "
            "registrations > your application > Authentication > Advanced settings, set Allow "
            "public client flows to Yes, and save. Confirm GRAPH_CLIENT_ID is the Application "
            "(client) ID and GRAPH_TENANT_ID is the Directory (tenant) ID. Do not add a client "
            "secret to this desktop/device-code sample."
        )
        if isinstance(correlation_id, str) and correlation_id:
            message += f" Correlation ID: {correlation_id}."
        return MsalAuthenticationError(message)

    details = [action]
    if isinstance(error, str) and error:
        details.append(error)
    if isinstance(description, str) and description:
        details.append(description)
    if isinstance(correlation_id, str) and correlation_id:
        details.append(f"correlation ID: {correlation_id}")
    return MsalAuthenticationError(": ".join(details))


class MsalDeviceCodeTokenProvider:
    """Acquire delegated Graph tokens with MSAL and its in-memory token cache."""

    def __init__(self, *, tenant_id: str, client_id: str) -> None:
        self._application = msal.PublicClientApplication(
            client_id=client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
        )
        self._lock = asyncio.Lock()

    def _get_token_sync(self) -> str:
        scopes = list(GRAPH_DELEGATED_SCOPES)
        logger.info("[1/4] Authentication: checking the MSAL token cache.")
        accounts = self._application.get_accounts()
        for account in accounts:
            token = _access_token(
                self._application.acquire_token_silent(scopes, account=account)
            )
            if token is not None:
                logger.info("[1/4] Authentication: using a cached delegated Graph token.")
                return token

        logger.info(
            "[1/4] Authentication: no usable cached token; starting MSAL device-code sign-in."
        )
        flow = self._application.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise _authentication_error("Could not start Microsoft Graph sign-in", flow)

        _show_device_code(flow)
        result = self._application.acquire_token_by_device_flow(flow)
        token = _access_token(result)
        if token is None:
            raise _authentication_error("Microsoft Graph sign-in failed", result)
        logger.info("[1/4] Authentication: delegated Graph sign-in completed.")
        return token

    async def get_token(self) -> str:
        async with self._lock:
            return await asyncio.to_thread(self._get_token_sync)
