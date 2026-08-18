from __future__ import annotations

from typing import Any

import pytest

from sharepoint_retrieval_agent.auth import (
    GRAPH_DELEGATED_SCOPES,
    MsalAuthenticationError,
    MsalDeviceCodeTokenProvider,
    _show_device_code,
)


class FakeMsalApplication:
    def __init__(
        self,
        *,
        accounts: list[dict[str, Any]] | None = None,
        silent_result: dict[str, Any] | None = None,
        flow: dict[str, Any] | None = None,
        device_result: dict[str, Any] | None = None,
    ) -> None:
        self.accounts = accounts or []
        self.silent_result = silent_result
        self.flow = flow or {}
        self.device_result = device_result or {}
        self.silent_scopes: list[str] | None = None
        self.device_scopes: list[str] | None = None
        self.device_flow_completed = False

    def get_accounts(self, username: str | None = None) -> list[dict[str, Any]]:
        del username
        return self.accounts

    def acquire_token_silent(
        self,
        scopes: list[str],
        account: dict[str, Any],
    ) -> dict[str, Any] | None:
        del account
        self.silent_scopes = scopes
        return self.silent_result

    def initiate_device_flow(self, scopes: list[str]) -> dict[str, Any]:
        self.device_scopes = scopes
        return self.flow

    def acquire_token_by_device_flow(self, flow: dict[str, Any]) -> dict[str, Any]:
        assert flow is self.flow
        self.device_flow_completed = True
        return self.device_result


def test_device_code_prompt_explains_the_wait(capsys) -> None:
    _show_device_code(
        {
            "verification_uri": "https://microsoft.com/devicelogin",
            "user_code": "ABCD-EFGH",
            "expires_in": 900,
        }
    )

    output = capsys.readouterr().err
    assert "Open https://microsoft.com/devicelogin" in output
    assert "Enter code ABCD-EFGH" in output
    assert "Waiting for browser sign-in" in output


@pytest.mark.asyncio
async def test_msal_reuses_a_silent_cached_token() -> None:
    application = FakeMsalApplication(
        accounts=[{"home_account_id": "account-1"}],
        silent_result={"access_token": "cached-token"},
    )
    provider = MsalDeviceCodeTokenProvider(
        tenant_id="tenant-id",
        client_id="client-id",
        application=application,
    )

    token = await provider.get_token()

    assert token == "cached-token"
    assert application.silent_scopes == list(GRAPH_DELEGATED_SCOPES)
    assert application.device_scopes is None
    assert not application.device_flow_completed


@pytest.mark.asyncio
async def test_msal_uses_device_flow_after_cache_miss(capsys) -> None:
    application = FakeMsalApplication(
        flow={
            "verification_uri": "https://microsoft.com/devicelogin",
            "user_code": "ABCD-EFGH",
            "expires_in": 900,
        },
        device_result={"access_token": "interactive-token"},
    )
    provider = MsalDeviceCodeTokenProvider(
        tenant_id="tenant-id",
        client_id="client-id",
        application=application,
    )

    token = await provider.get_token()

    assert token == "interactive-token"
    assert application.device_scopes == list(GRAPH_DELEGATED_SCOPES)
    assert application.device_flow_completed
    assert "Enter code ABCD-EFGH" in capsys.readouterr().err


@pytest.mark.asyncio
async def test_msal_errors_expose_diagnostics_without_dumping_result() -> None:
    application = FakeMsalApplication(
        flow={
            "error": "invalid_client",
            "error_description": "The public client is not configured correctly.",
            "correlation_id": "correlation-123",
            "access_token": "must-not-be-printed",
        }
    )
    provider = MsalDeviceCodeTokenProvider(
        tenant_id="tenant-id",
        client_id="client-id",
        application=application,
    )

    with pytest.raises(MsalAuthenticationError) as caught:
        await provider.get_token()

    message = str(caught.value)
    assert "invalid_client" in message
    assert "correlation-123" in message
    assert "must-not-be-printed" not in message


@pytest.mark.asyncio
async def test_msal_explains_public_client_configuration_error() -> None:
    application = FakeMsalApplication(
        flow={
            "verification_uri": "https://login.microsoft.com/device",
            "user_code": "ABCD-EFGH",
        },
        device_result={
            "error": "invalid_client",
            "error_description": (
                "AADSTS7000218: The request body must contain the following parameter: "
                "'client_assertion' or 'client_secret'."
            ),
            "correlation_id": "correlation-7000218",
        },
    )
    provider = MsalDeviceCodeTokenProvider(
        tenant_id="tenant-id",
        client_id="client-id",
        application=application,
    )

    with pytest.raises(MsalAuthenticationError) as caught:
        await provider.get_token()

    message = str(caught.value)
    assert "Allow public client flows to Yes" in message
    assert "Application (client) ID" in message
    assert "Do not add a client secret" in message
    assert "correlation-7000218" in message