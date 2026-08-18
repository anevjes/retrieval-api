from __future__ import annotations

import logging

from sharepoint_retrieval_agent.cli import _configure_logging


def test_normal_logging_suppresses_expected_device_code_polling(monkeypatch) -> None:
    app_logger = logging.getLogger("sharepoint_retrieval_agent")
    http_logger = logging.getLogger("azure.core.pipeline.policies.http_logging_policy")
    msal_logger = logging.getLogger("msal")
    original_app_level = app_logger.level
    original_http_level = http_logger.level
    original_msal_level = msal_logger.level
    try:
        monkeypatch.setenv("LOG_LEVEL", "INFO")
        monkeypatch.delenv("AUTH_SDK_LOG_LEVEL", raising=False)
        monkeypatch.delenv("AZURE_SDK_LOG_LEVEL", raising=False)

        _configure_logging(debug=False)

        assert app_logger.isEnabledFor(logging.INFO)
        assert not http_logger.isEnabledFor(logging.INFO)
        assert http_logger.isEnabledFor(logging.WARNING)
        assert not msal_logger.isEnabledFor(logging.INFO)
        assert msal_logger.isEnabledFor(logging.WARNING)
    finally:
        app_logger.setLevel(original_app_level)
        http_logger.setLevel(original_http_level)
        msal_logger.setLevel(original_msal_level)


def test_auth_sdk_logging_can_be_enabled_explicitly(monkeypatch) -> None:
    msal_logger = logging.getLogger("msal")
    original_level = msal_logger.level
    try:
        monkeypatch.setenv("AUTH_SDK_LOG_LEVEL", "INFO")

        _configure_logging(debug=False)

        assert msal_logger.isEnabledFor(logging.INFO)
    finally:
        msal_logger.setLevel(original_level)