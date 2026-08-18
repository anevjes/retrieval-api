from __future__ import annotations

import pytest

from sharepoint_retrieval_agent.scope import (
    ScopeValidationError,
    SharePointScope,
    canonicalize_site_url,
)


def test_all_sites_has_no_filter_but_still_excludes_onedrive() -> None:
    scope = SharePointScope.all_accessible_sites()

    assert scope.filter_expression is None
    assert scope.allows("https://contoso.sharepoint.com/sites/HR/policy.docx")
    assert not scope.allows("https://contoso-my.sharepoint.com/personal/alex/file.docx")
    assert not scope.allows("https://outlook.office.com/mail/item")


def test_selected_sites_build_validated_kql() -> None:
    scope = SharePointScope.selected_sites(
        [
            "https://Contoso.sharepoint.com/sites/Engineering",
            "https://contoso.sharepoint.com/sites/Policies/",
        ]
    )

    assert scope.filter_expression == (
        'Path:"https://contoso.sharepoint.com/sites/Engineering/" OR '
        'Path:"https://contoso.sharepoint.com/sites/Policies/"'
    )


def test_document_library_path_uses_canonical_spaces_in_kql() -> None:
    scope = SharePointScope.selected_sites(
        ["https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/"]
    )

    assert scope.filter_expression == (
        'Path:"https://contoso.sharepoint.com/sites/Sales/Shared Documents/"'
    )


def test_post_filter_uses_path_boundaries() -> None:
    scope = SharePointScope.selected_sites(
        ["https://contoso.sharepoint.com/sites/HR/"]
    )

    assert scope.allows("https://contoso.sharepoint.com/sites/HR/Shared%20Documents/a.docx")
    assert not scope.allows("https://contoso.sharepoint.com/sites/HR2/a.docx")
    assert not scope.allows("https://fabrikam.sharepoint.com/sites/HR/a.docx")


@pytest.mark.parametrize(
    "url",
    [
        "http://contoso.sharepoint.com/sites/HR",
        "https://contoso-my.sharepoint.com/personal/alex",
        "https://contoso.sharepoint.com/sites/HR?sharing=abc",
        "https://contoso.sharepoint.com/sites/HR/%2e%2e/Finance",
        "https://outlook.office.com/sites/HR",
        "https://user@contoso.sharepoint.com/sites/HR",
    ],
)
def test_invalid_site_paths_fail_closed(url: str) -> None:
    with pytest.raises(ScopeValidationError):
        canonicalize_site_url(url)
