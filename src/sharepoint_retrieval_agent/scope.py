"""Safe construction and enforcement of SharePoint site scopes."""

from __future__ import annotations

import posixpath
import re
from dataclasses import dataclass
from urllib.parse import quote, unquote, urlsplit, urlunsplit

_SHAREPOINT_HOST_PATTERN = re.compile(r"^[a-z0-9][a-z0-9-]*\.sharepoint\.com$")
_URL_PATH_SAFE_CHARACTERS = "/:@!$&'()*+,;=-._~"


class ScopeValidationError(ValueError):
    """Raised when a SharePoint site URL cannot safely become a KQL filter."""


def _normalize_path(path: str, *, trailing_slash: bool) -> str:
    decoded = unquote(path or "/")
    if "\x00" in decoded or "\\" in decoded:
        raise ScopeValidationError("SharePoint paths cannot contain nulls or backslashes.")
    if any(segment in {".", ".."} for segment in decoded.split("/")):
        raise ScopeValidationError("SharePoint paths cannot contain dot segments.")
    normalized = posixpath.normpath("/" + decoded.lstrip("/"))
    encoded = quote(normalized, safe=_URL_PATH_SAFE_CHARACTERS)
    if trailing_slash and encoded != "/" and not encoded.endswith("/"):
        encoded += "/"
    return encoded


def _validated_sharepoint_host(hostname: str | None) -> str:
    host = (hostname or "").rstrip(".").casefold()
    if not _SHAREPOINT_HOST_PATTERN.fullmatch(host):
        raise ScopeValidationError(
            "Site URLs must use a tenant SharePoint Online host such as "
            "https://contoso.sharepoint.com/."
        )
    if host.endswith("-my.sharepoint.com"):
        raise ScopeValidationError("OneDrive personal-site hosts are intentionally not allowed.")
    return host


def canonicalize_site_url(site_url: str) -> str:
    """Validate and canonicalize a SharePoint site path for KQL filtering."""

    candidate = site_url.strip()
    if not candidate:
        raise ScopeValidationError("SharePoint site URLs cannot be empty.")

    parsed = urlsplit(candidate)
    if parsed.scheme.casefold() != "https":
        raise ScopeValidationError("SharePoint site URLs must use HTTPS.")
    if parsed.username or parsed.password:
        raise ScopeValidationError("SharePoint site URLs cannot contain user information.")
    try:
        if parsed.port is not None:
            raise ScopeValidationError("SharePoint site URLs cannot specify a port.")
    except ValueError as error:
        raise ScopeValidationError("The SharePoint site URL has an invalid port.") from error
    if parsed.query or parsed.fragment:
        raise ScopeValidationError(
            "Use the canonical SharePoint path from the Details pane, without query strings "
            "or fragments."
        )

    host = _validated_sharepoint_host(parsed.hostname)
    path = _normalize_path(parsed.path, trailing_slash=True)
    return urlunsplit(("https", host, path, "", ""))


def is_sharepoint_content_url(url: str) -> bool:
    """Return whether a result URL is SharePoint Online rather than OneDrive."""

    try:
        parsed = urlsplit(url)
        if parsed.scheme.casefold() != "https" or parsed.username or parsed.password:
            return False
        if parsed.port is not None:
            return False
        _validated_sharepoint_host(parsed.hostname)
    except (ScopeValidationError, ValueError):
        return False
    return True


def _comparison_parts(url: str) -> tuple[str, str]:
    parsed = urlsplit(url)
    host = (parsed.hostname or "").casefold()
    path = posixpath.normpath("/" + unquote(parsed.path or "/").lstrip("/")).casefold()
    return host, path


@dataclass(frozen=True, slots=True)
class SharePointScope:
    """All accessible SharePoint sites or an explicit site-path allowlist."""

    site_urls: tuple[str, ...] = ()

    @classmethod
    def all_accessible_sites(cls) -> SharePointScope:
        """Create a scope covering all SharePoint content visible to the signed-in user."""

        return cls()

    @classmethod
    def selected_sites(cls, site_urls: list[str] | tuple[str, ...]) -> SharePointScope:
        """Create a canonical, deduplicated allowlist of SharePoint site paths."""

        if not site_urls:
            raise ScopeValidationError("Provide at least one SharePoint site URL.")
        canonical: dict[str, str] = {}
        for site_url in site_urls:
            normalized = canonicalize_site_url(site_url)
            canonical.setdefault(normalized.casefold(), normalized)
        return cls(tuple(canonical.values()))

    @property
    def includes_all_accessible_sites(self) -> bool:
        return not self.site_urls

    @property
    def filter_expression(self) -> str | None:
        """Build KQL solely from validated application input."""

        if self.includes_all_accessible_sites:
            return None
        return " OR ".join(f'Path:"{url}"' for url in self.site_urls)

    def allows(self, result_url: str) -> bool:
        """Post-filter a result to fail closed if the Retrieval API filter is bypassed."""

        if not is_sharepoint_content_url(result_url):
            return False
        if self.includes_all_accessible_sites:
            return True

        result_host, result_path = _comparison_parts(result_url)
        for site_url in self.site_urls:
            site_host, site_path = _comparison_parts(site_url)
            if result_host != site_host:
                continue
            if (
                site_path == "/"
                or result_path == site_path
                or result_path.startswith(site_path + "/")
            ):
                return True
        return False
