"""SharePoint-only agent powered by the Microsoft 365 Copilot Retrieval API."""

from .agent import SharePointAnswerAgent
from .models import GroundedAnswer, RetrievalExtract, RetrievalHit
from .scope import SharePointScope

__all__ = [
    "GroundedAnswer",
    "RetrievalExtract",
    "RetrievalHit",
    "SharePointAnswerAgent",
    "SharePointScope",
]

__version__ = "0.1.0"
