from enum import Enum


class ReferenceLineMarker(Enum):
    """Substrings that mark a paragraph's opening characters as reference-like."""

    HTTP = "http"
    DOI = "doi.org"
    HTTPS = "https"
    ISBN = "ISBN"
