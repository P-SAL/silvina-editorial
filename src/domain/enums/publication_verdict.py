from enum import Enum


class PublicationVerdict(Enum):
    """Final publication verdict for an analyzed document."""

    CRITICAL = "critica"
    WARNING = "advertencia"
    APPROVED = "aprobado"
