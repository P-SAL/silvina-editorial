from abc import ABC, abstractmethod


class FileGatewayPort(ABC):
    """File gateway Port."""

    @abstractmethod
    def read(self, file_path: str) -> str:
        """Read the contents of the file."""

    @abstractmethod
    def write(self, file_path: str, content: str) -> None:
        """Write the content to a file."""

    @abstractmethod
    def remove(self, file_path: str) -> None:
        """Remove file."""
