import os

from src.domain.gateway.file_gateway_port import FileGatewayPort


class FileGatewayAdapter(FileGatewayPort):
    """File gateway adapter."""

    def read(self, file_path: str) -> str:
        """Read the contents of the file."""
        with open(file_path, "r") as file:
            return file.read()

    def write(self, file_path: str, content: str) -> None:
        """Write the content to a file."""
        with open(file_path, "w") as file:
            file.write(content)

    def remove(self, file_path: str) -> None:
        """Remove file."""
        if os.path.exists(file_path):
            os.remove(file_path)
