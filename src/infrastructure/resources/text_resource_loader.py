from os.path import join
from pathlib import Path


def read_text_resource(directory: str, filename: str) -> str:
    """Read a UTF-8 text resource file from the given directory."""
    return Path(join(directory, filename)).read_text(encoding="utf-8")
