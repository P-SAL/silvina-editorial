import tempfile
from pathlib import Path
from unittest import TestCase

from src.infrastructure.adapters.gateway.file_gateway_adapter import FileGatewayAdapter

ACCENTED_SPANISH_TEXT = "línea de investigación — análisis operativo…"


class TestFileGatewayAdapter(TestCase):
    def setUp(self):
        self.adapter = FileGatewayAdapter()
        self._tmp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp_dir.cleanup)
        self.file_path = str(Path(self._tmp_dir.name) / "sample.txt")

    def test_write_then_read_round_trips_accented_spanish_text(self):
        self.adapter.write(self.file_path, ACCENTED_SPANISH_TEXT)

        result = self.adapter.read(self.file_path)

        self.assertEqual(result, ACCENTED_SPANISH_TEXT)

    def test_write_produces_utf8_encoded_bytes_on_disk(self):
        self.adapter.write(self.file_path, ACCENTED_SPANISH_TEXT)

        raw_bytes = Path(self.file_path).read_bytes()

        self.assertEqual(raw_bytes.decode("utf-8"), ACCENTED_SPANISH_TEXT)

    def test_read_decodes_an_externally_written_utf8_file_correctly(self):
        Path(self.file_path).write_bytes(ACCENTED_SPANISH_TEXT.encode("utf-8"))

        result = self.adapter.read(self.file_path)

        self.assertEqual(result, ACCENTED_SPANISH_TEXT)
