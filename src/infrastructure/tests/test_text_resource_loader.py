from os.path import join
from pathlib import Path
from unittest import TestCase

from src.infrastructure.resources.prompts.quality import PROMPTS_DIR
from src.infrastructure.resources.text_resource_loader import read_text_resource


class TestTextResourceLoader(TestCase):
    def test_reads_utf8_text_resource_file(self):
        expected_content = Path(join(PROMPTS_DIR, "clarity_coherence_prompt.txt")).read_text(
            encoding="utf-8"
        )

        result = read_text_resource(directory=PROMPTS_DIR, filename="clarity_coherence_prompt.txt")

        self.assertEqual(result, expected_content)
