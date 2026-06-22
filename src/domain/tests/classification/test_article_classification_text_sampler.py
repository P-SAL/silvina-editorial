from unittest import TestCase

from src.domain.classification.article_classification_text_sampler import (
    ArticleClassificationTextSampler,
)
from src.domain.dtos.document_content_dto import DocumentContentDTO


class TestArticleClassificationTextSampler(TestCase):
    def _build_document_content(self, paragraphs: list[str]) -> DocumentContentDTO:
        full_text = " ".join(paragraphs)
        return DocumentContentDTO(
            word_count=len(full_text.split()),
            char_count=len(full_text),
            paragraphs=paragraphs,
        )

    def test_bibliography_section_is_excluded_from_the_sample(self):
        paragraphs = (
            ["Intro uno " + "palabra " * 600]
            + ["Referencias"]
            + ["Smith 2020 entrada bibliografica excluida " + "palabra " * 100]
        )
        document_content = self._build_document_content(paragraphs)
        sampler = ArticleClassificationTextSampler()

        sample = sampler.build_sample(document_content)

        self.assertNotIn("entrada bibliografica excluida", sample)

    def test_sample_combines_intro_and_ending_segments(self):
        paragraphs = ["Intro " + "palabra " * 1500] + ["Final " + "palabra " * 700]
        document_content = self._build_document_content(paragraphs)
        full_text = " ".join(paragraphs)
        sampler = ArticleClassificationTextSampler()

        sample = sampler.build_sample(document_content)

        self.assertEqual(sample, (full_text[:3500] + " " + full_text[-2500:]).strip())

    def test_empty_sample_falls_back_to_first_six_thousand_characters_of_full_text(self):
        paragraphs = ["", "Referencias", "Entrada bibliografica " + "palabra " * 600]
        document_content = self._build_document_content(paragraphs)
        full_text = " ".join(paragraphs)
        sampler = ArticleClassificationTextSampler()

        sample = sampler.build_sample(document_content)

        self.assertEqual(sample, full_text[:6000])
