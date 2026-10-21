from inspect import signature
from unittest import TestCase

from src.domain.document.reference_extraction_port import ReferenceExtractionPort
from src.domain.dtos.reference_dto import ReferenceDTO
from src.domain.tests.document.fake_reference_extraction_port import FakeReferenceExtractionPort


class TestReferenceExtractionPort(TestCase):
    def test_s2a_cannot_instantiate_abstract_class(self):
        with self.assertRaises(TypeError):
            ReferenceExtractionPort()

    def test_s2b_extract_references_has_docx_path_str_parameter(self):
        sig = signature(ReferenceExtractionPort.extract_references)
        self.assertIn("docx_path", sig.parameters)
        self.assertEqual(sig.parameters["docx_path"].annotation, str)

    def test_s2b_extract_references_returns_tuple_of_list_and_str(self):
        sig = signature(ReferenceExtractionPort.extract_references)
        self.assertEqual(sig.return_annotation, tuple[list[ReferenceDTO], str])

    def test_s6a_fake_returns_configured_result(self):
        reference = ReferenceDTO(text="Smith, J. (2020). Title. Journal.")
        result_tuple = ([reference], "Referencias")
        fake = FakeReferenceExtractionPort(result=result_tuple)
        result = fake.extract_references(docx_path="test.docx")
        self.assertEqual(result, result_tuple)

    def test_s6b_fake_raises_configured_exception(self):
        error = ValueError("extraction failed")
        fake = FakeReferenceExtractionPort(error=error)
        with self.assertRaises(ValueError):
            fake.extract_references(docx_path="test.docx")
