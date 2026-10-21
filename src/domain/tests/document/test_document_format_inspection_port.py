from inspect import getsource
from sys import modules
from unittest import TestCase

from src.domain.document.document_format_inspection_port import DocumentFormatInspectionPort


class TestDocumentFormatInspectionPort(TestCase):
    def test_is_abstract_base_class(self):
        with self.assertRaises(TypeError):
            DocumentFormatInspectionPort()

    def test_declares_exactly_one_abstract_method_inspect(self):
        self.assertEqual(
            DocumentFormatInspectionPort.__abstractmethods__,
            frozenset({"inspect"}),
        )

    def test_module_has_no_docx_or_infrastructure_imports(self):
        module_source = getsource(modules[DocumentFormatInspectionPort.__module__])
        self.assertNotIn("import docx", module_source)
        self.assertNotIn("src.infrastructure", module_source)
