from unittest import TestCase

from src.domain.entities.base_entity import BaseEntity
from src.domain.enums.citation_type import CitationType


class TestCitation(TestCase):
    def _import_citation(self):
        from src.domain.citation.citation import Citation

        return Citation

    def test_citation_is_subclass_of_base_entity(self):
        Citation = self._import_citation()
        self.assertTrue(issubclass(Citation, BaseEntity))

    def test_citation_instantiation_with_required_fields_only(self):
        Citation = self._import_citation()
        citation = Citation(text="Some text", citation_type=CitationType.AUTHOR_YEAR, location=0)
        self.assertIsNone(citation.author)
        self.assertIsNone(citation.year)

    def test_citation_as_dict_contains_expected_keys(self):
        Citation = self._import_citation()
        citation = Citation(text="Some text", citation_type=CitationType.NUMERIC, location=1)
        result = citation.as_dict()
        self.assertIn("text", result)
        self.assertIn("citation_type", result)
        self.assertIn("location", result)
        self.assertIn("author", result)
        self.assertIn("year", result)

    def test_citation_str_truncates_at_50_chars(self):
        Citation = self._import_citation()
        long_text = "A" * 60
        citation = Citation(text=long_text, citation_type=CitationType.FOOTNOTE, location=2)
        result = str(citation)
        self.assertTrue(result.startswith("Citation("))
        self.assertIn("...", result)

    def test_citation_type_hints_use_modern_syntax(self):
        Citation = self._import_citation()
        import inspect

        source = inspect.getsource(Citation)
        self.assertNotIn("Optional[", source)
        self.assertNotIn("List[", source)
        self.assertNotIn("Dict[", source)
