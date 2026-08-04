from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError, SrcBaseWarning


class TestCountError(TestCase):
    def test_is_subclass_of_base_src_error(self):
        from src.domain.exceptions.count_errors import CountError

        self.assertTrue(issubclass(CountError, BaseSrcError))


class TestCharacterCountUnavailable(TestCase):
    def test_is_subclass_of_src_base_warning(self):
        from src.domain.exceptions.count_errors import CharacterCountUnavailable

        self.assertTrue(issubclass(CharacterCountUnavailable, SrcBaseWarning))

    def test_is_catchable_as_character_count_unavailable(self):
        from src.domain.exceptions.count_errors import CharacterCountUnavailable

        with self.assertRaises(CharacterCountUnavailable):
            raise CharacterCountUnavailable()

    def test_does_not_propagate_when_caught(self):
        from src.domain.exceptions.count_errors import CharacterCountUnavailable

        caught = False
        try:
            raise CharacterCountUnavailable()
        except CharacterCountUnavailable:
            caught = True

        self.assertTrue(caught)
