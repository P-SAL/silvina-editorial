from unittest import TestCase

from src.domain.enums.analysis_dimension import AnalysisDimension
from src.domain.enums.quality_dimension import QualityDimension


class TestQualityDimension(TestCase):
    def test_enum_has_exactly_four_members(self):
        self.assertEqual(len(QualityDimension), 4)

    def test_enum_members_are_claridad_coherencia_argumentacion_conclusiones(self):
        self.assertEqual(QualityDimension.CLARIDAD.value, "claridad")
        self.assertEqual(QualityDimension.COHERENCIA.value, "coherencia")
        self.assertEqual(QualityDimension.ARGUMENTACION.value, "argumentacion")
        self.assertEqual(QualityDimension.CONCLUSIONES.value, "conclusiones")

    def test_quality_dimension_does_not_subclass_analysis_dimension(self):
        self.assertFalse(issubclass(QualityDimension, AnalysisDimension))
        self.assertFalse(issubclass(AnalysisDimension, QualityDimension))
