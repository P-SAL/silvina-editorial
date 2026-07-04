from src.domain.dtos.publication_verdict_dto import PublicationVerdictDTO
from src.domain.enums.publication_verdict import PublicationVerdict
from src.domain.recommendation.analysis_context import AnalysisContext


class PublicationVerdictEvaluator:
    """Evaluates the final publication verdict from the full analysis context."""

    def evaluate(self, context: AnalysisContext) -> PublicationVerdictDTO:
        match_rate = context.citation_match_rate

        has_critical_issues = (
            context.quality.overall_score < context.settings.critical_quality_threshold
            or context.grammar.score < context.settings.critical_grammar_threshold
            or not context.structure.is_valid
            or match_rate < context.settings.critical_citation_match_threshold
        )
        has_warnings = (
            context.quality.overall_score < context.settings.publish_threshold
            or context.grammar.score < context.settings.publish_threshold
            or match_rate < context.settings.citation_match_threshold
            or len(context.apa_validation.violations) > 0
        )

        if has_critical_issues:
            return PublicationVerdictDTO(
                verdict=PublicationVerdict.CRITICAL,
                message="❌ NO APTO PARA PUBLICACIÓN. El documento presenta errores críticos que deben corregirse.",
            )
        if context.citations.total_citations == 0:
            return PublicationVerdictDTO(
                verdict=PublicationVerdict.CRITICAL,
                message="❌ NO APTO PARA PUBLICACIÓN. No se detectaron citas APA en el texto. Verifique el formato de citación según normas APA 7.",
            )
        if has_warnings:
            return PublicationVerdictDTO(
                verdict=PublicationVerdict.WARNING,
                message="⚠️ REQUIERE REVISIÓN antes de publicación. Corrija los problemas identificados.",
            )
        return PublicationVerdictDTO(
            verdict=PublicationVerdict.APPROVED,
            message="✅ APTO PARA PUBLICACIÓN. El documento cumple con los estándares de calidad.",
        )
