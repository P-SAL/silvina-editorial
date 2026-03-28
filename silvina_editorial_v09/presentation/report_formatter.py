"""
report_formatter.py
Generates formatted text reports from analysis results.
Part of Silvina Editorial Assistant v0.7
"""

from datetime import datetime
from typing import Dict, List, Any
from domain.enums import CitationType, ClassificationCategory, QualityLevel


class ReportFormatter:
    """Formats analysis results into readable text reports."""
    
    def __init__(self):
        """Initialize the report formatter."""
        self.separator_thick = "=" * 80
        self.separator_thin = "-" * 80
        self.separator_section = "·" * 80
    
    def _determine_publishability(self, analysis_results):
        """
        Decide if the document is publishable based on analysis results.
        """

        quality = analysis_results.get("quality")
        structure = analysis_results.get("structure")

        # Defensive defaults
        if not quality or not structure:
            return False, "Información insuficiente para determinar publicabilidad"

        # Quality rules
        if quality.overall_score < 6.0:
            return False, "Calidad académica insuficiente"

        # Structure rules
        if not structure.is_complete:
            return False, "Estructura académica incompleta"

        return True, "Cumple criterios mínimos de publicación"


    def generate_full_report(self, analysis_results: Dict[str, Any]) -> str:
        """Generate complete formatted report."""
        report_lines = []
        
        # Header
        report_lines.append("=" * 80)
        report_lines.append("INFORME DE ANÁLISIS EDITORIAL - SILVINA v0.7")
        report_lines.append("=" * 80)
        report_lines.append("")
        
        # **ADD THIS NEW SECTION**
        report_lines.append("━" * 80)
        report_lines.append("DECISIÓN DE PUBLICACIÓN")
        report_lines.append("━" * 80)
        report_lines.append("")
        
        # Determine publishability
        can_publish, decision_reason = self._determine_publishability(analysis_results)
        
        if can_publish:
            report_lines.append("✅ RECOMENDACIÓN: APTO PARA PUBLICACIÓN")
        else:
            report_lines.append("❌ RECOMENDACIÓN: REQUIERE REVISIÓN ANTES DE PUBLICACIÓN")
        
        report_lines.append("")
        report_lines.append(f"Justificación: {decision_reason}")
        report_lines.append("")
        report_lines.append("━" * 80)
        report_lines.append("")
        
         
        report_sections = [
            self._generate_header(analysis_results),
            self._generate_document_info(analysis_results),
            self._generate_classification_section(analysis_results),
            self._generate_quality_section(analysis_results),
            self._generate_structure_section(analysis_results),
            self._generate_citations_section(analysis_results),
            self._generate_recommendations_section(analysis_results),
            self._generate_footer()
        ]
        
        return "\n\n".join(report_sections)
    
    def _generate_header(self, results: Dict[str, Any]) -> str:
        """Generate report header."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        header = f"""
{self.separator_thick}
                    SILVINA EDITORIAL ASSISTANT v0.7
                      Informe de Análisis Editorial
{self.separator_thick}

Fecha de análisis: {timestamp}
Documento analizado: {results.get('filename', 'N/A')}
        """
        return header.strip()
    
    def _generate_document_info(self, results: Dict[str, Any]) -> str:
        """Generate document information section."""
        doc_info = results.get('document_info', {})
        
        section = f"""
{self.separator_thin}
1. INFORMACIÓN DEL DOCUMENTO
{self.separator_thin}

Título: {doc_info.get('title', 'No especificado')}
Autor(es): {doc_info.get('authors', 'No especificado')}
Total de palabras: {doc_info.get('word_count', 0):,}
Total de caracteres (con espacios): {doc_info.get('char_count', 0):,}
Total de párrafos: {doc_info.get('paragraph_count', 0)}
Páginas estimadas: {doc_info.get('estimated_pages', 0)}

     """
        return section.strip()
    
    def _generate_classification_section(self, results: Dict[str, Any]) -> str:
        """Generate article classification section."""
        classification = results.get('classification', {})
        
        category = classification.get('category', ClassificationCategory.UNKNOWN)
        confidence = classification.get('confidence', 0.0)
        reasoning = classification.get('reasoning', 'No disponible')
        
        # Format category name
        category_name = self._format_category_name(category)
        confidence_bar = self._create_confidence_bar(confidence)
        
        section = f"""
{self.separator_thin}
2. CLASIFICACIÓN DEL ARTÍCULO
{self.separator_thin}

Categoría identificada: {category_name}
Nivel de confianza: {confidence:.1%} {confidence_bar}

Razonamiento:
{self._wrap_text(reasoning, indent=2)}
        """
        return section.strip()
    
    def _generate_quality_section(self, results: Dict[str, Any]) -> str:
        """Generate quality analysis section."""
        quality = results.get('quality_analysis', {})
        
        overall_score = quality.get('overall_score', 0.0)
        quality_level = quality.get('quality_level', QualityLevel.NEEDS_IMPROVEMENT)
        
        section = f"""
{self.separator_thin}
3. ANÁLISIS DE CALIDAD
{self.separator_thin}

Puntuación general: {overall_score:.1f}/10.0 {self._get_score_indicator(overall_score)}
Nivel de calidad: {self._format_quality_level(quality_level)}

{self.separator_section}
DIMENSIONES EVALUADAS:
{self.separator_section}
"""
        
        # Add dimension scores
        dimensions = quality.get('dimensions', {})
        for dimension_name, dimension_data in dimensions.items():
            score = dimension_data.get('score', 0.0)
            feedback = dimension_data.get('feedback', 'No disponible')
            
            section += f"""
{dimension_name.upper().replace('_', ' ')}:
  Puntuación: {score:.1f}/10.0 {self._get_score_indicator(score)}
  {self._wrap_text(feedback, indent=2)}
"""
        
        return section.strip()
    
    def _generate_structure_section(self, results: Dict[str, Any]) -> str:
        """Generate structure validation section."""
        structure = results.get('structure_validation', {})
        
        is_valid = structure.get('is_valid', False)
        missing_sections = structure.get('missing_sections', [])
        validation_details = structure.get('details', {})
        
        status_icon = "✓" if is_valid else "✗"
        status_text = "VÁLIDA" if is_valid else "INCOMPLETA"
        
        section = f"""
{self.separator_thin}
4. VALIDACIÓN DE ESTRUCTURA
{self.separator_thin}

Estado: {status_icon} {status_text}
"""
        
        if missing_sections:
            section += f"""
Secciones faltantes:
{self._format_list(missing_sections, bullet='  ⚠')}
"""
        
        section += f"""
{self.separator_section}
DETALLES DE SECCIONES:
{self.separator_section}
"""
        
        for section_name, section_data in validation_details.items():
            found = section_data.get('found', False)
            status = "✓ Presente" if found else "✗ Ausente"
            
            section += f"\n{section_name}: {status}"
            
            if found and 'content_preview' in section_data:
                preview = section_data['content_preview'][:100] + "..."
                section += f"\n  Vista previa: {preview}"
        
        return section.strip()
    
    def _generate_citations_section(self, results: Dict[str, Any]) -> str:
        """Generate citations analysis section."""
        citations = results.get('citations_analysis', {})
        
        total_citations = citations.get('total_citations', 0)
        citation_types = citations.get('by_type', {})
        matched_count = citations.get('matched_count', 0)
        unmatched_count = citations.get('unmatched_count', 0)
        
        match_rate = (matched_count / total_citations * 100) if total_citations > 0 else 0
        
        section = f"""
{self.separator_thin}
5. ANÁLISIS DE CITAS Y REFERENCIAS
{self.separator_thin}

Total de citas: {total_citations}
Citas coincidentes: {matched_count}
Citas sin coincidencia: {unmatched_count}
Tasa de coincidencia: {match_rate:.1f}% {self._create_progress_bar(match_rate)}

{self.separator_section}
DISTRIBUCIÓN POR TIPO:
{self.separator_section}
"""
        
        for citation_type, count in citation_types.items():
            type_name = self._format_citation_type(citation_type)
            percentage = (count / total_citations * 100) if total_citations > 0 else 0
            section += f"\n{type_name}: {count} ({percentage:.1f}%)"
        
        # Add unmatched citations if any
        unmatched_list = citations.get('unmatched_citations', [])
        if unmatched_list:
            section += f"""

{self.separator_section}
CITAS SIN REFERENCIA CORRESPONDIENTE:
{self.separator_section}
"""
            for i, citation in enumerate(unmatched_list[:10], 1):  # Show first 10
                section += f"\n{i}. {citation}"
            
            if len(unmatched_list) > 10:
                section += f"\n... y {len(unmatched_list) - 10} más"
        
        return section.strip()
    
    def _generate_recommendations_section(self, results: Dict[str, Any]) -> str:
        """Generate recommendations section."""
        recommendations = results.get('recommendations', [])
        
        if not recommendations:
            return f"""
{self.separator_thin}
6. RECOMENDACIONES
{self.separator_thin}

No hay recomendaciones específicas. El documento cumple con los estándares básicos.
            """.strip()
        
        section = f"""
{self.separator_thin}
6. RECOMENDACIONES
{self.separator_thin}

Se han identificado las siguientes áreas de mejora:
"""
        
        for i, rec in enumerate(recommendations, 1):
            priority = rec.get('priority', 'media')
            priority_icon = self._get_priority_icon(priority)
            message = rec.get('message', 'Sin descripción')
            
            section += f"\n{i}. {priority_icon} {message}"
        
        return section.strip()
    
    def _generate_footer(self) -> str:
        """Generate report footer."""
        footer = f"""
{self.separator_thick}
Este informe fue generado automáticamente por Silvina Editorial Assistant.
Para más información sobre las Normas EUMIC, consulte la documentación oficial.
{self.separator_thick}
        """
        return footer.strip()
    
    # Helper formatting methods
    
    def _format_category_name(self, category: ClassificationCategory) -> str:
        """Format category name for display."""
        category_names = {
            ClassificationCategory.RESEARCH_ARTICLE: "Artículo de Investigación",
            ClassificationCategory.REVIEW_ARTICLE: "Artículo de Revisión",
            ClassificationCategory.REFLECTION_ARTICLE: "Artículo de Reflexión",
            ClassificationCategory.SHORT_ARTICLE: "Artículo Corto",
            ClassificationCategory.CASE_REPORT: "Reporte de Caso",
            ClassificationCategory.UNKNOWN: "No Clasificado"
        }
        return category_names.get(category, str(category))
    
    def _format_quality_level(self, level: QualityLevel) -> str:
        """Format quality level for display."""
        level_names = {
            QualityLevel.EXCELLENT: "⭐ EXCELENTE",
            QualityLevel.GOOD: "✓ BUENO",
            QualityLevel.ACCEPTABLE: "○ ACEPTABLE",
            QualityLevel.NEEDS_IMPROVEMENT: "△ NECESITA MEJORAS",
            QualityLevel.POOR: "✗ DEFICIENTE"
        }
        return level_names.get(level, str(level))
    
    def _format_citation_type(self, citation_type: CitationType) -> str:
        """Format citation type for display."""
        type_names = {
            CitationType.AUTHOR_YEAR: "Autor-año",
            CitationType.NUMERIC: "Numérica",
            CitationType.FOOTNOTE: "Nota al pie",
            CitationType.UNKNOWN: "Desconocida"
        }
        return type_names.get(citation_type, str(citation_type))
    
    def _create_confidence_bar(self, confidence: float) -> str:
        """Create a visual confidence bar."""
        bar_length = 20
        filled = int(confidence * bar_length)
        empty = bar_length - filled
        return f"[{'█' * filled}{'░' * empty}]"
    
    def _create_progress_bar(self, percentage: float) -> str:
        """Create a visual progress bar."""
        bar_length = 20
        filled = int(percentage / 100 * bar_length)
        empty = bar_length - filled
        return f"[{'█' * filled}{'░' * empty}]"
    
    def _get_score_indicator(self, score: float) -> str:
        """Get visual indicator for score."""
        if score >= 9.0:
            return "⭐⭐⭐"
        elif score >= 7.0:
            return "⭐⭐"
        elif score >= 5.0:
            return "⭐"
        else:
            return "○"
    
    def _get_priority_icon(self, priority: str) -> str:
        """Get icon for priority level."""
        icons = {
            'alta': '🔴',
            'media': '🟡',
            'baja': '🟢'
        }
        return icons.get(priority.lower(), '○')
    
    def _wrap_text(self, text: str, width: int = 76, indent: int = 0) -> str:
        """Wrap text to specified width with optional indentation."""
        import textwrap
        wrapper = textwrap.TextWrapper(
            width=width,
            initial_indent=' ' * indent,
            subsequent_indent=' ' * indent
        )
        return wrapper.fill(text)
    
    def _format_list(self, items: List[str], bullet: str = '  •') -> str:
        """Format a list of items with bullets."""
        return '\n'.join(f"{bullet} {item}" for item in items)
    
    def generate_summary_report(self, analysis_results: Dict[str, Any]) -> str:
        """
        Generate a short summary report.
        
        Args:
            analysis_results: Dictionary containing all analysis results
            
        Returns:
            Formatted summary report as string
        """
        classification = analysis_results.get('classification', {})
        quality = analysis_results.get('quality_analysis', {})
        structure = analysis_results.get('structure_validation', {})
        
        category = self._format_category_name(
            classification.get('category', ClassificationCategory.UNKNOWN)
        )
        quality_score = quality.get('overall_score', 0.0)
        structure_valid = "✓" if structure.get('is_valid', False) else "✗"
        
        summary = f"""
{self.separator_thick}
RESUMEN EJECUTIVO - SILVINA v0.7
{self.separator_thick}

Documento: {analysis_results.get('filename', 'N/A')}
Categoría: {category}
Calidad: {quality_score:.1f}/10.0
Estructura: {structure_valid}

{self.separator_thick}
        """
        return summary.strip()


# Convenience function for quick report generation
def create_report(analysis_results: Dict[str, Any], summary_only: bool = False) -> str:
    """
    Create a formatted report from analysis results.
    
    Args:
        analysis_results: Dictionary containing all analysis results
        summary_only: If True, generate only summary report
        
    Returns:
        Formatted report as string
    """
    formatter = ReportFormatter()
    
    if summary_only:
        return formatter.generate_summary_report(analysis_results)
    else:
        return formatter.generate_full_report(analysis_results)
    
def _determine_publishability(self, results: Dict[str, Any]) -> tuple:
    """
    Determine if document can be published based on analysis.
    
    Returns:
        (can_publish: bool, reason: str)
    """
    quality_score = results['quality_analysis']['overall_score']
    is_structure_valid = results['structure_validation']['is_valid']
    citations_matched = results['citations_analysis']['matched_count']
    citations_total = results['citations_analysis']['total_citations']
    
    # Calculate citation match rate
    if citations_total > 0:
        citation_rate = citations_matched / citations_total
    else:
        citation_rate = 1.0  # No citations is acceptable for some article types
    
    # Decision logic
    if quality_score >= 7.0 and is_structure_valid and citation_rate >= 0.9:
        return (True, "El documento cumple con los estándares de calidad, estructura y referencias requeridos por EUMIC.")
    
    elif quality_score >= 6.0 and is_structure_valid:
        return (False, f"Calidad aceptable ({quality_score:.1f}/10) pero requiere mejoras menores. Tasa de coincidencia de citas: {citation_rate:.1%}")
    
    elif not is_structure_valid:
        missing = ", ".join(results['structure_validation']['missing_sections'])
        return (False, f"Estructura incompleta. Faltan secciones: {missing}")
    
    else:
        return (False, f"Calidad insuficiente ({quality_score:.1f}/10). Requiere revisión sustancial.")