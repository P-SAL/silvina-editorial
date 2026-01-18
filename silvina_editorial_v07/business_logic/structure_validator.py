"""
Structure Validator - Validates IMRyD structure in scientific articles
"""

from typing import List, Dict
from domain.models import Section
from config import REQUIRED_SECTIONS
from domain.models import StructureValidationResult

class StructureValidator:
    """Validates IMRyD structure in scientific articles."""
    
    def validate(self, sections: List[Section]) -> Dict:
        """Validate structure and return issues."""
        issues = {
            "missing_sections": [],
            "out_of_order": [],
            "too_short": []
        }
        
        # Check for missing sections
        found_names = {s.name for s in sections}
        for required in REQUIRED_SECTIONS.keys():
            if required not in found_names:
                issues["missing_sections"].append(required)
        
        # Check order
        for i in range(len(sections) - 1):
            if sections[i].expected_order > sections[i + 1].expected_order:
                issues["out_of_order"].append((sections[i].name, sections[i + 1].name))
        
        # Check minimum length
        for section in sections:
            min_words = REQUIRED_SECTIONS[section.name]["min_words"]
            if section.word_count < min_words:
                issues["too_short"].append({
                    "section": section.name,
                    "current": section.word_count,
                    "minimum": min_words
                })
        
        return issues
    
    def validate_structure(self, document_content, article_type):
        if article_type == ArticleType.CIENTIFICO:
            required = ["Resumen", "Introducción", "Metodología", "Resultados", "Discusión", "Conclusiones", "Referencias"]
            justification = "Científico requiere IMRyD con razonamiento crítico y citación académica"
        elif article_type == ArticleType.DIVULGACION:
            required = ["Resumen", "Introducción", "Desarrollo", "Conclusiones", "Referencias"]
            justification = "Divulgación enfatiza reflexión crítica sin IMRyD rígido"
        elif article_type == ArticleType.OPINION:
            required = ["Introducción", "Argumentación", "Conclusiones"]
            justification = "Opinión privilegia crítica reflexiva sin validación empírica"
        else:
            required = ["Introducción", "Conclusiones"]
            justification = "Estructura mínima requerida"
        
        present = self._extract_present_sections(document_content)
        missing = [s for s in required if s.lower() not in [p.lower() for p in present]]
        
        details = {s: {'present': s.lower() in [p.lower() for p in present], 'required': True} 
                for s in required}
        
        return StructureValidationResult(
            is_valid=len(missing) == 0,
            missing_sections=missing,
            section_details=details,
            justification=justification
        )

        def _get_required_sections(self, category):
            """Get required sections based on article category."""
            from domain.enums import ClassificationCategory
            
            if category == ClassificationCategory.RESEARCH_ARTICLE:
                return ["Resumen", "Introducción", "Metodología", "Resultados", "Discusión", "Conclusiones", "Referencias"]
            elif category == ClassificationCategory.REVIEW_ARTICLE:
                return ["Resumen", "Introducción", "Desarrollo", "Conclusiones", "Referencias"]
            else:
                return ["Introducción", "Desarrollo", "Conclusiones", "Referencias"]

    def _extract_present_sections(self, document_content):
        """Extract section headers from document."""
        sections = []
        keywords = ["resumen", "abstract", "introducción", "metodología", "método", 
                    "resultados", "discusión", "conclusiones", "referencias", "bibliografía", "desarrollo"]
        
        for para in document_content.paragraphs:
            text_lower = para.lower().strip()
            for keyword in keywords:
                if keyword in text_lower and len(text_lower) < 50:
                    sections.append(para.strip())
                    break
    
        return sections