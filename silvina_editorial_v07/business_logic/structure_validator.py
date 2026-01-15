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
    
    def validate_structure(self, document_content, category):
        """
        Validate if document has required sections for its category.
        
        Args:
            document_content: DocumentContent object
            category: ClassificationCategory enum
            
        Returns:
            StructureValidationResult with validation details
        """
        required_sections = self._get_required_sections(category)
        present_sections = self._extract_present_sections(document_content)
        
        missing_sections = [
            section for section in required_sections 
            if section.lower() not in [s.lower() for s in present_sections]
        ]
        
        section_details = {}
        for section in required_sections:
            is_present = section.lower() in [s.lower() for s in present_sections]
            section_details[section] = {
                'present': is_present,
                'required': True
            }
        
        from domain.models import StructureValidationResult
        return StructureValidationResult(
            is_valid=len(missing_sections) == 0,
            missing_sections=missing_sections,
            section_details=section_details
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