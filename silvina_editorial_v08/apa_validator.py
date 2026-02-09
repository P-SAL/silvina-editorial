"""
apa_validator.py
APA 7 Spanish Citation Validator
Part of Silvina Editorial Assistant v0.7

Validates APA 7 format compliance for Spanish-language citations.
Distinguishes between structural matching and formal correctness.
"""

import re
from typing import List, Tuple, Optional
from dataclasses import dataclass
from enum import Enum


class APAErrorType(Enum):
    """Types of APA format violations"""
    CONJUNCTION_ERROR = "Conjunción incorrecta"
    COMMA_ERROR = "Puntuación incorrecta"
    CAPITALIZATION_ERROR = "Mayúsculas/minúsculas incorrectas"
    ET_AL_FORMAT_ERROR = "Formato 'et al.' incorrecto"
    PAGE_FORMAT_ERROR = "Formato de página incorrecto"
    SPACING_ERROR = "Espaciado incorrecto"
    YEAR_FORMAT_ERROR = "Formato de año incorrecto"
    PARENTHESES_ERROR = "Paréntesis incorrectos"

@dataclass
class APAViolation:
    """Represents an APA format violation"""
    citation_text: str
    error_type: APAErrorType
    location: int  # Paragraph index
    explanation: str
    correction: str
    paragraph_preview: str = ""  # First 30 chars of paragraph


class APAValidator:
    """
    Validates APA 7 Spanish citation format compliance.
    
    Based on APA 7th edition Spanish language guidelines:
    - Use "y" (not "&") for conjunctions
    - Format: (Apellido, Año) or Apellido (Año)
    - Et al. for 3+ authors (no period in "al")
    - Page references: (p. 23) or (pp. 45-67)
    """
    
    def __init__(self):
        self.violations: List[APAViolation] = []
        
        # APA 7 Spanish rules
        self.rules = {
            'conjunction': 'y',  # Spanish uses "y", not "&"
            'et_al': 'et al.',   # No period in "al"
            'page_single': 'p.',
            'page_multiple': 'pp.',
        }
    
    def validate_citation(self, citation_text: str, paragraph_index: int, paragraph_text: str = "") -> List[APAViolation]:
        """
        Validate a single citation for APA 7 compliance.
        
        Args:
            citation_text: The citation text (e.g., "(García, 2020)")
            paragraph_index: Location in document
            
        Returns:
            List of violations found (empty if compliant)
        """
        violations = []
                
        # Determine citation type
        is_parenthetical = citation_text.startswith('(') and citation_text.endswith(')')
        
        # Create preview (first 30 chars)
        preview = paragraph_text[:30] + "..." if len(paragraph_text) > 30 else paragraph_text

        if is_parenthetical:
            violations.extend(self._validate_parenthetical(citation_text, paragraph_index, preview))
        else:
            violations.extend(self._validate_narrative(citation_text, paragraph_index, preview))

        return violations
    
    def _validate_parenthetical(self, citation: str, location: int, preview: str = "") -> List[APAViolation]:
        """Validate parenthetical citation: (Author, Year)"""
        violations = []
        
        # Remove outer parentheses for analysis
        inner = citation[1:-1].strip()
        
        # Check 1: Ampersand instead of "y"
        if ' & ' in inner:
            violations.append(APAViolation(
                citation_text=citation,
                error_type=APAErrorType.CONJUNCTION_ERROR,
                location=location,
                explanation='APA 7 español requiere "y" en lugar de "&" para citas parentéticas',
                correction=citation.replace(' & ', ' y '),
                paragraph_preview=preview
            ))
                                
        # Check 2: Missing comma between author and year
        # Pattern: (Author Year) instead of (Author, Year)
        pattern_no_comma = r'\(([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+)\s+(\d{4}[a-z]?)\)'
        if re.match(pattern_no_comma, citation) and ',' not in citation:
            author = re.match(pattern_no_comma, citation).group(1)
            year = re.match(pattern_no_comma, citation).group(2)
            violations.append(APAViolation(
                citation_text=citation,
                error_type=APAErrorType.COMMA_ERROR,
                location=location,
                explanation='Falta coma entre autor y año',
                correction=f'({author}, {year})',
                paragraph_preview=preview
            ))
        
        # Check 3: Lowercase author name
        # Author names should start with capital letter
        author_pattern = r'\(([a-záéíóúñ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+)'
        if re.search(author_pattern, citation):
            violations.append(APAViolation(
                citation_text=citation,
                error_type=APAErrorType.CAPITALIZATION_ERROR,
                location=location,
                explanation='El apellido debe comenzar con mayúscula',
                correction=citation.capitalize(),  # Simplified
                paragraph_preview=preview
            ))
        
        # Check 4: Et al. format errors
        if 'et al' in inner.lower():
            # Check for extra period: "et. al."
            if 'et. al' in inner:
                violations.append(APAViolation(
                    citation_text=citation,
                    error_type=APAErrorType.ET_AL_FORMAT_ERROR,
                    location=location,
                    explanation='Formato incorrecto: debe ser "et al." (sin punto en "et")',
                    correction=citation.replace('et. al', 'et al'),
                    paragraph_preview=preview
                ))
            
            # Check for missing period after "al"
            if re.search(r'et al[,\)]', inner):
                violations.append(APAViolation(
                    citation_text=citation,
                    error_type=APAErrorType.ET_AL_FORMAT_ERROR,
                    location=location,
                    explanation='Falta punto después de "al": debe ser "et al."',
                    correction=citation.replace('et al', 'et al.'),
                    paragraph_preview=preview
                ))
        
        # Check 5: Page format errors
        if 'pág' in inner.lower() or 'página' in inner.lower():
            violations.append(APAViolation(
                citation_text=citation,
                error_type=APAErrorType.PAGE_FORMAT_ERROR,
                location=location,
                explanation='Usar abreviatura en inglés: "p." para página única, "pp." para múltiples',
                correction=citation.replace('pág.', 'p.').replace('págs.', 'pp.'),
                paragraph_preview=preview
            ))
        
        # Check 6: Excessive spacing
        if '  ' in citation:  # Double space
            violations.append(APAViolation(
                citation_text=citation,
                error_type=APAErrorType.SPACING_ERROR,
                location=location,
                explanation='Espaciado excesivo detectado',
                correction=' '.join(citation.split()),
                paragraph_preview=preview
            ))
        
        return violations
    
    def _validate_narrative(self, citation: str, location: int, preview: str = "") -> List[APAViolation]:
        """Validate narrative citation: Author (Year)"""
        violations = []
        
        # Pattern: Author (Year)
        pattern = r'([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-\s]+(?:et al\.)?)[\s]*\((\d{4}[a-z]?)\)'
        match = re.match(pattern, citation)
        
        if not match:
            return violations  # Cannot validate malformed citation
        
        author_part = match.group(1).strip()
        year_part = match.group(2)
        
        # Check 1: Ampersand in author part
        if ' & ' in author_part:
            violations.append(APAViolation(
                citation_text=citation,
                error_type=APAErrorType.CONJUNCTION_ERROR,
                location=location,
                explanation='APA 7 español requiere "y" en lugar de "&" para citas narrativas',
                correction=citation.replace(' & ', ' y '),
                paragraph_preview=preview
            ))
        
        # Check 2: Et al. format
        if 'et al' in author_part.lower():
            if 'et. al' in author_part:
                violations.append(APAViolation(
                    citation_text=citation,
                    error_type=APAErrorType.ET_AL_FORMAT_ERROR,
                    location=location,
                    explanation='Formato incorrecto: debe ser "et al." (sin punto en "et")',
                    correction=citation.replace('et. al', 'et al'),
                    paragraph_preview=preview
                ))
        
        # Check 3: Space before parenthesis
        if not re.search(r'\s\(\d{4}[a-z]?\)', citation):
            violations.append(APAViolation(
                citation_text=citation,
                error_type=APAErrorType.SPACING_ERROR,
                location=location,
                explanation='Debe haber un espacio entre el autor y el año',
                correction=re.sub(r'([A-Za-z])\(', r'\1 (', citation),
                paragraph_preview=preview
            ))
        
        return violations
    
    def validate_all_citations(self, citations: List[Tuple[str, int, str]]) -> List[APAViolation]:
        """
        Validate multiple citations.
        
        Args:
            citations: List of (citation_text, paragraph_index, paragraph_text) tuples
            
        Returns:
            List of all violations found
        """
        all_violations = []
        
        for citation_text, location, paragraph_text in citations:
            violations = self.validate_citation(citation_text, location, paragraph_text)
            all_violations.extend(violations)
        
        return all_violations
    
    def generate_report(self, violations: List[APAViolation]) -> str:
        """
        Generate human-readable validation report.
        
        Args:
            violations: List of violations found
            
        Returns:
            Formatted report string
        """
        if not violations:
            return "✅ No se detectaron errores de formato APA 7"
        
        report = f"\n⚠️  ERRORES DE FORMATO APA 7 DETECTADOS: {len(violations)}\n"
        report += "=" * 80 + "\n"
        
        # Group by error type
        by_type = {}
        for v in violations:
            type_name = v.error_type.value
            if type_name not in by_type:
                by_type[type_name] = []
            by_type[type_name].append(v)
        
        # Display grouped errors
        for error_type, error_list in by_type.items():
            report += f"\n🔴 {error_type.upper()} ({len(error_list)}):\n"
            for i, violation in enumerate(error_list[:5], 1):  # Show max 5 per type
                if violation.paragraph_preview:
                    report += f"   {i}. Ubicación: \"{violation.paragraph_preview}\"\n"
                else:
                    report += f"   {i}. Ubicación: Párrafo {violation.location + 1}\n"
                report += f"      Citación: {violation.citation_text}\n"
                report += f"      Problema: {violation.explanation}\n"
                report += f"      Corrección: {violation.correction}\n"
            
            if len(error_list) > 5:
                report += f"   ... y {len(error_list) - 5} error(es) más de este tipo\n"
        
        report += "\n" + "=" * 80 + "\n"
        
        return report


def validate_apa_citations(citations: List[Tuple[str, int, str]]) -> Tuple[List[APAViolation], str]:
    """
    Convenience function to validate citations and get report.
    
    Args:
        citations: List of (citation_text, paragraph_index) tuples
        
    Returns:
        Tuple of (violations_list, formatted_report)
    """
    validator = APAValidator()
    violations = validator.validate_all_citations(citations)
    report = validator.generate_report(violations)
    
    return violations, report


# Example usage
if __name__ == "__main__":
    # Test cases
    test_citations = [
        ("(García, 2020)", 0),           # ✅ Correct
        ("(García & Pérez, 2020)", 1),   # ❌ Wrong conjunction
        ("(garcía, 2020)", 2),           # ❌ Lowercase
        ("(García 2020)", 3),            # ❌ Missing comma
        ("(García et. al., 2020)", 4),   # ❌ Wrong et al.
        ("García (2020)", 5),            # ✅ Correct narrative
        ("García & Pérez (2020)", 6),    # ❌ Wrong conjunction
        ("(García, 2020, pág. 5)", 7),   # ❌ Wrong page format
    ]
    
    validator = APAValidator()
    violations, report = validate_apa_citations(test_citations)
    
    print(report)
    print(f"\nTotal violations: {len(violations)}")
