"""
Citation Matcher - Matches in-text citations with reference list
"""

import re
from typing import List
from domain.models import Citation, Reference
from domain.enums import SeverityLevel


class CitationMatcher:
    """Matches in-text citations with reference list."""
    
    def __init__(self, citations: List[Citation], references: List[Reference]):
        self.citations = citations
        self.references = references
        
        # Build lookup keys
        self.citation_keys = {}
        for cit in citations:
            key = cit.text
            self.citation_keys[key] = cit
        
        self.reference_keys = {}
        for ref in references:
            key = self._ref_key(ref)
            self.reference_keys[key] = ref
    

    def extract_all_citations(self, doc_path: str, citation_parser) -> List[Citation]:
        """Extract both in-text and footnote citations."""
        from docx import Document
        
        all_citations = []
        
        # 1. Extract footnote citations
        doc = Document(doc_path)
        footnote_citations = citation_parser.extract_footnotes(doc)
        all_citations.extend(footnote_citations)
        
        # 2. Extract in-text citations from paragraphs
        for i, para in enumerate(doc.paragraphs):
            text_citations = citation_parser.parse(para.text, i)
            all_citations.extend(text_citations)
        
        return all_citations

    @staticmethod
    def _ref_key(reference: Reference) -> str:
        """Generate matching key with smart organizational handling."""
        # Remove bullets/dashes
        clean_text = re.sub(r'^[-–—•]+\s*', '', reference.text)
        
        # Try to match organizational pattern with year
        org_pattern = r'^([A-ZÁ-ÚÑ][A-Za-záéíóúñ\s&,\-]{5,}?)\s+\((\d{4}[a-z]?)'
        match = re.search(org_pattern, clean_text)
        
        if match:
            org_name = match.group(1).strip()
            year = match.group(2)
            
            # Check for abbreviation in parentheses
            abbrev_match = re.search(r'[-–—]\s*([A-Z]{2,})\s*[-–—]', org_name)
            if abbrev_match:
                # Use abbreviation: "Central Intelligence Agency -CIA-" → "cia"
                key_word = abbrev_match.group(1).lower()
            else:
                # Use last word: "Ministerio de Economía" → "economía"
                words = org_name.split()
                skip_words = {'de', 'del', 'la', 'el', 'los', 'las', 'y'}
                significant_words = [w for w in words if w.lower() not in skip_words]
                key_word = significant_words[-1].lower() if significant_words else words[-1].lower()
            
            return f"{key_word}_{year}"
        
        # Fallback: personal author
        match = re.match(r'([A-ZÁ-ÚÑ][a-zá-úñ\-]+)', clean_text)
        if match:
            first_author = match.group(1).lower()
            year = reference.year if hasattr(reference, "year") else "n.d."
            return f"{first_author}_{year if year else 'unknown'}"
        
        return "unknown_unknown"
    
    def find_orphaned_citations(self) -> List[Citation]:
        """Citations without matching references."""
        orphaned = []
        for cit in self.citations:
            if cit.text not in self.reference_keys:
                orphaned.append(cit)
        return orphaned
    
    def find_orphaned_references(self) -> List[Reference]:
        """References never cited in text."""
        orphaned = []
        for ref in self.references:
            if self._ref_key(ref) not in self.citation_keys:
                orphaned.append(ref)
        return orphaned
    
    def generate_report(self, section_type: str) -> str:
        """Generate citation integrity report."""
        report = []
        report.append("\n" + "=" * 70)
        report.append("INTEGRIDAD DE CITAS Y REFERENCIAS")
        report.append("=" * 70)
        
        report.append(f"Citas en texto: {len(self.citations)}")
        report.append(f"Referencias bibliográficas: {len(self.references)}")
        report.append(f"Tipo de sección: {section_type.upper()}")
        
        # Orphaned citations (ALWAYS CRITICAL)
        orphaned_cits = self.find_orphaned_citations()
        if orphaned_cits:
            report.append(f"\n{SeverityLevel.CRITICO.value}: Citas Sin Referencia")
            report.append(f"Encontradas {len(orphaned_cits)} citas sin entrada bibliográfica:")
            for cit in orphaned_cits[:5]:
                report.append(f"  • {cit}")
            if len(orphaned_cits) > 5:
                report.append(f"  ... y {len(orphaned_cits) - 5} más")
        
        # Orphaned references (severity depends on section type)
        orphaned_refs = self.find_orphaned_references()
        if orphaned_refs:
            if section_type == "Referencias":
                severity = SeverityLevel.ADVERTENCIA
                msg = "En 'Referencias', se espera citar todas las entradas."
            else:
                severity = SeverityLevel.INFORMATIVO
                msg = "En 'Bibliografía', es aceptable incluir fuentes consultadas."
            
            report.append(f"\n{severity.value}: Referencias Sin Citar")
            report.append(msg)
            report.append(f"Encontradas {len(orphaned_refs)} referencias no citadas:")
            for ref in orphaned_refs[:5]:
                report.append(f"  • {ref.text[:60]}...")
            if len(orphaned_refs) > 5:
                report.append(f"  ... y {len(orphaned_refs) - 5} más")
        
        # Final verdict
        if not orphaned_cits and not orphaned_refs:
            report.append(f"\n✅ Sistema de citación íntegro")
        elif not orphaned_cits:
            report.append(f"\n✅ Todas las citas tienen referencia válida")
        
        return '\n'.join(report)
    
    def match_citations_to_references(self, section_type: str):
        from domain.models import CitationAnalysisResult

        orphaned_citations = self.find_orphaned_citations()

        return CitationAnalysisResult(
            total_citations=len(self.citations),
            total_references=len(self.references),
            matched_count=len(self.citations) - len(orphaned_citations),
            unmatched_count=len(orphaned_citations),
            citations_by_type={},
            unmatched_citations=[c.text for c in orphaned_citations]
        )
  
    
