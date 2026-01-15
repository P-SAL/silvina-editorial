"""
content_extractor.py
Extracts structured content from document paragraphs.
Part of Silvina Editorial Assistant v0.7
"""

from typing import List, Dict, Optional
import re
from domain.models import DocumentContent


class ContentExtractor:
    """Extracts structured content from raw document paragraphs."""
    
    def __init__(self):
        """Initialize the content extractor."""
        self.section_patterns = {
            'title': r'^(?:TÍTULO|TITLE)[:\s]*(.*)',
            'abstract': r'^(?:RESUMEN|ABSTRACT)[:\s]*(.*)',
            'keywords': r'^(?:PALABRAS CLAVE|KEYWORDS)[:\s]*(.*)',
            'authors': r'^(?:AUTOR|AUTORES|AUTHOR|AUTHORS)[:\s]*(.*)',
        }
    
    def extract_content(self, paragraphs: List[str]) -> DocumentContent:
        """
        Extract structured content from document paragraphs.
        
        Args:
            paragraphs: List of paragraph texts
            
        Returns:
            DocumentContent object with extracted information
        """
        content = DocumentContent()
        
        # Extract title (usually first non-empty paragraph or marked with TÍTULO)
        content.title = self._extract_title(paragraphs)
        
        # Extract authors
        content.authors = self._extract_authors(paragraphs)
        
        # Extract abstract
        content.abstract = self._extract_abstract(paragraphs)
        
        # Extract keywords
        content.keywords = self._extract_keywords(paragraphs)
        
        # Store all paragraphs
        content.paragraphs = paragraphs
        
        # Extract sections
        content.sections = self._extract_sections(paragraphs)
        
        # Calculate word count
        content.word_count = sum(len(p.split()) for p in paragraphs)
        
        return content
    
    def _extract_title(self, paragraphs: List[str]) -> Optional[str]:
        """Extract document title."""
        for para in paragraphs[:10]:  # Check first 10 paragraphs
            # Check for explicit title marker
            match = re.match(self.section_patterns['title'], para, re.IGNORECASE)
            if match:
                return match.group(1).strip()
            
            # If first substantial paragraph (likely title)
            if len(para.split()) >= 3 and len(para) < 200:
                return para.strip()
        
        return None
    
    def _extract_authors(self, paragraphs: List[str]) -> Optional[str]:
        """Extract author information."""
        for para in paragraphs[:15]:
            match = re.match(self.section_patterns['authors'], para, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        return None
    
    def _extract_abstract(self, paragraphs: List[str]) -> Optional[str]:
        """Extract abstract/resumen."""
        abstract_lines = []
        in_abstract = False
        
        for para in paragraphs:
            # Check if this is the abstract header
            if re.match(r'^(?:RESUMEN|ABSTRACT)\s*$', para, re.IGNORECASE):
                in_abstract = True
                continue
            
            # If we're in abstract section
            if in_abstract:
                # Stop at next section header
                if re.match(r'^[A-Z\sÁÉÍÓÚÑ]{3,}$', para):
                    break
                abstract_lines.append(para)
                
                # Stop after reasonable abstract length
                if len(' '.join(abstract_lines).split()) > 300:
                    break
        
        return ' '.join(abstract_lines) if abstract_lines else None
    
    def _extract_keywords(self, paragraphs: List[str]) -> List[str]:
        """Extract keywords."""
        for para in paragraphs:
            match = re.match(self.section_patterns['keywords'], para, re.IGNORECASE)
            if match:
                keywords_text = match.group(1).strip()
                # Split by common separators
                keywords = re.split(r'[;,]', keywords_text)
                return [kw.strip() for kw in keywords if kw.strip()]
        
        return []
    
    def _extract_sections(self, paragraphs: List[str]) -> Dict[str, str]:
        """Extract document sections."""
        sections = {}
        current_section = None
        current_content = []
        
        section_headers = [
            'INTRODUCCIÓN', 'INTRODUCTION',
            'METODOLOGÍA', 'METHODOLOGY', 'MÉTODOS', 'METHODS',
            'RESULTADOS', 'RESULTS',
            'DISCUSIÓN', 'DISCUSSION',
            'CONCLUSIONES', 'CONCLUSIONS',
            'REFERENCIAS', 'REFERENCES', 'BIBLIOGRAFÍA', 'BIBLIOGRAPHY'
        ]
        
        for para in paragraphs:
            # Check if this is a section header
            para_upper = para.strip().upper()
            is_header = any(header in para_upper for header in section_headers)
            
            if is_header and len(para.split()) <= 5:
                # Save previous section
                if current_section and current_content:
                    sections[current_section] = '\n'.join(current_content)
                
                # Start new section
                current_section = para_upper
                current_content = []
            elif current_section:
                # Add to current section
                current_content.append(para)
        
        # Save last section
        if current_section and current_content:
            sections[current_section] = '\n'.join(current_content)
        
        return sections


# Convenience function
def extract_content_from_paragraphs(paragraphs: List[str]) -> DocumentContent:
    """
    Extract content from paragraphs.
    
    Args:
        paragraphs: List of paragraph texts
        
    Returns:
        DocumentContent object
    """
    extractor = ContentExtractor()
    return extractor.extract_content(paragraphs)