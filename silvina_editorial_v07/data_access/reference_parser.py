"""
Reference Parser - Extracts bibliography references
"""

import re
from typing import List, Tuple
from domain.models import Reference


class ReferenceParser:
    """Parses bibliography section and creates Reference objects."""
    
    def parse_section(self, text: str) -> Tuple[List[Reference], str]:
        """
        Parse bibliography text into Reference objects.
        Returns: (list of References, section_type)
        """
        # Determine section type
        section_type = "Referencias"
        if "Bibliografía" in text or "Fuentes bibliográficas" in text:
            section_type = "Bibliografía"
        
        # Split into paragraphs
        paragraphs = text.split('\n')
        references = []
        
        for para in paragraphs:
            para = para.strip()
            
            # Skip short lines (headers, empty lines)
            if len(para) < 30:
                continue
            
            # Check if paragraph has multiple references (by counting years)
            years = re.findall(r'\(\d{4}\)', para)
            
            if len(years) >= 2:
                # Try to split by period before capital letter
                split_pattern = r'\.(?=[A-Z][a-z]+,\s+[A-Z]\.)'
                parts = re.split(split_pattern, para, maxsplit=1)
                
                for part in parts:
                    part = part.strip()
                    if len(part) > 30:
                        if not part.endswith('.'):
                            part += '.'
                        references.append(Reference(part))
            else:
                references.append(Reference(para))
        
        return references, section_type
    
    def extract_from_paragraphs(self, paragraphs: List[str]) -> Tuple[str, str]:
        """
        Extract bibliography section from document paragraphs.
        Returns: (bibliography_text, section_type)
        """
        found_start = False
        section_type = "Referencias"
        referencias_paras = []
        
        for para in paragraphs:
            para_text = para.strip()
            
            if not found_start:
                # Check for section headers
                if "Bibliografía" in para_text or "Fuentes bibliográficas" in para_text:
                    section_type = "Bibliografía"
                    found_start = True
                    continue
                elif "Referencias" in para_text and "bibliográficas" in para_text:
                    section_type = "Referencias"
                    found_start = True
                    continue
            
            if found_start and para_text and len(para_text) > 30:
                referencias_paras.append(para_text)
        
        bibliography_text = '\n'.join(referencias_paras)
        return bibliography_text, section_type