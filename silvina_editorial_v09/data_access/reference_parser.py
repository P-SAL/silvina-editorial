"""
Reference Parser - Extracts bibliography references
FINAL WORKING VERSION - Reads from XML with proper HTML decoding
"""

import re
import html
import zipfile
from typing import List, Tuple, Optional
from domain.models import Reference


class ReferenceParser:
    """Parses bibliography section and creates Reference objects."""
    
    def parse_from_docx(self, docx_path: str) -> Tuple[List[Reference], str]:
        """
        Extract bibliography directly from DOCX XML.
        
        Args:
            docx_path: Path to .docx file
            
        Returns:
            Tuple of (list of References, section_type string)
        """
        try:
            # Extract text from document.xml
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                doc_xml = zip_ref.read('word/document.xml').decode('utf-8')
            
            # Extract all text elements
            text_pattern = r'<w:t[^>]*>([^<]+)</w:t>'
            all_texts = re.findall(text_pattern, doc_xml)
            
            # Decode HTML entities (&amp; -> &)
            all_texts = [html.unescape(text) for text in all_texts]
            
            # Join into full text
            full_text = ' '.join(all_texts)
            
            # Find Bibliografia/Referencias section
            bib_match = re.search(r'(Bibliograf[íi]a|Referencias)\s+(.{100,})', full_text, re.IGNORECASE | re.DOTALL)
            
            if not bib_match:
                return [], "Referencias"
            
            section_type = "Bibliografía" if "ibliograf" in bib_match.group(1).lower() else "Referencias"
            bib_text = bib_match.group(2)
            
            # Parse references
            references = self._parse_references(bib_text)
            
            return references, section_type
            
        except Exception as e:
            print(f"      ⚠️  Error extracting bibliography from XML: {e}")
            return [], "Referencias"
    
    def _parse_references(self, bib_text: str) -> List[Reference]:
        """Parse bibliography text into individual references.
        
        Splits by year pattern, then cleans up leftover text from previous reference.
        """
        references = []
        
        # Pattern: (Year). or (DD de Month de Year).
        year_end_pattern = r'\((?:\d{1,2}\s+de\s+\w+\s+de\s+)?\d{4}[a-z]?\)\.?'
        
        # Split by year-end pattern but keep the match
        parts = re.split(f'({year_end_pattern})', bib_text)
        
        # Reconstruct references by pairing parts
        current_ref = ""
        for i, part in enumerate(parts):
            if re.match(year_end_pattern, part):
                # This is a year ending - complete the reference
                current_ref += part
                current_ref = current_ref.strip()
                
                if len(current_ref) > 30:
                    # Clean: Find where THIS reference's author starts
                    # Pattern: Author, I. (Year) at the beginning or after previous ref
                    # Look for: Capital letter + name + comma + initial
                    author_match = re.search(
                        r'([A-ZÁ-ÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁ-ÚÑ]?[a-záéíóúñ]+)*,\s+[A-ZÁÉÍÓÚÑ]\..*)',
                        current_ref
                    )
                    
                    if author_match:
                        # Keep only from the author onwards
                        clean_ref = author_match.group(1).strip()
                    else:
                        # Fallback: keep as is
                        clean_ref = current_ref
                    
                    # Remove bullets/dashes
                    clean_ref = re.sub(r'^[-–—•]+\s*', '', clean_ref)
                    references.append(Reference(clean_ref))
                
                current_ref = ""
            else:
                # This is reference content
                current_ref += part
        
        # Don't forget the last reference
        if current_ref.strip() and len(current_ref.strip()) > 30:
            # Same cleanup
            author_match = re.search(
                r'([A-ZÁ-ÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁ-ÚÑ]?[a-záéíóúñ]+)*,\s+[A-ZÁÉÍÓÚÑ]\..*)',
                current_ref
            )
            if author_match:
                clean_ref = author_match.group(1).strip()
            else:
                clean_ref = current_ref.strip()
            
            clean_ref = re.sub(r'^[-–—•]+\s*', '', clean_ref)
            references.append(Reference(clean_ref))
        
        return references


    def parse_section(self, text: Optional[str]) -> Tuple[List[Reference], str]:
        """
        Parse bibliography text into Reference objects (compatibility method).
        
        Args:
            text: Bibliography section text
            
        Returns:
            Tuple of (list of References, section_type)
        """
        if not text:
            return [], "Referencias"
        
        section_type = "Referencias"
        if "Bibliografía" in text or "BIBLIOGRAFÍA" in text:
            section_type = "Bibliografía"
        
        references = self._parse_references(text)
        return references, section_type
   
    def extract_from_paragraphs(self, paragraphs: List) -> Tuple[Optional[str], str]:
        """
        Extract bibliography section from document paragraphs (FALLBACK METHOD).
        
        NOTE: This is kept for compatibility but parse_from_docx() is preferred.
        
        Returns:
            Tuple of (bibliography_text or None, section_type)
        """
        found_start = False
        section_type = "Referencias"
        referencias_paras = []
        
        # Convert to text
        para_texts = []
        for para in paragraphs:
            if hasattr(para, 'text'):
                para_texts.append(para.text.strip())
            else:
                para_texts.append(str(para).strip())
        
        # Search for section header
        HEADER_PATTERNS = [
            r'^\s*BIBLIOGRAF[ÍI]A\s*$',
            r'^\s*REFERENCIAS\s*(?:BIBLIOGR[ÁA]FICAS?)?\s*$',
        ]
        
        for i, para_text in enumerate(para_texts):
            if not para_text:
                continue
            
            # Check for header
            if not found_start:
                for pattern in HEADER_PATTERNS:
                    if re.match(pattern, para_text, re.IGNORECASE):
                        section_type = "Bibliografía" if "IBLIOGRAF" in para_text.upper() else "Referencias"
                        found_start = True
                        break
                continue
            
            # Collect paragraphs
            if found_start:
                if len(para_text) < 20:
                    continue
                
                # Stop at next major section (all caps)
                if para_text.isupper() and len(para_text) > 5:
                    break
                
                referencias_paras.append(para_text)
        
        if not referencias_paras:
            return None, section_type
        
        return '\n'.join(referencias_paras), section_type
