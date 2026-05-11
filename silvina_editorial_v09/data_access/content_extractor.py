"""
content_extractor.py
Extracts structured content from document paragraphs.
FIXED VERSION - Returns "Autor no identificado" + blacklist headers
"""

from typing import List, Dict, Optional
import re
from domain.models import DocumentContent
from data_access.word_counter import WordCounter, WIN32COM_AVAILABLE


class ContentExtractor:
    """Extracts structured content from raw document paragraphs."""
    
    # BLACKLIST: Common section headers that are NOT authors
    AUTHOR_BLACKLIST = {
        'RESUMEN', 'ABSTRACT', 'INTRODUCCIÓN', 'INTRODUCTION',
        'METODOLOGÍA', 'METHODOLOGY', 'MÉTODOS', 'METHODS',
        'RESULTADOS', 'RESULTS', 'DISCUSIÓN', 'DISCUSSION',
        'CONCLUSIONES', 'CONCLUSIONS', 'CONCLUSIÓN', 'CONCLUSION',
        'REFERENCIAS', 'REFERENCES', 'BIBLIOGRAFÍA', 'BIBLIOGRAPHY',
        'PALABRAS CLAVE', 'KEYWORDS', 'AGRADECIMIENTOS', 'ACKNOWLEDGMENTS',
        'APÉNDICE', 'APPENDIX', 'ANEXO', 'ANNEX',
        'TABLA', 'TABLE', 'FIGURA', 'FIGURE',
        'ÍNDICE', 'INDEX', 'CONTENIDO', 'CONTENTS'
    }
    
    def __init__(self):
        """Initialize the content extractor."""
        self.section_patterns = {
            'title': r'^(?:TÍTULO|TITLE)[:\s]*(.*)',
            'abstract': r'^(?:RESUMEN|ABSTRACT)[:\s]*(.*)',
            'keywords': r'^(?:PALABRAS CLAVE|KEYWORDS)[:\s]*(.*)',
            'authors': r'^(?:AUTORES|AUTOR|AUTHORS|AUTHOR)[:\s]*(.*)',
            
        }
    
    def extract_content(self, paragraphs: List[str], docx_path: str = None) -> DocumentContent:
        """
        Extract structured content from document paragraphs.
        
        Args:
            paragraphs: List of paragraph objects or strings
            docx_path: Optional path to .docx file for accurate counting
        """

        # 1. Normalize paragraphs
        clean_paragraphs = [
            p.text.strip() if hasattr(p, "text") else str(p).strip()
            for p in paragraphs
            if (p.text if hasattr(p, "text") else str(p)).strip()
        ]

        if not clean_paragraphs:
            raise ValueError("No valid paragraphs after cleaning")

        # 2. Get initial counts from text
        word_count = sum(len(p.split()) for p in clean_paragraphs)
        char_count = sum(len(p) for p in clean_paragraphs)
        paragraph_count = len(clean_paragraphs)

        # 3. Try to get accurate Word counts
        if docx_path and WIN32COM_AVAILABLE:
            counter = WordCounter()
            accurate_counts = counter.get_accurate_counts(docx_path)
            if accurate_counts:
                word_count = accurate_counts['word_count']
                char_count = accurate_counts['char_count']
                paragraph_count = accurate_counts['paragraph_count']
        
        # 4. Extract structured fields        
        title = self._extract_title(clean_paragraphs)
        # If title contains ' — ' it was built from 2 paragraphs
        title_lines = 2 if title and ' — ' in title else 1
        # If first paragraph was an institution header (skipped in title extraction), add 1
        if clean_paragraphs and re.search(
            r'^(?:Universidad|Facultad|Escuela|Instituto|Centro|Ministerio|Comando|Departamento)',
            clean_paragraphs[0], re.IGNORECASE):
            title_lines += 1
        authors = self._extract_authors(clean_paragraphs, title_lines)
        abstract = self._extract_abstract(clean_paragraphs)
        keywords = self._extract_keywords(clean_paragraphs)
        sections = self._extract_sections(clean_paragraphs)

        sections = self._extract_sections(clean_paragraphs)

        # 5. Extract references
        references = []
        if docx_path:
            from data_access.reference_parser import ReferenceParser
            references, _ = ReferenceParser().parse_from_docx(docx_path)

        # 6. Return DocumentContent
        return DocumentContent(
            title=title,
            authors=authors,
            abstract=abstract,
            keywords=keywords,
            sections=sections,
            references=references,
            paragraphs=clean_paragraphs,
            word_count=word_count,
            char_count=char_count,
            paragraph_count=paragraph_count
        )

    def _extract_title(self, paragraphs: List[str]) -> Optional[str]:
        """Extract document title - combines first two short paragraphs if both look like title parts."""
        title_parts = []
        
        for para in paragraphs[:5]:
            # Stop if explicit title marker
            match = re.match(self.section_patterns['title'], para, re.IGNORECASE)
            if match:
                return match.group(1).strip()
            
            # Skip institution/organization headers
            if re.search(r'^(?:Universidad|Facultad|Escuela|Instituto|Centro|Ministerio|Comando|Departamento)', para, re.IGNORECASE):
                continue

            # Collect short paragraphs that look like title lines
            if len(para.split()) >= 2 and len(para) < 200:
                title_parts.append(para.strip())
                if len(title_parts) == 2:
                    break
        
        if len(title_parts) >= 2:
            # Don't combine if second part looks like an author name
            second = title_parts[1]
            looks_like_author = (
                len(second.split()) <= 10 and
                not any(c in second for c in ['—', '?', ':', 'de', 'del', 'para']) and
                (re.search(r'^(?:Dr|Dra|Lic|Mag|CF|CN|CNVGM|Prof)', second) or
                 re.search(r'^[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+\s*$', second) or
                 re.search(r'^[A-Z]{2,}', second))
            ) 
            
            if looks_like_author:
                return title_parts[0]
            part0 = title_parts[0].rstrip(':').strip()
            return f"{part0} — {title_parts[1]}"
           
        elif len(title_parts) == 1:
            return title_parts[0]
                  
        return None

    def _extract_authors(self, paragraphs: List[str], title_lines: int = 1) -> Optional[str]:
        """
        Extract author information from first page after title.
        FIXED: Returns "Autor no identificado" if not found.
        """
        
        # Skip first paragraph (usually title)
        for i, para in enumerate(paragraphs[title_lines:15], start=title_lines):
            para_stripped = para.strip()
            para_upper = para_stripped.upper()
            
            # BLACKLIST CHECK: Skip if it's a section header
            if any(header in para_upper for header in self.AUTHOR_BLACKLIST):
                continue
            
            # Pattern 1: Explicit "Autor:" or "Author:"
            match = re.match(self.section_patterns['authors'], para, re.IGNORECASE)
            if match:
                author_text = match.group(1).strip()
                if author_text and not any(bl in author_text.upper() for bl in self.AUTHOR_BLACKLIST):
                    return author_text
                # Authors are on the next lines — collect them
                author_lines = []
                for next_para in paragraphs[i+1:i+6]:
                    if re.match(r'^[A-ZÁÉÍÓÚÑ]', next_para) and len(next_para.split()) <= 10:
                        if any(header in next_para.upper() for header in self.AUTHOR_BLACKLIST):
                            break
                        author_lines.append(next_para.strip())
                    else:
                        break
                if author_lines:
                    return ', '.join(author_lines)
            
            # Pattern 1b: Parenthetical author format
            # Matches: "(Director de Proyecto Com. (R) Mág Pablo Andrés Farias)1"
            paren_match = re.search(r'\((?:Director|Autor|Investigador|Coordinador).*?([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,3})\s*\)\d*', para_stripped, re.IGNORECASE)
            if paren_match:
                return paren_match.group(1).strip()

            # Pattern 2: Name-like pattern after title
            # Matches: "Adriana Baravalle" or "Juan Pérez, María González"
            if i <= 3:  # Check first 3 paragraphs after title
       
                # Check if looks like a name
                if (len(para.split()) <= 10 and 
                    para[0].isupper() and 
                    not para.isupper() and  # Not all caps
                    not para.endswith(':') and  # Not a label
                    re.search(r'^[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+', para)):
                    
                    # Final blacklist check
                    if not any(bl in para_upper for bl in self.AUTHOR_BLACKLIST):
                        return para_stripped
        
        # NOT FOUND: Return standard message
        return "Autor no identificado"

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
