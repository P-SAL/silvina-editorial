"""
Citation Parser - Extracts APA citations from Spanish text
FINAL WORKING VERSION - Extracts from XML with proper HTML decoding
"""

import re
import html
import zipfile
from typing import List, Optional
from domain.models import Citation
from domain.enums import CitationType


class CitationParser:
    """Extracts APA citations from Spanish text using XML-based extraction."""
    
    def __init__(self):
        """Initialize the citation parser."""
        pass
   
    def extract_from_docx(self, docx_path: str) -> List[Citation]:
        """
        Extract citations directly from DOCX XML with proper HTML decoding.
        
        Args:
            docx_path: Path to .docx file
            
        Returns:
            List of Citation objects
        """
        try:
            # Extract raw XML from the DOCX
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                doc_xml = zip_ref.read('word/document.xml').decode('utf-8')
            
            # Parse XML properly
            import xml.etree.ElementTree as ET
            root = ET.fromstring(doc_xml)
            
            # Define namespace
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            
            # Extract text paragraph by paragraph
            paragraphs = []
            for para in root.findall('.//w:p', ns):
                texts = []
                for t in para.findall('.//w:t', ns):
                    if t.text:
                        texts.append(html.unescape(t.text))
                para_text = ''.join(texts)  # Join within paragraph (no spaces needed)
                if para_text.strip():
                    paragraphs.append(para_text)
            
            # Join paragraphs with space
            full_text = ' '.join(paragraphs)
            
                       
            # Extract citations from full text
            return self._extract_citations(full_text)
            
        except Exception as e:
            print(f"      ⚠️ Error extracting citations from XML: {e}")
            return []

    def _extract_citations(self, full_text: str) -> List[Citation]:
        """
        Extract all APA citations from document text.
        
        Handles both parenthetical and narrative citations.
        """
        citations = []
        seen = set()  # Avoid duplicates
        
        # ========== PARENTHETICAL CITATIONS ==========
        # Pattern: Find everything with a 4-digit year in parentheses
        # (Author, 2020) | (Author1 & Author2, 2020) | (Author et al., 2020)
        
        all_parenthetical = re.findall(r'\([^)]*(?:19|20)\d{2}[^)]*\)', full_text)
        
        for pattern in all_parenthetical:
            # Skip date-only patterns like (5 de abril de 2021)
            if re.match(r'^\(\d+\s+de\s+', pattern):
                continue
            
            # Extract year
            year_match = re.search(r'(\d{4}[a-z]?)', pattern)
            if not year_match:
                continue
            
            year = year_match.group(1)
            
            # Extract author - everything before the year
            author_part = pattern[1:].split(year)[0].strip().rstrip(',').strip()
            
            # Must have valid author name
            if not author_part or len(author_part) < 2:
                continue
            
            citation_key = f"{author_part}|{year}"
            if citation_key not in seen:
                seen.add(citation_key)
                citations.append(Citation(
                    text=pattern,
                    citation_type=CitationType.AUTHOR_YEAR,
                    location=-1,
                    author=author_part,
                    year=year
                ))
        
        # ========== NARRATIVE CITATIONS ==========
        
        # Initialize tracking dictionaries
        multi_author_names = {}
        first_authors_by_year = {}
        
        # Pre-process: Extract first authors from parenthetical multi-author citations
        # This prevents false positives like "Dhingra (2021)" when we already have 
        # "(Dhingra, Samo, Schaninger, & Schrimper, 2021)"
        for cite in citations:
            if cite.text.startswith('(') and ('&' in cite.author or ',' in cite.author):
                # Multi-author parenthetical - get first author
                first_author = re.match(r'([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)', cite.author)
                if first_author:
                    if cite.year not in first_authors_by_year:
                        first_authors_by_year[cite.year] = set()
                    first_authors_by_year[cite.year].add(first_author.group(1))
        
        # Pattern 1: Multi-author narrative citations
        # "Craig y Snook (2023)" | "Hansen y Keltner (2023)" | "Ramos e Iñaki Vélaez (2023)"
        
        narrative_multi = r'(?<![a-záéíóúñ])\b([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ]+(?:\s+[ye&]\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\s]+?)+)\s+\((\d{4}[a-z]?)\)'
        
        for match in re.finditer(narrative_multi, full_text):
            author = match.group(1).strip()
            year = match.group(2)
            
            # Skip if too long (captured too much context)
            if len(author) > 100:
                continue
            
            # Skip if starts with common intro phrases
            if re.match(r'^(Como|Según|Si|No|En|El|La|Los|Las|Un|Una)\s', author, re.IGNORECASE):
                continue
            
            citation_key = f"{author}|{year}"
            if citation_key not in seen:
                seen.add(citation_key)
                
                # Store all individual authors from this citation to avoid duplicates
                individual_authors = re.findall(r'[A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ]+', author)
                if year not in multi_author_names:
                    multi_author_names[year] = set()
                multi_author_names[year].update(individual_authors)
                
                citations.append(Citation(
                    text=f"{author} ({year})",
                    citation_type=CitationType.AUTHOR_YEAR,
                    location=-1,
                    author=author,
                    year=year
                ))
        
        # Pattern 2: Single author narrative citations
        # "Coleman (2023)" | "Schein (1982)"
        # But NOT if they're already part of a multi-author citation

        single_narrative = r'(?<![(\[])\b([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)\s+\((\d{4}[a-z]?)\)'

        
        for match in re.finditer(single_narrative, full_text):
            author = match.group(1).strip()
            year = match.group(2)
            
            # Check 1: First author of multi-work?
            if year in first_authors_by_year and author in first_authors_by_year[year]:
                continue
            
            # Check 2: Already in a multi-author narrative?
            if year in multi_author_names and author in multi_author_names[year]:
                continue
            
            # Check 3: Already added?
            citation_key = f"{author}|{year}"
            if citation_key in seen:
                continue
            
            # All checks passed - add it!
            seen.add(citation_key)
            citations.append(Citation(
                text=f"{author} ({year})",
                citation_type=CitationType.AUTHOR_YEAR,
                location=-1,
                author=author,
                year=year
            ))
              
        return citations


    def extract_footnotes(self, doc) -> List[Citation]:
        """
        Extract Word footnote references from document.
        
        Args:
            doc: python-docx Document object
            
        Returns:
            List of Citation objects representing footnotes
        """
        citations = []
        seen_numbers = set()
        
        # Search for Word footnotes in the document XML
        for i, para in enumerate(doc.paragraphs):
            if hasattr(para, '_element'):
                footnote_refs = para._element.findall(
                    './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnoteReference'
                )
                for ref in footnote_refs:
                    fn_id = ref.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    if fn_id and fn_id not in seen_numbers:
                        seen_numbers.add(fn_id)
                        citations.append(Citation(
                            text=f"[Footnote {fn_id}]",
                            citation_type=CitationType.FOOTNOTE,
                            location=i,
                            author=None,
                            year=None
                        ))
        
        if len(citations) > 0:
            print(f"      ✓ {len(citations)} notas al pie detectadas")
            print(f"      ⚠️  Las notas al pie NO son citas bibliográficas")
        
        return citations

    def parse(self, text: str, paragraph_index: int) -> List[Citation]:
        """
        Extract citations from one paragraph (FALLBACK METHOD).
        
        NOTE: This method is kept for compatibility with existing code,
        but extract_from_docx() is strongly preferred because it handles
        HTML entities correctly.
        
        Args:
            text: Paragraph text
            paragraph_index: Index of paragraph in document
            
        Returns:
            List of Citation objects
        """
        citations = []
        
        # Parenthetical citations
        all_parenthetical = re.findall(r'\([^)]*(?:19|20)\d{2}[^)]*\)', text)
        
        for pattern in all_parenthetical:
            if re.match(r'^\(\d+\s+de\s+', pattern):
                continue
            
            year_match = re.search(r'(\d{4}[a-z]?)', pattern)
            if not year_match:
                continue
            
            year = year_match.group(1)
            author_part = pattern[1:].split(year)[0].strip().rstrip(',').strip()
            
            if author_part and len(author_part) > 2:
                citations.append(Citation(
                    text=pattern,
                    citation_type=CitationType.AUTHOR_YEAR,
                    location=paragraph_index,
                    author=author_part,
                    year=year
                ))
        
        # Narrative citations (simplified for single paragraph)
        narrative_pattern = r'\b([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\s]+?)\s+\((\d{4}[a-z]?)\)'
        
        for match in re.finditer(narrative_pattern, text):
            author = match.group(1).strip()
            year = match.group(2)
            
            if len(author) < 100:  # Sanity check
                citations.append(Citation(
                    text=f"{author} ({year})",
                    citation_type=CitationType.AUTHOR_YEAR,
                    location=paragraph_index,
                    author=author,
                    year=year
                ))
        
        return citations