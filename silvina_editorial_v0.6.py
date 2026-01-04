# silvina_editorial_v0.6.py
"""
SILVINA Editorial Assistant v0.6
Citation Integrity & IMRyD Validation
Universidad de la Defensa Nacional
"""

from dataclasses import dataclass
from typing import List, Optional
import re
from pathlib import Path
import sys

# Try to import pywin32, but don't fail if not available
try:
    import win32com.client as win32
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    print("⚠️ pywin32 no instalado - modo documento deshabilitado")


# ============================================================
# CITATION DATA CLASS
# ============================================================

@dataclass
class Citation:
    """Stores one citation with its location in the document."""
    
    authors: List[str]
    year: str
    paragraph_index: int
    citation_type: str
    raw_text: str
    page: Optional[str] = None
    start_pos: int = 0
    is_secondary: bool = False           
    secondary_source: Optional[str] = None  
    
    def __repr__(self):
        """Show citation in readable format."""
        authors_text = " y ".join(self.authors)
        page_text = f", p. {self.page}" if self.page else ""
        type_marker = "📖" if self.citation_type == "narrativa" else "📎"
        
        # Show secondary citation marker
        if self.is_secondary:
            return f"🔗 {authors_text} ({self.year}{page_text}) [como se cita en {self.secondary_source}] [¶{self.paragraph_index}]"
        
        return f"{type_marker} {authors_text} ({self.year}{page_text}) [¶{self.paragraph_index}]"

@dataclass
class Reference:
    """Stores one bibliographic reference with parsed metadata."""
    
    authors: List[str]        # ["García", "López"] or ["IBM Research"]
    year: str                 # "2020" or "2020a"
    title: str                # Article/book title
    raw_text: str            # Full reference text
    paragraph_index: int     # Which paragraph it appears in
    
    def __post_init__(self):
        """Normalize author names (remove extra whitespace)."""
        self.authors = [a.strip() for a in self.authors if a.strip()]
    
    @property
    def reference_key(self) -> str:
        """Generate unique key for matching (first_author_year)."""
        if not self.authors:
            return "unknown_unknown"
        
        # Get first author's last name
        first_author = self.authors[0]
        # Remove initials like "García, A." → "García"
        last_name = first_author.split(',')[0].strip()
        
        return f"{last_name}_{self.year}".lower()
    
    @property
    def display_text(self) -> str:
        """Human-readable reference for reports."""
        if len(self.authors) == 1:
            author_text = self.authors[0]
        elif len(self.authors) == 2:
            author_text = f"{self.authors[0]} y {self.authors[1]}"
        else:
            author_text = f"{self.authors[0]} et al."
        
        return f"{author_text} ({self.year})"
    
    def __repr__(self):
        return f"Reference({self.display_text} @ ¶{self.paragraph_index})"



# ============================================================
# CITATION EXTRACTOR
# ============================================================

class CitationExtractor:
    """Finds APA citations in Spanish text."""
        
    def __init__(self):
        # Pattern 1: Parenthetical citations
        # Now supports:
        #   (García, 2020)
        #   (García et al., 2020)
        #   (García y López, 2020)
        #   (NIST, 2022)
        
        self.pattern_simple = re.compile(
            r'\('                                           # Opening parenthesis
            r'([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+'        # First author (includes hyphens)
            r'(?:\s+et\s+al\.)?'                            # Optional "et al."
            r'(?:\s+y\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+)?)'  # Optional "y Second Author"
            r',\s*'                                         # Comma + spaces
            r'(\d{4}[a-z]?)'                               # Year (2020 or 2020a)
            r'(?:,\s*(?:pp?\.|párr\.)\s*([\d\-]+))?'       # Optional page/paragraph
            r'\)'                                           # Closing parenthesis
        )
        
        # Pattern 2: Narrative citations
        # Now supports:
        #   García (2020)
        #   García et al. (2019)
        #   García y López (2020)
        
        self.pattern_narrative = re.compile(
            r'([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+'        # First author
            r'(?:\s+et\s+al\.)?'                            # Optional "et al."
            r'(?:\s+y\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+)?)'  # Optional second author
            r'\s+\('                                        # Space + opening paren
            r'(\d{4}[a-z]?)'                               # Year
            r'(?:,\s*(?:pp?\.|párr\.)\s*([\d\-]+))?'       # Optional page
            r'\)'                                           # Closing parenthesis
        )

    def extract_simple(self, text: str, para_index: int) -> List[Citation]:
        """Find parenthetical citations like (García, 2020, p. 45)."""
        citations = []
        
        for match in self.pattern_simple.finditer(text):
            authors_raw = match.group(1)
            year = match.group(2)
            page = match.group(3) if match.lastindex >= 3 else None
            
            # Parse authors (handles "y" and "et al.")
            authors = self._parse_authors(authors_raw)
            
            citation = Citation(
                authors=authors,
                year=year,
                paragraph_index=para_index,
                citation_type="parentética",
                raw_text=match.group(0),
                page=page,
                start_pos=match.start()
            )
            citations.append(citation)
        
        return citations
    
    def extract_narrative(self, text: str, para_index: int) -> List[Citation]:
        """Find narrative citations like García (2020)."""
        citations = []
        
        for match in self.pattern_narrative.finditer(text):
            authors_raw = match.group(1)
            year = match.group(2)
            page = match.group(3) if match.lastindex >= 3 else None
            
            # Parse authors
            authors = self._parse_authors(authors_raw)
            
            citation = Citation(
                authors=authors,
                year=year,
                paragraph_index=para_index,
                citation_type="narrativa",
                raw_text=match.group(0),
                page=page,
                start_pos=match.start()
            )
            citations.append(citation)
        
        return citations
  
    def extract_multiple(self, text: str, para_index: int) -> List[Citation]:
        """
        Find multiple citations in one parenthesis.
        Example: (García, 2020; López et al., 2019; Pérez y Martínez, 2018)
        """
        # Pattern to find parentheses containing semicolons
        pattern_multiple = re.compile(r'\(([^)]+;[^)]+)\)')
        
        citations = []
        
        for match in pattern_multiple.finditer(text):
            full_text = match.group(0)  # Full citation with parentheses
            inner_text = match.group(1)  # Text without parentheses
            match_start = match.start()
            
            # Split by semicolon to get individual citations
            individual_cits = inner_text.split(';')
            
            for cit_text in individual_cits:
                cit_text = cit_text.strip()
                
                # Parse each citation (Author, Year [, page])
                # Pattern: Author [et al.] [y Author2], Year [, page]
                cit_pattern = re.compile(
                    r'([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+'    # First author
                    r'(?:\s+et\s+al\.)?'                       # Optional "et al."
                    r'(?:\s+y\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+)?)'  # Optional second author
                    r',\s*'                                    # Comma
                    r'(\d{4}[a-z]?)'                          # Year
                    r'(?:,\s*(?:pp?\.|párr\.)\s*([\d\-]+))?'  # Optional page
                )
                
                cit_match = cit_pattern.match(cit_text)
                
                if cit_match:
                    authors_raw = cit_match.group(1)
                    year = cit_match.group(2)
                    page = cit_match.group(3) if cit_match.lastindex >= 3 else None
                    
                    authors = self._parse_authors(authors_raw)
                    
                    citation = Citation(
                        authors=authors,
                        year=year,
                        paragraph_index=para_index,
                        citation_type="parentética",
                        raw_text=full_text,  # Keep full parenthesis for context
                        page=page,
                        start_pos=match_start
                    )
                    citations.append(citation)
        
        return citations


    def extract_secondary(self, text: str, para_index: int) -> List[Citation]:
        """
        Find secondary citations: (Author, Year, como se cita en SecondAuthor, SecondYear)
        Example: (Saussure, 1916, como se cita en Godel, 1969)
        """
        # Pattern for secondary citations
        pattern_secondary = re.compile(
            r'\('                                           # Opening paren
            r'([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+'        # Primary author
            r'(?:\s+et\s+al\.)?'
            r'(?:\s+y\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+)?)'
            r',\s*'
            r'(\d{4}[a-z]?)'                               # Primary year
            r',\s*como\s+se\s+cita\s+en\s+'               # "como se cita en"
            r'([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+'        # Secondary author
            r'(?:\s+et\s+al\.)?'
            r'(?:\s+y\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+)?)'
            r',\s*'
            r'(\d{4}[a-z]?)'                               # Secondary year
            r'(?:,\s*(?:pp?\.|párr\.)\s*([\d\-]+))?'       # Optional page
            r'\)'                                           # Closing paren
        )
        
        citations = []
        
        for match in pattern_secondary.finditer(text):
            primary_author = match.group(1)
            primary_year = match.group(2)
            secondary_author = match.group(3)
            secondary_year = match.group(4)
            page = match.group(5) if match.lastindex >= 5 else None
            
            authors = self._parse_authors(primary_author)
            
            # Create citation for primary source (the one being cited indirectly)
            citation = Citation(
                authors=authors,
                year=primary_year,
                paragraph_index=para_index,
                citation_type="parentética",
                raw_text=match.group(0),
                page=page,
                start_pos=match.start(),
                is_secondary=True,
                secondary_source=f"{secondary_author}, {secondary_year}"
            )
            citations.append(citation)
        
        return citations

    def extract_all(self, text: str, para_index: int) -> List[Citation]:
        """Find ALL citations in one paragraph."""
        citations = []
        
        # Extract secondary citations FIRST (most specific pattern)
        secondary_cits = self.extract_secondary(text, para_index)
        citations.extend(secondary_cits)
        secondary_positions = {cit.start_pos for cit in secondary_cits}
        
        # Extract multiple citations
        multiple_cits = self.extract_multiple(text, para_index)
        # Avoid positions already captured
        for cit in multiple_cits:
            if cit.start_pos not in secondary_positions:
                citations.append(cit)
        
        multiple_positions = {cit.start_pos for cit in multiple_cits}
        all_positions = secondary_positions | multiple_positions
        
        # Extract simple parenthetical citations
        simple_cits = self.extract_simple(text, para_index)
        for cit in simple_cits:
            if cit.start_pos not in all_positions:
                citations.append(cit)
        
        # Extract narrative citations
        narrative_cits = self.extract_narrative(text, para_index)
        citations.extend(narrative_cits)
        
        return citations
    
       
    @staticmethod
    def _parse_authors(authors_text: str) -> List[str]:
        """
        Parse author string into list.
        Examples:
            "García" → ["García"]
            "García y López" → ["García", "López"]
            "García et al." → ["García et al."]
            "NIST" → ["NIST"]
        """
        # Handle et al. case
        if "et al." in authors_text:
            first_author = authors_text.split("et al.")[0].strip()
            return [f"{first_author} et al."]
        
        # Handle "y" separator for two authors
        if " y " in authors_text:
            return [a.strip() for a in authors_text.split(" y ")]
        
        # Single author (could be person or institution)
        return [authors_text.strip()]



class ReferenceExtractor:
    """Extracts APA references from bibliography section."""
    
    def __init__(self):
        # Pattern for APA reference: Author, I. (Year). Title...
        # Examples:
        #   García, A. (2020). Title...
        #   García, A. & López, B. (2019). Title...
        #   IBM Research. (2024). Title...
        
        self.pattern_reference = re.compile(
            r'^([A-ZÁÉÍÓÚÑ][^\(]+?)'     # Authors (up to year parenthesis)
            r'\s*\((\d{4}[a-z]?)\)\.'    # Year in parentheses with dot
            r'\s*(.+?)(?:\.|$)',          # Title (up to first period or end)
            re.MULTILINE | re.IGNORECASE
        )
    
    def detect_section_type(self, paragraphs: List[str]) -> tuple:
        """
        Detect if the list is 'Referencias' or 'Bibliografía'.
        
        Returns:
            (section_type, paragraph_index) where section_type is:
            - "referencias" = strict APA (must cite everything)
            - "bibliografia" = consulted sources (citation optional)
            - "unknown" = couldn't determine
        """
        # Look for section headers in last 20 paragraphs
        search_start = max(0, len(paragraphs) - 20)
        
        for i in range(search_start, len(paragraphs)):
            para_lower = paragraphs[i].lower().strip()
            
            # Check for "Referencias" variations
            if para_lower in ['referencias', 'referencias bibliográficas', 'references']:
                return ("referencias", i)
            
            # Check for "Bibliografía" variations
            if any(keyword in para_lower for keyword in [
                'bibliografía',
                'bibliografia',
                'fuentes bibliográficas',
                'fuentes consultadas',
                'bibliography'
            ]):
                return ("bibliografia", i)
        
        return ("unknown", -1)
    
    def extract_from_paragraphs(self, paragraphs: List[str], start_index: int = 0) -> List[Reference]:
        """
        Extract references from paragraphs (usually the last section).
        
        Args:
            paragraphs: List of paragraph texts
            start_index: Which paragraph to start from (references are usually at end)
        
        Returns:
            List of Reference objects
        """
        references = []
        
        for i in range(start_index, len(paragraphs)):
            para_text = paragraphs[i].strip()
            
            # Skip empty paragraphs
            if not para_text:
                continue
            
            # Try to match APA reference pattern
            match = self.pattern_reference.match(para_text)
            
            if match:
                authors_raw = match.group(1).strip()
                year = match.group(2)
                title = match.group(3).strip()
                
                # Parse authors
                authors = self._parse_authors(authors_raw)
                
                reference = Reference(
                    authors=authors,
                    year=year,
                    title=title,
                    raw_text=para_text,
                    paragraph_index=i
                )
                references.append(reference)
        
        return references
    
    @staticmethod
    def _parse_authors(authors_text: str) -> List[str]:
        """
        Parse author string into list.
        Examples:
            "García, A." → ["García, A."]
            "García, A. & López, B." → ["García, A.", "López, B."]
            "García, A., López, B., & Pérez, C." → ["García, A.", "López, B.", "Pérez, C."]
        """
        # Replace " & " with ", " for uniform splitting
        authors_text = authors_text.replace(' & ', ', ')
        
        # Split by comma, but keep "Apellido, Inicial" together
        # This is a simplified approach - we'll improve it later
        parts = authors_text.split(', ')
        
        authors = []
        i = 0
        while i < len(parts):
            # Each author is "Apellido, Initial"
            if i + 1 < len(parts) and len(parts[i+1]) <= 3:  # Likely an initial
                authors.append(f"{parts[i]}, {parts[i+1]}")
                i += 2
            else:
                # Institutional author or last name only
                authors.append(parts[i])
                i += 1
        
        return authors


class CitationMatcher:
    """Matches in-text citations with reference list entries."""
    
    def __init__(self, citations: List[Citation], references: List[Reference]):
        self.citations = citations
        self.references = references
        
        # Build lookup dictionaries for fast matching
        self.citation_keys = {cit.citation_key for cit in citations}
        self.reference_keys = {ref.reference_key for ref in references}
        
        # Build reference lookup by key
        self.ref_lookup = {ref.reference_key: ref for ref in references}
    
    def find_orphaned_citations(self) -> List[Citation]:
        """
        Find citations that don't have matching references.
        These are CRITICAL errors - cited but not in bibliography.
        """
        orphaned = []
        
        for citation in self.citations:
            if citation.citation_key not in self.reference_keys:
                orphaned.append(citation)
        
        return orphaned
    
    def find_orphaned_references(self) -> List[Reference]:
        """
        Find references that are never cited in text.
        These are WARNING level - unnecessary references.
        """
        orphaned = []
        
        for reference in self.references:
            if reference.reference_key not in self.citation_keys:
                orphaned.append(reference)
        
        return orphaned
    
    def find_year_discrepancies(self) -> List[tuple]:
        """
        Find cases where citation year doesn't match reference year.
        Example: Text says (García, 2020) but reference says (2019).
        """
        discrepancies = []
        
        for citation in self.citations:
            # Find matching reference
            ref = self.ref_lookup.get(citation.citation_key)
            
            if ref and citation.year != ref.year:
                discrepancies.append((citation, ref))
        
        return discrepancies
    
    def generate_report(self, section_type: str = "referencias") -> str:
        """
        Generate comprehensive citation integrity report.
        
        Args:
            section_type: "referencias" or "bibliografia" - affects severity
        """
        report = []
        report.append("=" * 60)
        report.append("SILVINA v0.6 - Reporte de Integridad de Citas")
        report.append("=" * 60)
        report.append("")
        
        # Summary statistics
        report.append("📊 Estadísticas:")
        report.append(f"  • Citas en texto: {len(self.citations)}")
        report.append(f"  • Entradas bibliográficas: {len(self.references)}")
        report.append(f"  • Tipo de sección: {section_type.upper()}")
        report.append("")
        
        # Orphaned citations (ALWAYS CRITICAL)
        orphaned_cits = self.find_orphaned_citations()
        if orphaned_cits:
            report.append("🔴 CRÍTICO: Citas sin Entrada Bibliográfica")
            report.append(f"  Encontradas {len(orphaned_cits)} citas que NO aparecen en la lista:")
            report.append("")
            for cit in orphaned_cits[:10]:
                report.append(f"  • {cit.display_text} [Párrafo {cit.paragraph_index}]")
                report.append(f"    └─ Texto: {cit.raw_text}")
            if len(orphaned_cits) > 10:
                report.append(f"  ... y {len(orphaned_cits) - 10} más")
            report.append("")
        
        # Orphaned references (severity depends on section type)
        orphaned_refs = self.find_orphaned_references()
        if orphaned_refs:
            if section_type == "referencias":
                # Strict APA - this is a WARNING (should cite everything)
                report.append("🟡 ADVERTENCIA: Referencias sin Citar en Texto")
                report.append(f"  En secciones 'Referencias', se espera citar todas las entradas.")
                report.append(f"  Encontradas {len(orphaned_refs)} referencias no citadas:")
            else:
                # Bibliography - this is just INFORMATIONAL
                report.append("🔵 INFORMATIVO: Entradas Bibliográficas sin Citar")
                report.append(f"  En 'Bibliografía', es aceptable incluir fuentes consultadas.")
                report.append(f"  Encontradas {len(orphaned_refs)} entradas no citadas:")
            
            report.append("")
            for ref in orphaned_refs[:10]:
                report.append(f"  • {ref.display_text} [Párrafo {ref.paragraph_index}]")
                report.append(f"    └─ {ref.title[:60]}...")
            if len(orphaned_refs) > 10:
                report.append(f"  ... y {len(orphaned_refs) - 10} más")
            report.append("")
        
        # Year discrepancies (ALWAYS CRITICAL)
        discrepancies = self.find_year_discrepancies()
        if discrepancies:
            report.append("🔴 CRÍTICO: Discrepancias de Año")
            report.append(f"  Encontradas {len(discrepancies)} inconsistencias:")
            report.append("")
            for cit, ref in discrepancies:
                report.append(f"  • Cita: {cit.display_text} vs Referencia: {ref.display_text}")
                report.append(f"    └─ [Párrafo {cit.paragraph_index}] → [Párrafo {ref.paragraph_index}]")
            report.append("")
        
        # Final verdict - FIXED LOGIC
        if not orphaned_cits and not discrepancies:
            if section_type == "referencias":
                if not orphaned_refs:
                    report.append("✅ PERFECTO: Sistema de Citación Íntegro")
                    report.append("  • Todas las citas tienen referencias")
                    report.append("  • Todas las referencias son citadas")
                    report.append("  • No hay discrepancias de año")
                else:
                    report.append("🟡 ADVERTENCIA: Referencias sin Citar")
                    report.append("  • Algunas referencias no son citadas en el texto")
            elif section_type == "bibliografia":
                if len(self.citations) > 0:
                    # Has citations + bibliography = good
                    report.append("✅ ACEPTABLE: Sistema de Citación Válido")
                    report.append("  • Todas las citas tienen entrada bibliográfica")
                    report.append("  • Bibliografía puede incluir fuentes consultadas")
                    report.append("  • No hay discrepancias de año")
                else:
                    # NO citations but HAS bibliography = CRITICAL PROBLEM
                    report.append("🔴 CRÍTICO: Documento Sin Sistema de Citación Formal")
                    report.append("")
                    report.append("  El documento tiene bibliografía pero NO tiene citas en formato APA.")
                    report.append("")
                    report.append("  📋 Problemas detectados:")
                    report.append("     • Ninguna cita formal en el texto (0 encontradas)")
                    report.append("     • Referencias bibliográficas presentes pero no utilizadas")
                    report.append("     • Posibles afirmaciones sin respaldo formal")
                    report.append("")
                    report.append("  ⚠️  Esto significa que:")
                    report.append("     • El lector no puede verificar las fuentes de las afirmaciones")
                    report.append("     • No se puede evaluar la solidez del argumento")
                    report.append("     • Viola estándares académicos básicos")
                    report.append("")
                    report.append("  ✅ Solución requerida:")
                    report.append("     • Agregar citas en formato APA en el texto:")
                    report.append("       Ejemplo parentético: (IBM Research, 2024)")
                    report.append("       Ejemplo narrativo: Según Gidney y Ekera (2024), ...")
                    report.append("     • Cada afirmación debe estar respaldada por una cita")
                    report.append("     • Relacionar las citas con las entradas bibliográficas")
            report.append("")
        
        return "\n".join(report)

     
    
# ============================================================
# WORD DOCUMENT READER
# ============================================================

class WordDocumentReader:
    """Reads paragraphs from Word documents using pywin32."""
    
    def __init__(self, file_path: str):
        if not HAS_WIN32:
            raise ImportError("pywin32 no está instalado. Instalar con: pip install pywin32")
        
        self.file_path = Path(file_path)
        self.word = None
        self.doc = None
    
    def open(self):
        """Open Word application and document."""
        try:
            self.word = win32.Dispatch("Word.Application")
            self.word.Visible = False
            self.doc = self.word.Documents.Open(str(self.file_path.absolute()))
            print(f"✓ Documento abierto: {self.file_path.name}")
            return True
        except Exception as e:
            print(f"✗ Error abriendo documento: {e}")
            return False
    
    def get_paragraphs(self) -> List[str]:
        """Extract all paragraph texts from document."""
        if not self.doc:
            return []
        
        paragraphs = []
        for para in self.doc.Paragraphs:
            text = para.Range.Text.strip()
            if text and not para.Style.NameLocal.startswith("Título"):
                paragraphs.append(text)
        
        print(f"✓ Extraídos {len(paragraphs)} párrafos")
        return paragraphs
    
    def close(self):
        """Close document and Word application."""
        if self.doc:
            self.doc.Close(SaveChanges=False)
        if self.word:
            self.word.Quit()
        print("✓ Documento cerrado")
    
    def __enter__(self):
        self.open()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


# ============================================================
# MAIN ANALYSIS FUNCTION
# ============================================================

def analyze_document_citations(docx_path: str):
    """
    Extract and analyze all citations from a Word document.
    
    Args:
        docx_path: Path to .docx file
    
    Returns:
        List of Citation objects found
    """
    print("\n" + "="*60)
    print("SILVINA v0.6 - Análisis de Citas")
    print("="*60)
    
    # Read document
    with WordDocumentReader(docx_path) as reader:
        paragraphs = reader.get_paragraphs()
    
    if not paragraphs:
        print("✗ No se encontraron párrafos")
        return []
    
    # Extract citations
    print("\n📊 Extrayendo citas...")
    extractor = CitationExtractor()
    
    all_citations = []
    for i, para_text in enumerate(paragraphs):
        citations = extractor.extract_all(para_text, para_index=i)
        all_citations.extend(citations)
    
    # Report results
    print(f"\n✓ Análisis completado")
    print(f"  • Total citas: {len(all_citations)}")
    print(f"  • Parentéticas: {sum(1 for c in all_citations if c.citation_type == 'parentética')}")
    print(f"  • Narrativas: {sum(1 for c in all_citations if c.citation_type == 'narrativa')}")
    
    # Show first 10 citations as sample
    if all_citations:
        print(f"\n📋 Primeras {min(10, len(all_citations))} citas encontradas:")
        for cit in all_citations[:10]:
            print(f"  {cit}")
    
    return all_citations

def debug_document_paragraphs(docx_path: str, max_paragraphs: int = 20):
    """
    Show first N paragraphs to debug citation detection.
    
    Args:
        docx_path: Path to .docx file
        max_paragraphs: Number of paragraphs to display
    """
    print("\n" + "="*60)
    print("SILVINA v0.6 - Modo Debug: Visualización de Párrafos")
    print("="*60)
    
    with WordDocumentReader(docx_path) as reader:
        paragraphs = reader.get_paragraphs()
    
    if not paragraphs:
        print("✗ No se encontraron párrafos")
        return
    
    print(f"\n📝 Mostrando los primeros {min(max_paragraphs, len(paragraphs))} párrafos:\n")
    
    for i, para in enumerate(paragraphs[:max_paragraphs]):
        print(f"--- Párrafo {i} ({len(para)} caracteres) ---")
        print(para)
        print()

def search_parentheses(docx_path: str):
    """Find all paragraphs containing parentheses (potential citations)."""
    print("\n" + "="*60)
    print("SILVINA v0.6 - Búsqueda de Paréntesis")
    print("="*60)
    
    with WordDocumentReader(docx_path) as reader:
        paragraphs = reader.get_paragraphs()
    
    print(f"\n🔍 Buscando párrafos con paréntesis...\n")
    
    found_count = 0
    for i, para in enumerate(paragraphs):
        if '(' in para and ')' in para:
            found_count += 1
            print(f"--- Párrafo {i} ---")
            # Extract content between parentheses
            import re
            matches = re.findall(r'\([^)]+\)', para)
            if matches:
                print(f"  Paréntesis encontrados: {len(matches)}")
                for match in matches[:3]:  # Show first 3
                    print(f"    • {match}")
            print(f"  Texto: {para[:200]}...")
            print()
    
    print(f"✓ Total: {found_count} párrafos con paréntesis de {len(paragraphs)} totales")

def check_citation_integrity(docx_path: str):
    """
    Check if document has orphaned references (references without in-text citations).
    This is a critical editorial problem.
    """
    print("\n" + "="*60)
    print("SILVINA v0.6 - Verificación de Integridad de Citas")
    print("="*60)
    
    with WordDocumentReader(docx_path) as reader:
        paragraphs = reader.get_paragraphs()
    
    # Extract citations
    extractor = CitationExtractor()
    all_citations = []
    for i, para_text in enumerate(paragraphs):
        citations = extractor.extract_all(para_text, para_index=i)
        all_citations.extend(citations)
    
    # Detect reference section (paragraphs with author names and years)
    reference_pattern = re.compile(r'^[A-Z][a-zA-Z]+,\s+[A-Z]')  # "Author, A."
    reference_paragraphs = []
    
    for i, para in enumerate(paragraphs):
        if reference_pattern.match(para.strip()):
            reference_paragraphs.append((i, para[:100]))
    
    # Generate report
    print(f"\n📊 Resultados del Análisis:\n")
    print(f"  • Total de párrafos: {len(paragraphs)}")
    print(f"  • Citas en texto encontradas: {len(all_citations)}")
    print(f"  • Referencias bibliográficas: {len(reference_paragraphs)}")
    
    # Critical issue detection
    if len(reference_paragraphs) > 0 and len(all_citations) == 0:
        print(f"\n🔴 CRÍTICO: Problema de Integridad de Citas Detectado")
        print(f"\n  El documento tiene {len(reference_paragraphs)} referencias bibliográficas")
        print(f"  pero NO tiene citas en el texto.")
        print(f"\n  📋 Esto significa que:")
        print(f"     • Las referencias nunca son citadas en el cuerpo del artículo")
        print(f"     • No se puede verificar qué afirmaciones están respaldadas")
        print(f"     • Viola normas APA y estándares académicos")
        
        print(f"\n  ⚠️  Referencias encontradas (primeras 5):")
        for i, (para_idx, ref_text) in enumerate(reference_paragraphs[:5]):
            print(f"     {i+1}. [Párrafo {para_idx}] {ref_text}...")
        
        print(f"\n  ✅ Solución requerida:")
        print(f"     • Agregar citas en formato APA en el texto:")
        print(f"       Ejemplo: (Gidney & Ekera, 2024)")
        print(f"       Ejemplo: Según IBM Research (2024), ...")
    
    elif len(all_citations) > 0 and len(reference_paragraphs) == 0:
        print(f"\n🔴 CRÍTICO: Citas sin Lista de Referencias")
        print(f"  El documento cita {len(all_citations)} fuentes pero no tiene")
        print(f"  una sección de Referencias bibliográficas.")
    
    elif len(all_citations) == 0 and len(reference_paragraphs) == 0:
        print(f"\n🟡 ADVERTENCIA: Sin Sistema de Citación")
        print(f"  El documento no tiene citas ni referencias.")
        print(f"  Si es un artículo académico, esto debe corregirse.")
    
    else:
        print(f"\n✅ Sistema de citación presente")
        print(f"  • {len(all_citations)} citas en texto")
        print(f"  • {len(reference_paragraphs)} referencias bibliográficas")


def test_reference_extraction(docx_path: str):
    """Test reference extraction from a document."""
    print("\n" + "="*60)
    print("SILVINA v0.6 - Test: Extracción de Referencias")
    print("="*60)
    
    with WordDocumentReader(docx_path) as reader:
        paragraphs = reader.get_paragraphs()
    
    print(f"\n📚 Buscando referencias en los últimos 15 párrafos...\n")
    
    # Extract references from last 15 paragraphs (where references usually are)
    extractor = ReferenceExtractor()
    start_para = max(0, len(paragraphs) - 15)
    references = extractor.extract_from_paragraphs(paragraphs, start_index=start_para)
    
    print(f"✓ Referencias encontradas: {len(references)}\n")
    
    for ref in references:
        print(f"  {ref}")
        print(f"    └─ Key: {ref.reference_key}")
        print(f"    └─ Autores: {ref.authors}")
        print(f"    └─ Título: {ref.title[:60]}...")
        print()
    
    return references

def analyze_citation_reference_matching(docx_path: str):
    """
    Complete citation-reference integrity analysis.
    Detects section type (Referencias vs Bibliografía) and adjusts validation.
    """
    print("\n" + "="*60)
    print("SILVINA v0.6 - Análisis Completo de Citas y Referencias")
    print("="*60)
    
    # Read document
    with WordDocumentReader(docx_path) as reader:
        paragraphs = reader.get_paragraphs()
    
    print(f"\n📖 Extrayendo citas del texto...")
    
    # Extract citations
    cit_extractor = CitationExtractor()
    all_citations = []
    for i, para_text in enumerate(paragraphs):
        citations = cit_extractor.extract_all(para_text, para_index=i)
        all_citations.extend(citations)
    
    print(f"  ✓ {len(all_citations)} citas encontradas")
    
    print(f"\n📚 Detectando tipo de sección bibliográfica...")
    
    # Detect section type
    ref_extractor = ReferenceExtractor()
    section_type, section_para = ref_extractor.detect_section_type(paragraphs)
    
    if section_type == "referencias":
        print(f"  ✓ Sección detectada: REFERENCIAS (Párrafo {section_para})")
        print(f"    └─ Norma APA: Todas deben ser citadas")
    elif section_type == "bibliografia":
        print(f"  ✓ Sección detectada: BIBLIOGRAFÍA (Párrafo {section_para})")
        print(f"    └─ Puede incluir fuentes consultadas sin citar")
    else:
        print(f"  ⚠️  Sección no identificada - asumiendo Referencias")
        section_type = "referencias"  # Default to strict
    
    print(f"\n📚 Extrayendo entradas bibliográficas...")
    
    # Extract references (from last 20 paragraphs)
    start_para = max(0, len(paragraphs) - 20)
    all_references = ref_extractor.extract_from_paragraphs(paragraphs, start_index=start_para)
    
    print(f"  ✓ {len(all_references)} entradas encontradas")
    
    print(f"\n🔍 Verificando integridad...")
    
    # Match citations with references
    matcher = CitationMatcher(all_citations, all_references)
    
    # Generate report with section type awareness
    report = matcher.generate_report(section_type=section_type)
    print("\n" + report)
    
    return matcher


# ============================================================
# MAIN ENTRY POINT
# ============================================================

if __name__ == "__main__":
    # Test mode (no arguments)
    if len(sys.argv) == 1:
        print("SILVINA v0.6 - Citation Extractor (Test Mode)")
        print("="*50)
        
        
       # NEW TEST DATA - Session 3 COMPLETE
        test_paragraphs = [
            "El cambio climático es real (García, 2020, p. 45).",
            "Según López et al. (2019) el problema es grave.",
            "Dos autores (Pérez y Martínez, 2021) confirman esto.",
            "Institución (NIST, 2022) publicó estándares.",
            "Apellido compuesto (García-López, 2018) analizó datos.",
            "Narrativa con dos: Sánchez y Rodríguez (2023) proponen un modelo.",
            "Múltiples citas (García, 2020; López et al., 2019; Pérez, 2021).",
            "Con páginas (Martínez, 2022, p. 10; Ruiz y Soto, 2021, pp. 5-8).",
            "Cita secundaria (Saussure, 1916, como se cita en Godel, 1969).",
        ]
                           
        extractor = CitationExtractor()
        all_citations = []
        
        for i, paragraph in enumerate(test_paragraphs):
            found = extractor.extract_all(paragraph, para_index=i)
            all_citations.extend(found)
            if found:
                print(f"\nPárrafo {i}: {paragraph}")
                for cit in found:
                    print(f"  → {cit}")
        
        
        print("\n💡 Comandos disponibles:")
        print("   python silvina_editorial_v0.6.py documento.docx            # Analizar")
        print("   python silvina_editorial_v0.6.py documento.docx --debug    # Ver párrafos")
        print("   python silvina_editorial_v0.6.py documento.docx --search   # Buscar paréntesis")
        print("   python silvina_editorial_v0.6.py documento.docx --check    # Verificar integridad")
        print("   python silvina_editorial_v0.6.py documento.docx --refs     # Extraer referencias")
        print("   python silvina_editorial_v0.6.py documento.docx --match     # Análisis completo")

    # Check for flags BEFORE default analysis
    elif len(sys.argv) >= 2:
        docx_file = sys.argv[1]
        
        if not Path(docx_file).exists():
            print(f"✗ Error: Archivo no encontrado: {docx_file}")
            sys.exit(1)
        
        # Now check which mode (all at same indentation level)
        # Check integrity mode
        if len(sys.argv) == 3 and sys.argv[2] == "--check":
            try:
                check_citation_integrity(docx_file)
            except ImportError as e:
                print(f"✗ Error: {e}")
                sys.exit(1)
                       

        # Search mode (find parentheses)
        elif len(sys.argv) == 3 and sys.argv[2] == "--search":
            try:
                search_parentheses(docx_file)
            except ImportError as e:
                print(f"✗ Error: {e}")
                sys.exit(1)
        
        # Debug mode (show paragraphs)
        elif len(sys.argv) >= 3 and sys.argv[2] == "--debug":
            start_para = int(sys.argv[3]) if len(sys.argv) > 3 else 15
            try:
                debug_document_paragraphs(docx_file, start=start_para, count=25)
            except ImportError as e:
                print(f"✗ Error: {e}")
                sys.exit(1)
        
        # Test reference extraction
        elif len(sys.argv) == 3 and sys.argv[2] == "--refs":
            try:
                test_reference_extraction(docx_file)
            except ImportError as e:
                print(f"✗ Error: {e}")
                sys.exit(1)

        # Full citation-reference matching analysis
        elif len(sys.argv) == 3 and sys.argv[2] == "--match":
            try:
                analyze_citation_reference_matching(docx_file)
            except ImportError as e:
                print(f"✗ Error: {e}")
                sys.exit(1)

        # Default: Document analysis mode (no flag)
        else:
            try:
                citations = analyze_document_citations(docx_file)
            except ImportError as e:
                print(f"✗ Error: {e}")
                sys.exit(1)
