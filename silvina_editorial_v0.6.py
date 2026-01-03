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
    
    def __repr__(self):
        """Show citation in readable format."""
        authors_text = " y ".join(self.authors)
        page_text = f", p. {self.page}" if self.page else ""
        type_marker = "📖" if self.citation_type == "narrativa" else "📎"
        return f"{type_marker} {authors_text} ({self.year}{page_text}) [¶{self.paragraph_index}]"


# ============================================================
# CITATION EXTRACTOR
# ============================================================

class CitationExtractor:
    """Finds APA citations in Spanish text."""
    
    def __init__(self):
        # Pattern 1: Parenthetical (García, 2020, p. 45)
        self.pattern_simple = re.compile(
            r'\('
            r'([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+et\s+al\.)?)'
            r',\s*'
            r'(\d{4}[a-z]?)'
            r'(?:,\s*(?:pp?\.|párr\.)\s*([\d\-]+))?'
            r'\)'
        )
        
        # Pattern 2: Narrative - García (2020)
        self.pattern_narrative = re.compile(
            r'([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+et\s+al\.)?)'
            r'\s+\('
            r'(\d{4}[a-z]?)'
            r'(?:,\s*(?:pp?\.|párr\.)\s*([\d\-]+))?'
            r'\)'
        )
    
    def extract_simple(self, text: str, para_index: int) -> List[Citation]:
        """Find parenthetical citations like (García, 2020, p. 45)."""
        citations = []
        
        for match in self.pattern_simple.finditer(text):
            author = match.group(1)
            year = match.group(2)
            page = match.group(3) if match.lastindex >= 3 else None
            
            citation = Citation(
                authors=[author],
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
            author = match.group(1)
            year = match.group(2)
            page = match.group(3) if match.lastindex >= 3 else None
            
            citation = Citation(
                authors=[author],
                year=year,
                paragraph_index=para_index,
                citation_type="narrativa",
                raw_text=match.group(0),
                page=page,
                start_pos=match.start()
            )
            citations.append(citation)
        
        return citations
    
    def extract_all(self, text: str, para_index: int) -> List[Citation]:
        """Find ALL citations (parenthetical + narrative) in one paragraph."""
        citations = []
        citations.extend(self.extract_simple(text, para_index))
        citations.extend(self.extract_narrative(text, para_index))
        return citations


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




# ============================================================
# MAIN ENTRY POINT
# ============================================================

if __name__ == "__main__":
    # Test mode (no arguments)
    if len(sys.argv) == 1:
        print("SILVINA v0.6 - Citation Extractor (Test Mode)")
        print("="*50)
        
        test_paragraphs = [
            "El cambio climático es real (García, 2020, p. 45).",
            "Según López et al. (2019) el problema es grave.",
            "Varios estudios (Pérez, 2021a) y Martínez (2018) lo confirman.",
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
        
        # Default: Document analysis mode (no flag)
        else:
            try:
                citations = analyze_document_citations(docx_file)
            except ImportError as e:
                print(f"✗ Error: {e}")
                sys.exit(1)
