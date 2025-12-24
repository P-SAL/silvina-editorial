"""
Silvina Editorial Assistant v0.4
Object-Oriented Refactor with Classes

NEW in v0.4:
- Reference class (encapsulates one citation)
- Document class (manages Word document)
- Cleaner, more maintainable code structure

Author: Pablo Salonio
Repository: https://github.com/P-SAL/silvina-editorial
"""

from datetime import datetime
import re
import win32com.client
import pythoncom
import time
import os


# === DOCUMENT CLASS ===
class Document:
    """Manages Word document loading and reference extraction."""
    
    def __init__(self, filepath):
        """Initialize with filepath only."""
        self.filepath = filepath
        self.word = None
        self.doc = None
        self.text = ""
        self.references = []
    
    def load(self):
        """Load document and extract references."""
        self._connect_to_word()
        self._extract_referencias()
        self._create_reference_objects()
    
    def _connect_to_word(self):
        """Open Word document."""
        pythoncom.CoInitialize()
        
        try:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = True
            abs_path = os.path.abspath(self.filepath)
            self.doc = self.word.Documents.Open(abs_path)
            self.word.Visible = False
            time.sleep(1.0)
            print(f"✅ Connected: {abs_path}")
        except Exception as e:
            print(f"❌ Error: {e}")
            self.word = None
            self.doc = None
    
    def _extract_referencias(self):
        """Extract Referencias section using paragraphs (no truncation)."""
        
        if not self.doc:
            return
        
        try:
            time.sleep(1.0)
            char_count = self.get_character_count()
            print(f"🔍 Characters: {char_count:,}")
            
            # Find the paragraph with referencias heading
            found_start = False
            referencias_paras = []
            
            for para in self.doc.Paragraphs:
                para_text = para.Range.Text.strip()
                
                # Check if this is the heading
                if not found_start:
                    if "Fuentes bibliográficas" in para_text or "Referencias" in para_text or "Bibliografía" in para_text:
                        found_start = True
                        continue  # Skip heading itself
                
                # After heading, collect all remaining paragraphs
                if found_start and para_text:
                    referencias_paras.append(para_text)
            
            # Join paragraphs with newlines
            self.text = '\n'.join(referencias_paras)
                            
        except Exception as e:
            print(f"❌ Extract error: {e}")
            self.text = ""
    
    def _create_reference_objects(self):
        """Create Reference objects from extracted paragraphs."""
        if not self.text:
            return
        
        # Split by newlines - each paragraph is a reference
        paragraphs = self.text.split('\n')
        
        for para in paragraphs:
            para = para.strip()
            if len(para) < 30:
                continue
            
            # Special case: check if paragraph has TWO years (two merged refs)
            years = re.findall(r'\(\d{4}\)', para)
            
            if len(years) >= 2:
                # Split at period before capital letter pattern
                split_pattern = r'\.(?=[A-Z][a-z]+,\s+[A-Z]\.)'
                parts = re.split(split_pattern, para, maxsplit=1)
                
                for part in parts:
                    part = part.strip()
                    if len(part) > 30:
                        if not part.endswith('.'):
                            part += '.'
                        self.references.append(Reference(part))
            else:
                # Single reference - add as is
                self.references.append(Reference(para))
        
        print(f"✅ Created {len(self.references)} Reference objects")
    
    def get_character_count(self):
        """Get accurate Word character count."""
        if not self.doc:
            return 0
        
        try:
            total = self.doc.Characters.Count
            for fn in self.doc.Footnotes:
                total += len(fn.Range.Text)
            for en in self.doc.Endnotes:
                total += len(en.Range.Text)
            return total
        except:
            return 0
    
    def close(self):
        """Clean up Word connection."""
        try:
            if self.doc:
                self.doc.Close(SaveChanges=False)
            if self.word:
                self.word.Quit()
        except:
            pass


    def generate_report(self):
        """Generate formatted validation report."""
        if not self.references:
            return "No references found."
        
        report = []
        report.append("=" * 70)
        report.append("SILVINA - VALIDACIÓN DE REFERENCIAS APA")
        report.append("=" * 70)
        report.append(f"\nDocumento: {os.path.basename(self.filepath)}")
        report.append(f"Caracteres totales: {self.get_character_count():,}")
        report.append(f"Referencias encontradas: {len(self.references)}")
        
        # Count valid/invalid
        valid_count = sum(1 for ref in self.references if ref.is_valid())
        invalid_count = len(self.references) - valid_count
        
        report.append(f"\n✅ Válidas: {valid_count}")
        report.append(f"❌ Con problemas: {invalid_count}")
        
        report.append("\n" + "-" * 70)
        report.append("DETALLE DE VALIDACIÓN")
        report.append("-" * 70 + "\n")
        
        for i, ref in enumerate(self.references, 1):
            rep = ref.get_validation_report()
            status = "✅ VÁLIDA" if rep['is_valid'] else "❌ REQUIERE REVISIÓN"
            
            report.append(f"{i}. {status}")
            report.append(f"   Texto: {rep['text']}")
            
            if not rep['valid_author']:
                report.append("   ⚠️ Formato de autor incorrecto (debe ser: Apellido, I.)")
            if not rep['valid_year']:
                report.append("   ⚠️ Año no encontrado o formato incorrecto (debe ser: (YYYY))")
            
            report.append("")
        
        report.append("=" * 70)
        report.append(f"Reporte generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        report.append("=" * 70)
        
        return '\n'.join(report)


# === REFERENCE CLASS ===
class Reference:
    """Represents a single bibliographic reference"""
    
    def __init__(self, text):
        """Initialize reference with citation text"""
        self.text = text
        
    def validate_author(self):
        """Check if reference has valid APA author format"""
        pattern = r'[A-ZÁ-ÚÑ][a-zá-úñ]+(?:-[A-ZÁ-ÚÑ][a-zá-úñ]+)?,\s[A-Z]\.'
        return bool(re.search(pattern, self.text))
    
    def validate_year(self):
        """Check if reference has valid year format (YYYY)"""
        pattern = r'\((\d{4})\)'
        match = re.search(pattern, self.text)
        if match:
            return True, match.group(1)
        return False, None
    
    def is_valid(self):
        """Check if reference meets all APA requirements"""
        has_author = self.validate_author()
        has_year, _ = self.validate_year()
        return has_author and has_year
    
    def get_validation_report(self):
        """Return detailed validation results"""
        has_author = self.validate_author()
        has_year, year = self.validate_year()
        
        return {
            'text': self.text[:80] + '...' if len(self.text) > 80 else self.text,
            'valid_author': has_author,
            'valid_year': has_year,
            'year': year,
            'is_valid': has_author and has_year
        }


# === MAIN EXECUTION ===
if __name__ == "__main__":
    print("\n" + "="*70)
    print("SILVINA v0.4 - ASISTENTE EDITORIAL")
    print("="*70 + "\n")
    
    # UPDATE THIS PATH
    filepath = r"C:\Users\usuario\Desktop\Escudo cuantico_AB_25092025.docx"
    
    doc = Document(filepath)
    doc.load()
    
    # Generate and display report
    report = doc.generate_report()
    print(report)

    # Save report to file
    report_filename = f"reporte_silvina_v04_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(report_filename, 'w', encoding='utf-8') as f:
        f.write(report)

    print(f"\n💾 Reporte guardado: {report_filename}")
    doc.close()

    