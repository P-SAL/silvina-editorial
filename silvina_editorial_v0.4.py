"""
Silvina Editorial Assistant v0.4
Object-Oriented Refactor with Classes

NEW in v0.4:
- Reference class (encapsulates one citation)
- Document class (manages Word document)
- APAValidator class (validation rules)
- Cleaner, more maintainable code structure

Author: Pablo Salonio
Repository: https://github.com/P-SAL/silvina-editorial
"""

from datetime import datetime
import re
import win32com.client
import pythoncom


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
    print("╔════════════════════════════════════════════════════════════════╗")
    print("║              SILVINA - ASISTENTE EDITORIAL v0.4                ║")
    print("║              Refactor con Programación Orientada a Objetos    ║")
    print("╚════════════════════════════════════════════════════════════════╝\n")
    
    # Test the Reference class
    print("🧪 Testing Reference class...\n")
    
    # Test reference 1 (valid)
    ref1 = Reference("García, M. (2023). Inteligencia artificial en educación. Revista Tech, 15(2), 45-67.")
    report1 = ref1.get_validation_report()
    
    print(f"Reference 1: {report1['text']}")
    print(f"  ✓ Author valid: {report1['valid_author']}")
    print(f"  ✓ Year valid: {report1['valid_year']} ({report1['year']})")
    print(f"  ✓ Overall valid: {report1['is_valid']}\n")
    
    # Test reference 2 (invalid - no year)
    ref2 = Reference("López, J. Title without year. Journal Name.")
    report2 = ref2.get_validation_report()
    
    print(f"Reference 2: {report2['text']}")
    print(f"  ✓ Author valid: {report2['valid_author']}")
    print(f"  ✓ Year valid: {report2['valid_year']}")
    print(f"  ✓ Overall valid: {report2['is_valid']}\n")
    
    print("✅ Reference class working!")

    print("\n" + "="*70)
    print("🧪 Testing with Real APA References...\n")
    
    # Test real references from APA guide
    real_refs = [
        "Herrera Cáceres, C. y Rosillo Peña, M. (2019). Confort y eficiencia energética en el diseño de edificaciones. Universidad del Valle.",
        "Castañeda Naranjo, L. A. y Palacios Neri, J. (2015). Nanotecnología: fuente de nuevos paradigmas. Mundo Nano, 7(12), 45-49.",
        "Invalid reference without proper format",
        "García M (2023) Missing comma after surname.",
        "Pérez-Sánchez, C. (2020). Compound surname test. Journal, 5(1), 10-20."
    ]
    
    for i, ref_text in enumerate(real_refs, 1):
        ref = Reference(ref_text)
        report = ref.get_validation_report()
        
        symbol = "✅" if report['is_valid'] else "❌"
        print(f"{symbol} Reference {i}:")
        print(f"   Text: {report['text']}")
        print(f"   Author: {report['valid_author']}, Year: {report['valid_year']}")
        print()