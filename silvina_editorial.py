# silvina_editorial.py - Editorial Assistant Agent v0.1
# Pablo Salonio - Module 1 Project
# Asistente editorial para Revista Visión Conjunta

from datetime import datetime
import os

def extract_document_info(doc_path):
    """Read statistics from already-open Word document"""
    try:
        import win32com.client
        
        print("   Buscando Word abierto...")
        
        # Connect to already-running Word
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            print("   ✓ Word encontrado")
        except:
            return None, "No se encontró Word abierto. Por favor, abra el documento en Word primero."
        
        # Check if any document is open
        print(f"   Documentos abiertos: {word.Documents.Count}")
        if word.Documents.Count == 0:
            return None, "Word está abierto pero no hay ningún documento abierto."
        
        # Get the active document
        print("   Obteniendo documento activo...")
        doc = word.ActiveDocument
        doc_name = doc.Name
        print(f"   ✓ Documento: {doc_name}")
        
        # Get statistics
        print("   Leyendo estadísticas...")
        total_chars_body = doc.Characters.Count
        total_words = doc.Words.Count
        total_paragraphs = doc.Paragraphs.Count

        # Add footnote/endnote characters (they're part of the document!)
        footnote_chars = 0
        for fn in doc.Footnotes:
            footnote_chars += len(fn.Range.Text)
    
        for en in doc.Endnotes:
            footnote_chars += len(en.Range.Text)

        # Total characters INCLUDING footnotes/endnotes
        total_chars = total_chars_body + footnote_chars

        print(f"   ✓ Caracteres (cuerpo): {total_chars_body:,}")
        print(f"   ✓ Caracteres (notas): {footnote_chars:,}")
        print(f"   ✓ Total caracteres: {total_chars:,}")
        print(f"   ✓ Palabras: {total_words:,}")

        # Get text
        print("   Extrayendo texto...")
        full_text = doc.Content.Text
        print("   ✓ Texto extraído")
        
        info = {
            'full_text': full_text,
            'total_chars': total_chars,
            'total_words': total_words,
            'paragraph_count': total_paragraphs,
            'doc_name': doc_name
        }
        
        return info, None
    
    except Exception as e:
        print(f"   ✗ Excepción: {str(e)}")
        return None, str(e)

def check_format_compliance(info):
    """Check basic format compliance with Visión Conjunta guidelines"""
    issues = []
    recommendations = []
    
    # Check character count (16000-24000 for short, 36000-40000 for long)
    chars = info['total_chars']
    
    if chars < 16000:
        issues.append(f"❌ Extensión insuficiente: {chars:,} caracteres (mínimo: 16,000)")
        recommendations.append("Ampliar el contenido para alcanzar la extensión mínima de artículo corto")
    elif 16000 <= chars <= 24000:
        issues.append(f"✅ Extensión válida para artículo corto: {chars:,} caracteres")
    elif 24000 < chars < 36000:
        issues.append(f"⚠️ Extensión intermedia: {chars:,} caracteres (no cumple formato estándar)")
        recommendations.append("Ajustar a artículo corto (16,000-24,000) o largo (36,000-40,000)")
    elif 36000 <= chars <= 40000:
        issues.append(f"✅ Extensión válida para artículo largo: {chars:,} caracteres")
    else:
        issues.append(f"❌ Extensión excesiva: {chars:,} caracteres (máximo: 40,000)")
        recommendations.append("Reducir extensión para cumplir con límite de artículo largo")
    
    return issues, recommendations


def generate_report(doc_path, info, issues, recommendations):
    """Generate editorial review report"""
    report = f"""
╔════════════════════════════════════════════════════════════════╗
║              SILVINA - ASISTENTE EDITORIAL v0.1                ║
║         Revista Visión Conjunta - Informe de Revisión          ║
╚════════════════════════════════════════════════════════════════╝

📄 DOCUMENTO: {doc_path}
📅 FECHA DE REVISIÓN: {datetime.now().strftime('%d/%m/%Y %H:%M')}

═══════════════════════════════════════════════════════════════
📊 ANÁLISIS BÁSICO (Estadísticas de Word)
═══════════════════════════════════════════════════════════════

- Total de caracteres con espacios: {info['total_chars']:,}
- Total de palabras: {info['total_words']:,}
- Total de párrafos: {info['paragraph_count']}

═══════════════════════════════════════════════════════════════
🔍 CUMPLIMIENTO DE FORMATO
═══════════════════════════════════════════════════════════════

"""
    
    for issue in issues:
        report += f"{issue}\n"
    
    if recommendations:
        report += f"\n{'═'*63}\n💡 RECOMENDACIONES\n{'═'*63}\n\n"
        for i, rec in enumerate(recommendations, 1):
            report += f"{i}. {rec}\n"
    
    report += f"\n{'═'*63}\n"
    report += "📝 Nota: Esta es la versión básica (v0.1) de Silvina.\n"
    report += "Próximas versiones incluirán análisis APA, estilo y contenido.\n"
    report += f"{'═'*63}\n"
    
    return report


def review_document(doc_path):
    """Main function: Review document and generate report"""
    print("\n🔄 Conectando con Word abierto...")
    
    # Extract document info using Word
    info, error = extract_document_info(doc_path)
    
    if error:
        print(f"\n❌ Error: {error}")
        print("💡 Asegúrese de que:")
        print("   1. Word esté abierto")
        print("   2. El documento esté abierto en Word")
        print("   3. El documento sea el activo (ventana visible)")
        return
    
    # Check compliance
    issues, recommendations = check_format_compliance(info)
    
    # Generate report
    report = generate_report(doc_path, info, issues, recommendations)
    
    # Display report
    print(report)
    
    # Save report to file
    report_filename = f"reporte_silvina_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(report_filename, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"💾 Reporte guardado en: {report_filename}\n")


# Main program
if __name__ == "__main__":
    print("╔════════════════════════════════════════════════════════════════╗")
    print("║              SILVINA - ASISTENTE EDITORIAL v0.1                ║")
    print("╚════════════════════════════════════════════════════════════════╝\n")
    
    doc_path = input("📁 Ingrese la ruta del documento .docx: ").strip()
    
    # Remove quotes if user pasted path with them
    doc_path = doc_path.strip('"').strip("'")
    
    if doc_path:
        review_document(doc_path)
    else:
        print("❌ No se proporcionó ruta de documento")