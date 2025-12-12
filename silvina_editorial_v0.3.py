"""
Silvina Editorial Assistant v0.3
APA Citation Format Validator for Revista Visión Conjunta

CHANGELOG v0.3:
- NEW: Extract Referencias/Bibliografía section from Word documents
- NEW: Validate APA 7 citation format using Regular Expressions
- NEW: Check author name patterns (Apellido, I.)
- NEW: Validate year format (YYYY)
- NEW: Generate APA compliance report
- KEPT: All v0.2 features (character count, LLM review)

Author: Pablo Salonio
Repository: https://github.com/P-SAL/silvina-editorial
"""

from datetime import datetime
import os
import re  # NEW: Regular expressions for APA validation

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

        # Add footnote/endnote characters
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


# NEW v0.3: Extract Referencias section
def extract_references_section(full_text):
    """
    Extract Referencias/Bibliografía section from document text.
    
    Args:
        full_text (str): Complete document text
        
    Returns:
        str: Text of references section, or None if not found
    """
    # Look for common Spanish reference section headers
    headers = ["Referencias", "Bibliografía", "REFERENCIAS", "BIBLIOGRAFÍA"]
    
    for header in headers:
        if header in full_text:
            # Find position of header
            start_pos = full_text.find(header)
            # Extract from header to end of document
            referencias_text = full_text[start_pos:]
            print(f"   ✓ Sección encontrada: '{header}'")
            return referencias_text
    
    print("   ⚠️ No se encontró sección de Referencias")
    return None

# NEW v0.3: Diagnostic function
def find_possible_reference_headers(full_text):
    """
    Search for possible reference section headers in document.
    Helps diagnose why Referencias section wasn't found.
    
    Args:
        full_text (str): Complete document text
        
    Returns:
        list: Possible headers found
    """
    # Common variations to search for
    possible_headers = [
        "Referencias",
        "Bibliografía", 
        "REFERENCIAS",
        "BIBLIOGRAFÍA",
        "Referencias bibliográficas",
        "REFERENCIAS BIBLIOGRÁFICAS",
        "Fuentes",
        "FUENTES",
        "Bibliografía consultada",
        "Literatura citada"
    ]
    
    found = []
    for header in possible_headers:
        if header in full_text:
            # Get context around the header (20 chars before and after)
            pos = full_text.find(header)
            context = full_text[max(0, pos-20):pos+len(header)+20]
            found.append((header, context))
    
    return found

# NEW v0.3: Validate author name pattern
def validate_author_pattern(reference):
    """
    Check if reference contains proper APA author format: Apellido, I.
    
    Args:
        reference (str): Single reference line
        
    Returns:
        bool: True if valid author pattern found
    """
    # Pattern: Apellido, I. (supports Spanish characters)
    # Matches: García, M. | López, J. A. | Pérez-Sánchez, C.
    author_pattern = r'[A-ZÁ-ÚÑ][a-zá-úñ]+(?:-[A-ZÁ-ÚÑ][a-zá-úñ]+)?,\s[A-Z]\.'
    
    return bool(re.search(author_pattern, reference))


# NEW v0.3: Validate year format
def validate_year_pattern(reference):
    """
    Check if reference contains year in parentheses: (YYYY)
    
    Args:
        reference (str): Single reference line
        
    Returns:
        tuple: (bool, str) - (valid, year_found)
    """
    # Pattern: (2015) | (2020) | (2023)
    year_pattern = r'\((\d{4})\)'
    
    match = re.search(year_pattern, reference)
    if match:
        return True, match.group(1)
    return False, None


# NEW v0.3: Basic APA validation
def check_apa_compliance(referencias_text):
    """
    Validate APA format compliance in Referencias section.
    
    Args:
        referencias_text (str): Text of Referencias section
        
    Returns:
        dict: Validation results with issues found
    """
    if not referencias_text:
        return None
    
    # Split into individual references (by line breaks)
    # Skip the header line
    lines = referencias_text.split('\n')
    references = [line.strip() for line in lines if line.strip() and len(line.strip()) > 50]
    
    results = {
        'total_refs': len(references),
        'valid_author': 0,
        'valid_year': 0,
        'issues': []
    }
    
    for i, ref in enumerate(references, 1):
        # Check author pattern
        has_author = validate_author_pattern(ref)
        if has_author:
            results['valid_author'] += 1
        else:
            results['issues'].append(f"Ref {i}: Formato de autor incorrecto")
        
        # Check year pattern
        has_year, year = validate_year_pattern(ref)
        if has_year:
            results['valid_year'] += 1
        else:
            results['issues'].append(f"Ref {i}: Año no encontrado o formato incorrecto")
    
    return results


def check_format_compliance(info):
    """Check basic format compliance with Visión Conjunta guidelines"""
    issues = []
    recommendations = []
    
    # Check character count
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


def check_grammar_style(text, max_chars=3000):
    """
    Use local LLM to review grammar and style (from v0.2)
    Only analyzes first portion of text to avoid overwhelming small model
    """
    try:
        import ollama
        
        # Truncate if too long for small model
        sample = text[:max_chars]
        if len(text) > max_chars:
            sample += "\n\n[...texto truncado para análisis...]"
        
        prompt = f"""Eres un revisor editorial de textos académicos en español para una revista científica.

Analiza este fragmento y proporciona:
1. Principales errores gramaticales (máximo 3)
2. Sugerencias de estilo académico (máximo 3)
3. Calificación: Excelente/Bueno/Necesita revisión

Sé conciso y profesional.

TEXTO:
{sample}"""

        response = ollama.chat(
            model='llama3.2:1b',
            messages=[{'role': 'user', 'content': prompt}]
        )
        
        return response['message']['content'], None
    
    except ImportError:
        return None, "Módulo 'ollama' no instalado (pip install ollama)"
    except Exception as e:
        return None, f"Error LLM: {str(e)}"


# NEW v0.3: Enhanced report with APA validation
def generate_report(doc_path, info, issues, recommendations, llm_review=None, apa_results=None):
    """Generate editorial review report (enhanced in v0.3 with APA validation)"""
    report = f"""
╔════════════════════════════════════════════════════════════════╗
║              SILVINA - ASISTENTE EDITORIAL v0.3                ║
║         Revista Visión Conjunta - Informe de Revisión          ║
╚════════════════════════════════════════════════════════════════╝

📄 DOCUMENTO: {doc_path}
📅 FECHA DE REVISIÓN: {datetime.now().strftime('%d/%m/%Y %H:%M')}

═══════════════════════════════════════════════════════════════
📊 ANÁLISIS BÁSICO
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
        report += f"\n{'═'*63}\n💡 RECOMENDACIONES DE FORMATO\n{'═'*63}\n\n"
        for i, rec in enumerate(recommendations, 1):
            report += f"{i}. {rec}\n"
    
    # NEW v0.3: APA validation results
    if apa_results:
        report += f"\n{'═'*63}\n📚 VALIDACIÓN DE REFERENCIAS APA 7\n{'═'*63}\n\n"
        report += f"Total de referencias encontradas: {apa_results['total_refs']}\n"
        report += f"Referencias con formato de autor válido: {apa_results['valid_author']}/{apa_results['total_refs']}\n"
        report += f"Referencias con año válido: {apa_results['valid_year']}/{apa_results['total_refs']}\n"
        
        if apa_results['issues']:
            report += f"\n⚠️ PROBLEMAS ENCONTRADOS:\n"
            for issue in apa_results['issues'][:10]:  # Show max 10 issues
                report += f"   • {issue}\n"
            if len(apa_results['issues']) > 10:
                report += f"   ... y {len(apa_results['issues']) - 10} más\n"
    
    # Add LLM review if available (from v0.2)
    if llm_review:
        report += f"\n{'═'*63}\n🤖 REVISIÓN DE GRAMÁTICA Y ESTILO (LLM)\n{'═'*63}\n\n"
        report += llm_review + "\n"
    
    report += f"\n{'═'*63}\n"
    report += "📝 Versión 0.3 - Incluye validación APA + revisión LLM\n"
    report += f"{'═'*63}\n"
    
    return report


def review_document(doc_path, use_llm=True, check_apa=True):  # NEW: check_apa parameter
    """Main function: Review document and generate report (enhanced v0.3)"""
    print("\n🔄 Conectando con Word abierto...")
    
    # Extract document info
    info, error = extract_document_info(doc_path)
    
    if error:
        print(f"\n❌ Error: {error}")
        print("💡 Asegúrese de que:")
        print("   1. Word esté abierto")
        print("   2. El documento esté abierto en Word")
        print("   3. El documento sea el activo (ventana visible)")
        return
    
    # Check format compliance
    issues, recommendations = check_format_compliance(info)
    
    # NEW v0.3: APA citation validation
    apa_results = None
    if check_apa:
        print("\n📚 Validando referencias APA...")
        referencias_text = extract_references_section(info['full_text'])
        if referencias_text:
            apa_results = check_apa_compliance(referencias_text)
            print(f"   ✓ {apa_results['total_refs']} referencias analizadas")
        else:
            print("   ⚠️ Validación APA omitida (sección no encontrada)")
    
    # LLM grammar/style review (from v0.2)
    llm_review = None
    if use_llm:
        print("\n🤖 Analizando gramática y estilo con LLM...")
        llm_review, llm_error = check_grammar_style(info['full_text'])
        if llm_error:
            print(f"   ⚠️ {llm_error}")
            print("   ℹ️ Continuando sin revisión LLM...")
        else:
            print("   ✓ Revisión LLM completada")
    
    # Generate report (NEW: includes apa_results)
    report = generate_report(doc_path, info, issues, recommendations, llm_review, apa_results)
    
    # Display report
    print(report)
    
    # Save report
    report_filename = f"reporte_silvina_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(report_filename, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"💾 Reporte guardado en: {report_filename}\n")


# Main program
if __name__ == "__main__":
    print("╔════════════════════════════════════════════════════════════════╗")
    print("║              SILVINA - ASISTENTE EDITORIAL v0.3                ║")
    print("║                  + Validación APA 7                            ║")
    print("╚════════════════════════════════════════════════════════════════╝\n")
    
    # NEW: Diagnostic mode
    print("Modos disponibles:")
    print("  1. Revisión completa (normal)")
    print("  2. Diagnóstico de Referencias (buscar encabezados)\n")
    
    mode = input("Seleccione modo (1/2, Enter=1): ").strip()
    
    if mode == "2":
        # Diagnostic mode
        print("\n🔍 MODO DIAGNÓSTICO\n")
        doc_path = input("📁 Ingrese la ruta del documento .docx: ").strip()
        doc_path = doc_path.strip('"').strip("'")
        
        if doc_path:
            print("\n🔄 Conectando con Word...")
            info, error = extract_document_info(doc_path)
            
            if error:
                print(f"\n❌ Error: {error}")
            else:
                print("\n🔍 Buscando posibles encabezados de Referencias...\n")
                found = find_possible_reference_headers(info['full_text'])
                
                if found:
                    print(f"✅ Encontrados {len(found)} posibles encabezados:\n")
                    for header, context in found:
                        print(f"  • '{header}'")
                        print(f"    Contexto: ...{context}...\n")
                else:
                    print("❌ No se encontraron encabezados de referencias")
                    print("\n💡 El documento podría usar un formato diferente.")
                    print("   Muestre las últimas líneas del documento:\n")
                    print(info['full_text'][-500:])
    else:
        # Normal mode
        doc_path = input("📁 Ingrese la ruta del documento .docx: ").strip()
        doc_path = doc_path.strip('"').strip("'")
        
        if doc_path:
            # Ask if user wants APA validation
            check_apa_input = input("📚 ¿Validar referencias APA? (s/n, Enter=sí): ").strip().lower()
            check_apa = check_apa_input != 'n'
            
            # Ask if user wants LLM review
            use_llm_input = input("🤖 ¿Usar revisión LLM? (s/n, Enter=sí): ").strip().lower()
            use_llm = use_llm_input != 'n'
            
            review_document(doc_path, use_llm, check_apa)
        else:
            print("❌ No se proporcionó ruta de documento")