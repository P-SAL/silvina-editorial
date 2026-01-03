"""
Silvina Editorial Assistant v0.5 - COMPLETE
Object-Oriented Refactor with Classes

v0.5 COMPLETE FEATURES:
- Article type detection (Divulgación/Científica with IMRyD)
- Enhanced Spanish APA 7 author validation (personal, organizational, et al.)
- Token calculator for LLM context management
- RAE grammar rules context for focused review
- Alphabetical order validation
- Spanish "y" vs "&" conjunction validation
- DOI/URL presence detection
- Old format detection ("Recuperado de")
- Spanish quotes validation (« » vs " ")
- Bibliografia vs Referencias section detection
- Duplicate reference detection
- Improved COM stability

Author: Pablo Salonio
Repository: https://github.com/P-SAL/silvina-editorial
"""

from datetime import datetime
import re
import win32com.client
import pythoncom
import time
import os
from difflib import SequenceMatcher


# === RAE GRAMMAR RULES CONTEXT ===
RAE_RULES_CONTEXT = """Reglas RAE para textos académicos (resumidas):

ERRORES COMUNES A DETECTAR:
1. Concordancia: sujeto-verbo, artículo-sustantivo
   ❌ "Los datos es claro" → ✅ "Los datos son claros"

2. Uso de comas: NO entre sujeto y verbo
   ❌ "El sistema cuántico, permite cifrado" → ✅ "El sistema cuántico permite cifrado"

3. Acentuación: palabras esdrújulas/sobresdrújulas
   ❌ "metodo" → ✅ "método"

4. Puntuación: punto después de abreviaturas
   ❌ "Dr Sánchez" → ✅ "Dr. Sánchez"

5. Gerundios incorrectos: NO para acciones posteriores
   ❌ "Se realizó el experimento, obteniendo resultados" 
   → ✅ "Se realizó el experimento y se obtuvieron resultados"

SOLO menciona errores EVIDENTES que veas en el texto."""


# === DOCUMENT CLASS ===
class Document:
    """Manages Word document loading and reference extraction."""
    
    def __init__(self, filepath):
        """Initialize with filepath only."""
        self.filepath = filepath #Stores the path to the Word file.
        self.word = None # Stores a reference to the Word application (COM object) that allows Python to control Microsoft Word
        self.doc = None # Stores the opened Word document object,
        self.text = ""
        self.references = []
        self.section_type = "Referencias"  # Default
    
    def load(self):
        """Load document and extract references."""
        self._connect_to_word()
        self._extract_referencias()
        self._create_reference_objects()
    
    def _connect_to_word(self):
        """Open Word document with robust COM initialization."""
        pythoncom.CoInitialize()
    
        try:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            abs_path = os.path.abspath(self.filepath)
            
            self.doc = self.word.Documents.Open(abs_path)
            
            time.sleep(2.0)
            self.doc.Activate()
            time.sleep(1.0)
            _ = self.doc.Characters.Count
            time.sleep(0.5)
            _ = len(self.doc.Paragraphs)
            
            print(f"✅ Connected: {abs_path}")
            print(f"✅ Document fully loaded: {len(self.doc.Paragraphs)} paragraphs")
            
        except Exception as e:
            print(f"❌ Connection Error: {e}")
            self.word = None
            self.doc = None
    
    def _extract_referencias(self):
        """Extract Referencias/Bibliografía section."""
        
        if not self.doc:
            print("⚠️ No document loaded")
            return
        
        try:
            time.sleep(1.0)
            
            char_count = self.get_character_count()
            if char_count == 0:
                print("⚠️ Document shows 0 characters - COM not ready")
                return
                
            print(f"🔍 Characters: {char_count:,}")
            
            try:
                total_paras = len(self.doc.Paragraphs)
                print(f"🔍 Total paragraphs: {total_paras}")
            except Exception as para_error:
                print(f"❌ Cannot access Paragraphs: {para_error}")
                return
            
            found_start = False
            referencias_paras = []
            
            for para in self.doc.Paragraphs:
                try:
                    para_text = para.Range.Text.strip()
                except:
                    continue
                
                if not found_start:
                    if "Bibliografía" in para_text:
                        self.section_type = "Bibliografía"
                        found_start = True
                        print(f"✅ Found Bibliografía section")
                        continue
                    elif "Fuentes bibliográficas" in para_text or "Referencias" in para_text:
                        self.section_type = "Referencias"
                        found_start = True
                        print(f"✅ Found Referencias section")
                        continue
                
                if found_start and para_text:
                    referencias_paras.append(para_text)
            
            self.text = '\n'.join(referencias_paras)
            print(f"✅ Extracted {len(referencias_paras)} reference paragraphs")
                            
        except Exception as e:
            print(f"❌ Extract error: {e}")
            import traceback
            traceback.print_exc()
            self.text = ""
    
    def _create_reference_objects(self):
        """Create Reference objects from extracted paragraphs."""
        if not self.text:
            return
        
        paragraphs = self.text.split('\n')
        
        for para in paragraphs:
            para = para.strip()
            if len(para) < 30:
                continue
            
            years = re.findall(r'\(\d{4}\)', para)
            
            if len(years) >= 2:
                split_pattern = r'\.(?=[A-Z][a-z]+,\s+[A-Z]\.)'
                parts = re.split(split_pattern, para, maxsplit=1)
                
                for part in parts:
                    part = part.strip()
                    if len(part) > 30:
                        if not part.endswith('.'):
                            part += '.'
                        self.references.append(Reference(part))
            else:
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
    
    def detectar_tipo_articulo(self):
        """Detecta el tipo de artículo según caracteres y estructura."""
        if not self.doc:
            return {
                'tipo': 'Indeterminado',
                'caracteres': 0,
                'cumple_limite': False,
                'mensaje': 'Documento no cargado'
            }
        
        caracteres = self.get_character_count()
        texto_completo = self.doc.Content.Text.lower()
        palabras_imryd = ['introducción', 'método', 'resultados', 'discusión', 'conclusión']
        tiene_imryd = sum(1 for palabra in palabras_imryd if palabra in texto_completo) >= 4
        
        if tiene_imryd and 30000 <= caracteres <= 50000:
            tipo = 'Científica'
            cumple = True
            mensaje = f'Artículo científico con {caracteres:,} caracteres (rango válido: 30,000-50,000)'
        elif tiene_imryd and caracteres > 50000:
            tipo = 'Científica'
            cumple = False
            mensaje = f'Excede límite científico: {caracteres:,} caracteres (máximo: 50,000)'
        elif tiene_imryd and caracteres < 30000:
            tipo = 'Científica'
            cumple = False
            mensaje = f'Debajo del mínimo científico: {caracteres:,} caracteres (mínimo: 30,000)'
        elif caracteres <= 35000:
            tipo = 'Divulgación'
            cumple = abs(caracteres - 30000) <= 5000
            if cumple:
                mensaje = f'Artículo de divulgación con {caracteres:,} caracteres (objetivo: ~30,000)'
            else:
                mensaje = f'Divulgación con {caracteres:,} caracteres (objetivo: ~30,000 ± 5,000)'
        else:
            tipo = 'Indeterminado'
            cumple = False
            mensaje = f'No se detectó estructura IMRyD, pero tiene {caracteres:,} caracteres (fuera del rango de Divulgación)'
        
        return {
            'tipo': tipo,
            'caracteres': caracteres,
            'cumple_limite': cumple,
            'mensaje': mensaje
        }
    
    def calcular_tokens(self, texto=None):
        """Estima tokens para validar si documento cabe en contexto LLM."""
        if texto is None:
            texto = self.doc.Content.Text if self.doc else ""
        
        caracteres = len(texto)
        tokens_estimados = caracteres // 4
        
        MAX_CONTEXT = 8192
        RESERVED_FOR_PROMPT = 500
        RESERVED_FOR_RESPONSE = 500
        
        contexto_disponible = MAX_CONTEXT - RESERVED_FOR_PROMPT - RESERVED_FOR_RESPONSE
        cabe = tokens_estimados <= contexto_disponible
        
        return {
            'caracteres': caracteres,
            'tokens_estimados': tokens_estimados,
            'cabe_en_contexto': cabe,
            'contexto_disponible': contexto_disponible,
            'porcentaje_uso': (tokens_estimados / contexto_disponible) * 100
        }
    
    def validar_orden_alfabetico(self):
        """Verifica si las referencias están en orden alfabético."""
        if not self.references or len(self.references) < 2:
            return {
                'ordenadas': True,
                'total_referencias': len(self.references),
                'problemas': []
            }
        
        problemas = []
        
        for i in range(len(self.references) - 1):
            ref_actual = self.references[i].text
            ref_siguiente = self.references[i + 1].text
            
            primera_palabra_actual = ref_actual.split()[0].rstrip('.,').lower()
            primera_palabra_siguiente = ref_siguiente.split()[0].rstrip('.,').lower()
            
            if primera_palabra_actual > primera_palabra_siguiente:
                problemas.append({
                    'posicion': i + 1,
                    'texto': ref_actual[:60] + '...' if len(ref_actual) > 60 else ref_actual,
                    'deberia_ir_antes_de': ref_siguiente[:60] + '...' if len(ref_siguiente) > 60 else ref_siguiente
                })
        
        return {
            'ordenadas': len(problemas) == 0,
            'total_referencias': len(self.references),
            'problemas': problemas
        }
    
    def detectar_duplicados(self):
        """
        Detecta referencias duplicadas o muy similares.
        
        Returns:
            dict: {
                'tiene_duplicados': bool,
                'duplicados': list of dicts with indices and similarity
            }
        """
        if len(self.references) < 2:
            return {
                'tiene_duplicados': False,
                'duplicados': []
            }
        
        duplicados = []
        
        for i in range(len(self.references)):
            for j in range(i + 1, len(self.references)):
                ref_i = self.references[i].text
                ref_j = self.references[j].text
                
                # Calculate similarity ratio
                similarity = SequenceMatcher(None, ref_i.lower(), ref_j.lower()).ratio()
                
                # If >85% similar, flag as potential duplicate
                if similarity > 0.85:
                    duplicados.append({
                        'ref1_index': i + 1,
                        'ref1_text': ref_i[:60] + '...' if len(ref_i) > 60 else ref_i,
                        'ref2_index': j + 1,
                        'ref2_text': ref_j[:60] + '...' if len(ref_j) > 60 else ref_j,
                        'similitud': f"{similarity * 100:.1f}%"
                    })
        
        return {
            'tiene_duplicados': len(duplicados) > 0,
            'duplicados': duplicados
        }
    
    def validar_comillas_espanolas(self):
        """
        Verifica uso de comillas españolas (« ») en vez de inglesas (" ").
        
        Returns:
            dict: {
                'usa_comillas_correctas': bool,
                'problemas': list of reference indices
            }
        """
        problemas = []
        
        for i, ref in enumerate(self.references):
            # Check for English-style quotes in reference
            if '"' in ref.text or "'" in ref.text:
                problemas.append({
                    'posicion': i + 1,
                    'texto': ref.text[:60] + '...' if len(ref.text) > 60 else ref.text
                })
        
        return {
            'usa_comillas_correctas': len(problemas) == 0,
            'problemas': problemas
        }
    
    def close(self):
        """Clean up Word connection."""
        try:
            if self.doc:
                self.doc.Close(SaveChanges=False)
            if self.word:
                self.word.Quit()
        except:
            pass

    def generate_report(self, include_llm=True):
        """Generate formatted validation report with optional LLM review."""
        if not self.references:
            return "No references found."
        
        report = []
        report.append("=" * 70)
        report.append("SILVINA - ASISTENTE EDITORIAL v0.5 COMPLETE")
        report.append("=" * 70)
        report.append(f"\nDocumento: {os.path.basename(self.filepath)}")
        report.append(f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        report.append(f"Caracteres totales: {self.get_character_count():,}")
        
        # ARTICLE TYPE SECTION
        info_tipo = self.detectar_tipo_articulo()
        report.append("\n" + "=" * 70)
        report.append("TIPO DE ARTÍCULO Y CUMPLIMIENTO EUMIC")
        report.append("=" * 70)
        report.append(f"Tipo detectado: {info_tipo['tipo']}")
        report.append(f"Caracteres: {info_tipo['caracteres']:,}")
        report.append(f"{'✅' if info_tipo['cumple_limite'] else '⚠️'} {info_tipo['mensaje']}")

        # LLM REVIEW SECTION (WITHOUT token analysis - moved to end)
        if include_llm:
            report.append("\n" + "=" * 70)
            report.append("REVISIÓN DE GRAMÁTICA Y ESTILO (LLM)")
            report.append("=" * 70)
            
            try:
                print("\n🤖 Analizando con LLM...")
                info_tokens = self.calcular_tokens()  # Calculate but don't display yet
                llm_review, llm_error = self.review_with_llm(info_tokens)
                if llm_error:
                    report.append(f"\n⚠️ {llm_error}")
                else:
                    report.append(f"\n{llm_review}")
                    
            except Exception as e:
                report.append(f"\n❌ Error en análisis LLM: {str(e)}")
                info_tokens = None  # Set to None if error
        else:
            info_tokens = None
        
        # REFERENCES VALIDATION SECTION
        report.append("\n" + "=" * 70)
        report.append("VALIDACIÓN DE REFERENCIAS APA")
        report.append("=" * 70)
        report.append(f"Tipo de sección: {self.section_type}")
        report.append(f"Referencias encontradas: {len(self.references)}")
        
        # Count valid/invalid
        valid_count = sum(1 for ref in self.references if ref.is_valid())
        invalid_count = len(self.references) - valid_count
        
        report.append(f"✅ Válidas: {valid_count}")
        report.append(f"❌ Con problemas: {invalid_count}")
        
        # ADDITIONAL CHECKS
        orden_info = self.validar_orden_alfabetico()
        duplicados_info = self.detectar_duplicados()
        comillas_info = self.validar_comillas_espanolas()
        
        # Summary of additional validations
        if orden_info['ordenadas']:
            report.append(f"✅ Referencias en orden alfabético")
        else:
            report.append(f"⚠️ Referencias NO están en orden alfabético ({len(orden_info['problemas'])} problemas)")
        
        if not duplicados_info['tiene_duplicados']:
            report.append(f"✅ No se detectaron referencias duplicadas")
        else:
            report.append(f"⚠️ Posibles duplicados encontrados: {len(duplicados_info['duplicados'])}")
        
        if comillas_info['usa_comillas_correctas']:
            report.append(f"✅ Comillas españolas correctas")
        else:
            report.append(f"⚠️ Uso de comillas inglesas en {len(comillas_info['problemas'])} referencias")
        
        # DOI/URL Summary
        refs_con_doi = sum(1 for ref in self.references if ref.tiene_doi_o_url()['tiene_doi'])
        refs_con_url = sum(1 for ref in self.references if ref.tiene_doi_o_url()['tiene_url'])
        refs_formato_antiguo = sum(1 for ref in self.references if ref.tiene_doi_o_url()['formato_antiguo'])
        
        report.append(f"📊 DOI: {refs_con_doi}/{len(self.references)} | URL: {refs_con_url}/{len(self.references)}")
        if refs_formato_antiguo > 0:
            report.append(f"⚠️ {refs_formato_antiguo} referencias usan formato antiguo ('Recuperado de')")
        
        report.append("\n" + "-" * 70)
        report.append("DETALLE DE VALIDACIÓN")
        report.append("-" * 70 + "\n")
        
        for i, ref in enumerate(self.references, 1):
            rep = ref.get_validation_report()
            
            if rep['is_valid']:
                # ✅ VALID: Show status only, no text
                report.append(f"{i}. ✅ VÁLIDA")
            else:
                # ❌ INVALID: Show status AND full details
                report.append(f"{i}. ❌ REQUIERE REVISIÓN")
                report.append(f"   Texto: {rep['text']}")
                
                if not rep['valid_author']:
                    report.append("   ⚠️ Formato de autor incorrecto (debe ser: Apellido, I.)")
                if not rep['valid_year']:
                    report.append("   ⚠️ Año no encontrado o formato incorrecto (debe ser: (YYYY))")
                if not rep['valid_conjuncion']:
                    report.append(f"   ⚠️ {rep['error_conjuncion']}")
                
                # DOI/URL info (only for invalid references)
                doi_url_info = rep['doi_url_info']
                if not doi_url_info['tiene_doi'] and not doi_url_info['tiene_url']:
                    report.append("   ℹ️ Sin DOI ni URL")
                if doi_url_info['formato_antiguo']:
                    report.append("   ⚠️ Usa formato antiguo 'Recuperado de' (debe omitirse)")
            
            report.append("")
        
        # ALPHABETICAL ORDER PROBLEMS DETAIL
        if not orden_info['ordenadas']:
            report.append("\n" + "-" * 70)
            report.append("PROBLEMAS DE ORDEN ALFABÉTICO")
            report.append("-" * 70 + "\n")
            
            for problema in orden_info['problemas']:
                report.append(f"⚠️ Referencia #{problema['posicion']}:")
                report.append(f"   {problema['texto']}")
                report.append(f"   Debería ir ANTES de: {problema['deberia_ir_antes_de']}\n")
        
        # DUPLICATE REFERENCES DETAIL
        if duplicados_info['tiene_duplicados']:
            report.append("\n" + "-" * 70)
            report.append("POSIBLES REFERENCIAS DUPLICADAS")
            report.append("-" * 70 + "\n")
            
            for dup in duplicados_info['duplicados']:
                report.append(f"⚠️ Referencias #{dup['ref1_index']} y #{dup['ref2_index']} son {dup['similitud']} similares:")
                report.append(f"   #{dup['ref1_index']}: {dup['ref1_text']}")
                report.append(f"   #{dup['ref2_index']}: {dup['ref2_text']}\n")
        
        # SPANISH QUOTES PROBLEMS
        if not comillas_info['usa_comillas_correctas']:
            report.append("\n" + "-" * 70)
            report.append("PROBLEMAS DE COMILLAS")
            report.append("-" * 70 + "\n")
            
            for problema in comillas_info['problemas']:
                report.append(f"⚠️ Referencia #{problema['posicion']} usa comillas inglesas:")
                report.append(f"   {problema['texto']}")
                report.append(f"   Debe usar comillas españolas (« »)\n")
        
        # TOKEN ANALYSIS (MOVED TO END - Technical info about Silvina)
        if include_llm and info_tokens:
            report.append("\n" + "=" * 70)
            report.append("ANÁLISIS TÉCNICO - CAPACIDAD LLM")
            report.append("=" * 70)
            report.append(f"Caracteres analizados: {info_tokens['caracteres']:,}")
            report.append(f"Tokens estimados: {info_tokens['tokens_estimados']:,}")
            report.append(f"Uso de contexto: {info_tokens['porcentaje_uso']:.1f}%")
            
            if not info_tokens['cabe_en_contexto']:
                report.append(f"⚠️ Documento excede contexto LLM - análisis parcial")
            else:
                report.append(f"✅ Documento completo analizado")
        
        report.append("\n" + "=" * 70)
        
        return '\n'.join(report)
   

    def review_with_llm(self, info_tokens):
        """Revisión gramatical con contexto RAE."""
        try:
            import ollama
            
            full_text = self.doc.Content.Text if self.doc else ""
            MAX_SAMPLE = 2000
            sample = full_text[:MAX_SAMPLE]
            
            prompt = f"""Eres un corrector de textos académicos en español.

INSTRUCCIÓN ÚNICA: Revisa este texto y lista SOLO errores gramaticales EVIDENTES.

PROHIBIDO:
- NO sugieras cambios de estilo
- NO comentes sobre estructura
- NO menciones títulos o keywords
- NO des consejos generales

Si NO HAY ERRORES, escribe EXACTAMENTE: "No se detectaron errores gramaticales."

TEXTO:
{sample}"""
            
            response = ollama.chat(
                model='llama3-gradient:8b',
                messages=[{'role': 'user', 'content': prompt}],
                options={
                    'num_predict': 500,
                    'temperature': 0.1
                }
            )
            
            return response['message']['content'], None
        
        except ImportError:
            return None, "Módulo 'ollama' no instalado"
        except Exception as e:
            return None, f"Error LLM: {str(e)}"


# === REFERENCE CLASS ===
class Reference:
    """Represents a single bibliographic reference"""
    
    def __init__(self, text):
        """Initialize reference with citation text"""
        self.text = text
    
    def validate_author(self):
        """Check if reference has valid APA 7 Spanish author format."""
        personal = r'[A-ZÁ-ÚÑ][a-zá-úñ]+(?:-[A-ZÁ-ÚÑ][a-zá-úñ]+)?,\s+[A-Z]\.'
        et_al = r'et\s+al\.'
        organizational = r'^[A-Z][A-Za-z\s&,\-]{10,}\.\s'
        
        has_personal = bool(re.search(personal, self.text))
        has_et_al = bool(re.search(et_al, self.text, re.IGNORECASE))
        has_organizational = bool(re.search(organizational, self.text))
        
        if has_organizational and not has_personal:
            return True
        
        return has_personal or has_et_al
    
    def validate_year(self):
        """Check if reference has valid year format (YYYY)"""
        pattern = r'\((\d{4})\)'
        match = re.search(pattern, self.text)
        if match:
            return True, match.group(1)
        return False, None
    
    def validar_conjuncion_espanola(self):
        """
        Verifica uso de 'y' en vez de '&' para referencias en español APA 7.
        Improved pattern to catch all cases.
        """
        # Pattern catches: "I. &" or "I., &" or "A., &"
        patron_ampersand = r'[A-Z]\.(?:,)?\s+&\s+[A-Z]'
        
        if re.search(patron_ampersand, self.text):
            return False, "Uso incorrecto de '&' (debe ser 'y' en español APA 7)"
        
        return True, None
    
    def tiene_doi_o_url(self):
        """
        Verifica presencia de DOI o URL.
        
        Returns:
            dict: {
                'tiene_doi': bool,
                'tiene_url': bool,
                'formato_antiguo': bool  # "Recuperado de"
            }
        """
        tiene_doi = bool(re.search(r'https?://doi\.org/[\w\.\-/]+', self.text, re.IGNORECASE))
        tiene_url = bool(re.search(r'https?://[^\s]+', self.text))
        formato_antiguo = bool(re.search(r'Recuperado\s+de\s+https?://', self.text, re.IGNORECASE))
        
        return {
            'tiene_doi': tiene_doi,
            'tiene_url': tiene_url,
            'formato_antiguo': formato_antiguo
        }
    
    def is_valid(self):
        """Check if reference meets all APA 7 Spanish requirements"""
        has_author = self.validate_author()
        has_year, _ = self.validate_year()
        conjuncion_valida, _ = self.validar_conjuncion_espanola()
        
        return has_author and has_year and conjuncion_valida
    
    def get_validation_report(self):
        """Return detailed validation results"""
        has_author = self.validate_author()
        has_year, year = self.validate_year()
        conjuncion_valida, error_conjuncion = self.validar_conjuncion_espanola()
        doi_url_info = self.tiene_doi_o_url()
        
        return {
            'text': self.text[:80] + '...' if len(self.text) > 80 else self.text,
            'valid_author': has_author,
            'valid_year': has_year,
            'valid_conjuncion': conjuncion_valida,
            'error_conjuncion': error_conjuncion,
            'doi_url_info': doi_url_info,
            'year': year,
            'is_valid': has_author and has_year and conjuncion_valida
        }


# === MAIN EXECUTION ===
if __name__ == "__main__":
    print("\n" + "="*70)
    print("SILVINA v0.5 - ASISTENTE EDITORIAL - COMPLETE")
    print("="*70 + "\n")
    
    filepath = r"C:\Users\usuario\Desktop\Escudo cuantico_AB_25092025.docx"
    
    doc = Document(filepath)
    doc.load()
    
    report = doc.generate_report(include_llm=True)
    print(report)
    
    report_filename = f"reporte_silvina_v05_COMPLETE_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(report_filename, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"\n💾 Reporte guardado: {report_filename}")
    
    doc.close()