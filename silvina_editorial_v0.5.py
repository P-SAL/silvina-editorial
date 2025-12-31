"""
Silvina Editorial Assistant v0.5
Object-Oriented Refactor with Classes

NEW in v0.5:
- Article type detection (Divulgación/Científica)
- Enhanced Spanish APA 7 author validation (personal, organizational, et al.)
- Token calculator for LLM context management
- RAE grammar rules context for focused review
- Improved COM stability with HomeKey fix

Author: Pablo Salonio
Repository: https://github.com/P-SAL/silvina-editorial
"""

from datetime import datetime
import re
import win32com.client
import pythoncom
import time
import os


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
        """Open Word document with robust COM initialization."""
        pythoncom.CoInitialize()
    
        try:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False  # Keep invisible for stability
            abs_path = os.path.abspath(self.filepath)
            
            # Open document
            self.doc = self.word.Documents.Open(abs_path)
            
            # CRITICAL: Force Word to fully load document
            # Multiple techniques to ensure COM object is ready
            time.sleep(2.0)  # Initial wait
            
            # Force document activation
            self.doc.Activate()
            time.sleep(1.0)
            
            # Access a property to force full initialization
            _ = self.doc.Characters.Count
            time.sleep(0.5)
            
            # Verify we can access paragraphs
            _ = len(self.doc.Paragraphs)
            
            print(f"✅ Connected: {abs_path}")
            print(f"✅ Document fully loaded: {len(self.doc.Paragraphs)} paragraphs")
            
        except Exception as e:
            print(f"❌ Connection Error: {e}")
            self.word = None
            self.doc = None
    
    def _extract_referencias(self):
        """Extract Referencias section using paragraphs (no truncation)."""
        
        if not self.doc:
            print("⚠️ No document loaded")
            return
        
        try:
            time.sleep(1.0)
            
            # Verify document is accessible
            char_count = self.get_character_count()
            if char_count == 0:
                print("⚠️ Document shows 0 characters - COM not ready")
                return
                
            print(f"🔍 Characters: {char_count:,}")
            
            # Check if Paragraphs is accessible
            try:
                total_paras = len(self.doc.Paragraphs)
                print(f"🔍 Total paragraphs: {total_paras}")
            except Exception as para_error:
                print(f"❌ Cannot access Paragraphs: {para_error}")
                return
            
            # Find the paragraph with referencias heading
            found_start = False
            referencias_paras = []
            
            for para in self.doc.Paragraphs:
                try:
                    para_text = para.Range.Text.strip()
                except:
                    continue  # Skip problematic paragraphs
                
                # Check if this is the heading
                if not found_start:
                    if "Fuentes bibliográficas" in para_text or "Referencias" in para_text or "Bibliografía" in para_text:
                        found_start = True
                        print(f"✅ Found Referencias section")
                        continue  # Skip heading itself
                
                # After heading, collect all remaining paragraphs
                if found_start and para_text:
                    referencias_paras.append(para_text)
            
            # Join paragraphs with newlines
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
    
    def detectar_tipo_articulo(self):
        """
        Detecta el tipo de artículo según caracteres y estructura.
        
        Tipos EUMIC:
        - Divulgación: ~30,000 caracteres
        - Científica: 30,000-50,000 caracteres con estructura IMRyD
        
        Returns:
            dict: {
                'tipo': 'Divulgación'/'Científica'/'Indeterminado',
                'caracteres': int,
                'cumple_limite': bool,
                'mensaje': str
            }
        """
        if not self.doc:
            return {
                'tipo': 'Indeterminado',
                'caracteres': 0,
                'cumple_limite': False,
                'mensaje': 'Documento no cargado'
            }
        
        caracteres = self.get_character_count()
        
        # Check for IMRyD structure (Científica indicator)
        texto_completo = self.doc.Content.Text.lower()
        palabras_imryd = ['introducción', 'método', 'resultados', 'discusión', 'conclusión']
        tiene_imryd = sum(1 for palabra in palabras_imryd if palabra in texto_completo) >= 4
        
        # Determine type
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
        
        elif caracteres <= 35000:  # Likely Divulgación
            tipo = 'Divulgación'
            cumple = abs(caracteres - 30000) <= 5000  # ±5k tolerance
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
        """
        Estima tokens para validar si documento cabe en contexto LLM.
        
        Regla aproximada: 1 token ≈ 4 caracteres en español
        llama3 context: 8,192 tokens (máximo seguro)
        
        Returns:
            dict: {
                'caracteres': int,
                'tokens_estimados': int,
                'cabe_en_contexto': bool,
                'contexto_disponible': int,
                'porcentaje_uso': float
            }
        """
        if texto is None:
            texto = self.doc.Content.Text if self.doc else ""
        
        caracteres = len(texto)
        tokens_estimados = caracteres // 4  # Aproximación: 4 chars = 1 token
        
        # llama3/gradient context window
        MAX_CONTEXT = 8192
        RESERVED_FOR_PROMPT = 500  # Para instrucciones + RAE rules
        RESERVED_FOR_RESPONSE = 500  # Para respuesta del LLM
        
        contexto_disponible = MAX_CONTEXT - RESERVED_FOR_PROMPT - RESERVED_FOR_RESPONSE
        cabe = tokens_estimados <= contexto_disponible
        
        return {
            'caracteres': caracteres,
            'tokens_estimados': tokens_estimados,
            'cabe_en_contexto': cabe,
            'contexto_disponible': contexto_disponible,
            'porcentaje_uso': (tokens_estimados / contexto_disponible) * 100
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
        report.append("SILVINA - ASISTENTE EDITORIAL v0.5")
        report.append("=" * 70)
        report.append(f"\nDocumento: {os.path.basename(self.filepath)}")
        report.append(f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        report.append(f"Caracteres totales: {self.get_character_count():,}")
        
        # ARTICLE TYPE SECTION (NEW IN V0.5)
        info_tipo = self.detectar_tipo_articulo()
        report.append("\n" + "=" * 70)
        report.append("TIPO DE ARTÍCULO Y CUMPLIMIENTO EUMIC")
        report.append("=" * 70)
        report.append(f"Tipo detectado: {info_tipo['tipo']}")
        report.append(f"Caracteres: {info_tipo['caracteres']:,}")
        report.append(f"{'✅' if info_tipo['cumple_limite'] else '⚠️'} {info_tipo['mensaje']}")

        # LLM REVIEW SECTION (IMPROVED IN V0.5)
        if include_llm:
            report.append("\n" + "=" * 70)
            report.append("REVISIÓN DE GRAMÁTICA Y ESTILO (LLM)")
            report.append("=" * 70)
            
            try:
                # Token analysis
                info_tokens = self.calcular_tokens()
                report.append(f"\n📊 Análisis de tokens:")
                report.append(f"   Caracteres: {info_tokens['caracteres']:,}")
                report.append(f"   Tokens estimados: {info_tokens['tokens_estimados']:,}")
                report.append(f"   Uso de contexto: {info_tokens['porcentaje_uso']:.1f}%")
                
                if not info_tokens['cabe_en_contexto']:
                    report.append(f"   ⚠️ Documento excede contexto LLM")
                    report.append(f"   Se analizará solo los primeros ~{info_tokens['contexto_disponible']:,} tokens\n")
                else:
                    report.append(f"   ✅ Documento cabe en contexto LLM\n")
                
                # LLM analysis
                print("\n🤖 Analizando con LLM...")
                llm_review, llm_error = self.review_with_llm(info_tokens)
                if llm_error:
                    report.append(f"\n⚠️ {llm_error}")
                else:
                    report.append(f"\n{llm_review}")
                    
            except Exception as e:
                report.append(f"\n❌ Error en análisis LLM: {str(e)}")
        
        # REFERENCES VALIDATION SECTION
        report.append("\n" + "=" * 70)
        report.append("VALIDACIÓN DE REFERENCIAS APA")
        report.append("=" * 70)
        report.append(f"Referencias encontradas: {len(self.references)}")
        
        # Count valid/invalid
        valid_count = sum(1 for ref in self.references if ref.is_valid())
        invalid_count = len(self.references) - valid_count
        
        report.append(f"✅ Válidas: {valid_count}")
        report.append(f"❌ Con problemas: {invalid_count}")
        
        report.append("\n" + "-" * 70)
        report.append("DETALLE DE VALIDACIÓN")
        report.append("-" * 70 + "\n")
        
        for i, ref in enumerate(self.references, 1):
            rep = ref.get_validation_report()

            if rep['is_valid']:
                # ✅ Valid reference: status only
                report.append(f"{i}. ✅ VÁLIDA")
            else:
                # ❌ Invalid reference: show details
                report.append(f"{i}. ❌ REQUIERE REVISIÓN")
                report.append(f"   Texto: {rep['text']}")

                if not rep['valid_author']:
                    report.append("   ⚠️ Formato de autor incorrecto (debe ser: Apellido, I.)")
                if not rep['valid_year']:
                    report.append("   ⚠️ Año no encontrado o formato incorrecto (debe ser: (YYYY))")

        report.append("" * 70)
           
        return '\n'.join(report)

    def review_with_llm(self, info_tokens):
        """
        Revisión gramatical con contexto RAE (DIAGNOSTIC VERSION).
        """
        try:
            import ollama
            
            print("   🔍 DEBUG: Starting LLM review...")
            
            # Get document text
            full_text = self.doc.Content.Text if self.doc else ""
            print(f"   🔍 DEBUG: Full text length: {len(full_text)} chars")
            
            # AGGRESSIVE TRUNCATION for testing
            MAX_SAMPLE = 2000  # Much smaller for testing
            sample = full_text[:MAX_SAMPLE]
            print(f"   🔍 DEBUG: Sample length: {len(sample)} chars")
            
            # SIMPLIFIED PROMPT (no RAE context for now)
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
               
       
            print(f"   🔍 DEBUG: Prompt length: {len(prompt)} chars")
            print("   🔍 DEBUG: Calling Ollama...")
            
            # Call with minimal options
            response = ollama.chat(
                model='llama3-gradient:8b',
                messages=[{'role': 'user', 'content': prompt}],
                options={
                    'num_predict': 500,  # Very short response
                    'temperature': 0.1
                }
            )
            
            print("   ✅ DEBUG: LLM responded!")
            return response['message']['content'], None
        
        except ImportError:
            return None, "Módulo 'ollama' no instalado"
        except Exception as e:
            print(f"   ❌ DEBUG: Exception caught: {type(e).__name__}")
            return None, f"Error LLM: {str(e)}"
        
            
# === REFERENCE CLASS ===
class Reference:
    """Represents a single bibliographic reference"""
    
    def __init__(self, text):
        """Initialize reference with citation text"""
        self.text = text
    
    def validate_author(self):
        """
        Check if reference has valid APA 7 Spanish author format.
        
        Accepts:
        - Personal: Apellido, I.
        - Multiple: Apellido, I. y Apellido, J.
        - Et al: Apellido, I. et al.
        - Organizational: Google, IBM Research, National Institute...
        """
        # Pattern 1: Personal author (Apellido, I.)
        personal = r'[A-ZÁ-ÚÑ][a-zá-úñ]+(?:-[A-ZÁ-ÚÑ][a-zá-úñ]+)?,\s+[A-Z]\.'
        
        # Pattern 2: Et al. format
        et_al = r'et\s+al\.'
        
        # Pattern 3: Organizational author
        # Starts with capital, has multiple words, ends with period
        organizational = r'^[A-Z][A-Za-z\s&,\-]{10,}\.\s'
        
        # Check matches
        has_personal = bool(re.search(personal, self.text))
        has_et_al = bool(re.search(et_al, self.text, re.IGNORECASE))
        has_organizational = bool(re.search(organizational, self.text))
        
        # Organizational is valid ONLY if no personal author format found
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
    print("SILVINA v0.5 - ASISTENTE EDITORIAL")
    print("="*70 + "\n")
    
    filepath = r"C:\Users\usuario\Desktop\Escudo cuantico_AB_25092025.docx"
    
    doc = Document(filepath)
    doc.load()
    
    # Generate report WITH LLM review (Session 1 testing)
    report = doc.generate_report(include_llm=True)
    print(report)
    
    # Save report to file
    report_filename = f"reporte_silvina_v05_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(report_filename, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"\n💾 Reporte guardado: {report_filename}")
    
    doc.close()
