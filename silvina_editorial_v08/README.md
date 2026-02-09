# Silvina Editorial Assistant v0.7

[![Version](https://img.shields.io/badge/version-v0.7-blue)](https://github.com/P-SAL/silvina-editorial)
[![Python](https://img.shields.io/badge/python-3.12-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Status](https://img.shields.io/badge/status-Active%20Development-yellow)](https://github.com/P-SAL/silvina-editorial)

**AI-powered manuscript review for Spanish academic journals** | EUMIC compliance • APA 7 validation • Modular architecture • LLM-powered quality analysis

---

## 📖 Overview

Silvina is an intelligent editorial assistant for **Revista Visión Conjunta** (Facultad Militar Conjunta - Universidad de la Defensa Nacional, Argentina). It automates academic manuscript review using **deterministic structural validation** and **selective AI-powered analysis**.

**Current Version:** v0.7 (January 2026)  
**Architecture:** Modular 4-layer design (Domain → Data Access → Business Logic → Presentation)  
**LLM Integration:** Ollama (llama3-gradient:8b-instruct-1048k-q4_K_M)

---

## ✨ Key Features

### 🏗️ **Modular Architecture (NEW in v0.7)**

**4-Layer Design:**
```
domain/           # Core models & enums (DocumentContent, Citation, Reference)
data_access/      # Document parsing (Word reader, citation/reference extraction)
business_logic/   # Analysis engines (classifier, quality analyzer, structure validator)
presentation/     # Output formatting (text/Word/JSON reports)
```

**Benefits:**
- Clean separation of concerns
- Easy testing and maintenance
- Scalable for future features
- Type-safe with dataclasses

---

### 📚 **Document Analysis Pipeline**

**7-Step Process:**

1. **📖 Document Reading** - Extracts paragraphs from .docx files
2. **🔍 Content Extraction** - Identifies title, authors, sections, word count
3. **📚 Citation & Reference Parsing** - Extracts in-text citations and bibliography
4. **🏷️ Article Classification** - Científico vs Divulgación (LLM-powered)
5. **⭐ Quality Analysis** - 5 dimensions: normativa, claridad, coherencia, argumentación, conclusiones
6. **📋 Structure Validation** - Verifies required sections per article type (EUMIC standards)
7. **🔗 Citation Matching** - Links citations to references, identifies orphaned entries

---

### 🆕 **EUMIC Compliance Verification (NEW in v0.7)**

**Automated Format Validation:**
- ✅ Document format (.docx)
- ✅ Margins (2.5 cm standard)
- ✅ Font (Times New Roman/Arial, 12pt)
- ✅ Text alignment (justified)
- ✅ Figure numbering and captions
- ✅ Table formatting and titles
- ✅ Formula presentation
- ✅ Abstract length (150-250 words)
- ✅ Keywords (3-5 required)

**Clean List Output:**
```
📋 VERIFICACIÓN EUMIC:
================================================================================

🔴 CRÍTICO (2):
   1. [Resumen y Palabras Clave] Falta sección de Resumen/Abstract
      → El documento debe incluir un resumen de 150-250 palabras
   2. [Resumen y Palabras Clave] Faltan palabras clave
      → Se requieren 3-5 palabras clave relevantes al contenido

🟡 ADVERTENCIAS (1):
   1. [Formato General] Margen superior no cumple estándar EUMIC
      → Requerido: 2.5 cm, Actual: 3.0 cm

================================================================================
```

---

### 🎯 **Classification System**

**Article Types (EUMIC-compliant):**
- **Artículo Científico** - Research with IMRyD structure, 3000-6000 words
- **Artículo de Divulgación** - Literature review, flexible structure
- **Artículo de Opinión** - Opinion/analysis piece
- **Artículo Corto** - Brief communication, 1000-2000 words

**Classification Method:** LLM-based with confidence scoring (0-100%)

---

### ⭐ **Quality Analysis (IMPROVED in v0.7)**

**5 Evaluation Dimensions:**
- **Normativa** - Orthographic and grammatical correctness
- **Claridad** - Writing clarity and readability
- **Coherencia** - Logical flow and consistency
- **Argumentación** - Strength of arguments and evidence
- **Conclusiones** - Quality and relevance of conclusions

**Improvements:**
- ✅ **Enhanced LLM feedback parsing** - Now correctly extracts all dimension feedback
- ✅ **Increased context window** - 8,000 characters (up from 3,500) for better analysis
- ✅ **Removed hardcoded scores** - Uses actual LLM evaluations
- ✅ **Cleaner output** - Removes unwanted "Nota:" sections from LLM responses

**Output:** Overall score (0-10), quality level (Excelente/Bueno/Aceptable/Necesita Mejora/Deficiente)

---

### 📋 **Structure Validation**

**Required Sections by Type:**

| Article Type | Required Sections |
|--------------|-------------------|
| Científico | Resumen, Introducción, Metodología, Resultados, Discusión, Conclusiones, Referencias |
| Divulgación | Resumen, Introducción, Desarrollo, Conclusiones, Referencias |
| Opinión | Introducción, Argumentación, Conclusiones, Referencias |

**Validation:** Detects missing sections, provides actionable feedback

---

### 🔗 **Citation-Reference Matching (IMPROVED in v0.7)**

**Features:**
- Extracts APA 7 Spanish citations: `(Autor, 2020)`, `Autor (2020)`
- Handles `et al.`, organizational authors, multiple citations
- Detects Markdown-style footnotes: `[^1]`, `[^2]`
- Matches citations to bibliography entries
- Identifies unmatched citations (red flag for publication)
- Calculates match rate percentage
- **NEW:** Clean output without unnecessary paragraph count warnings

---

### 📊 **Multi-Format Reports**

**3 Output Formats:**
1. **📄 Text Report** (`.txt`) - Complete analysis in plain text
2. **📘 Word Report** (`.docx`) - Formatted document with color-coded sections
3. **📊 JSON Data** (`.json`) - Structured data for further processing

**Report Includes:**
- Document metadata (title, authors, word count)
- Classification results with confidence
- Quality scores by dimension with detailed feedback
- EUMIC compliance verification (NEW in v0.7)
- Structure validation status
- Citation analysis with match rate
- Prioritized recommendations (Alta/Media/Baja)

---

## 🛠️ Technology Stack

| Component | Technology |
|-----------|------------|
| **Language** | Python 3.12 |
| **Document Parsing** | python-docx (Word .docx files) |
| **Word Automation** | win32com (Windows COM for accurate counts) |
| **LLM Integration** | Ollama (local inference) |
| **Data Models** | Dataclasses with type hints |
| **Architecture** | Modular 4-layer design |
| **Output** | python-docx for Word reports |

---

## 📦 Installation

### Prerequisites
- Python 3.12+
- [Ollama](https://ollama.ai/) installed and running
- 8GB+ RAM (16GB recommended)
- Windows (for accurate Word document statistics via COM automation)

### Setup
```bash
# 1. Clone repository
git clone https://github.com/P-SAL/silvina-editorial.git
cd silvina-editorial/silvina_editorial_v07

# 2. Create virtual environment
python -m venv venv312
source venv312/bin/activate  # Windows: venv312\Scripts\activate

# 3. Install dependencies
pip install python-docx ollama pywin32

# 4. Pull LLM model (one-time)
ollama pull llama3-gradient:8b-instruct-1048k-q4_K_M

# 5. Verify installation
python -c "import ollama; print('✅ Ollama ready')"
```

---

## 🚀 Usage

### Basic Analysis
```bash
python main.py "path/to/document.docx"
```

**Example:**
```bash
python main.py "C:\Users\usuario\Desktop\mi_articulo.docx"
```

**Interactive Mode:**
```bash
python main.py
# Will prompt: Ingrese la ruta del documento (.docx):
```

**Output:** 3 files in the same directory as input:
- `mi_articulo_analisis.txt`
- `mi_articulo_analisis.docx`
- `mi_articulo_analisis.json`

---

### Console Output Example
```
================================================================================
   SILVINA EDITORIAL ASSISTANT v0.7
   Asistente de Análisis Editorial para Documentos Académicos
================================================================================

📄 Analizando documento: mi_articulo.docx
================================================================================

[1/7] 📖 Leyendo documento...
      ✓ Documento leído correctamente

[2/7] 🔍 Extrayendo contenido estructurado...
      ✓ Conteos precisos obtenidos desde Word
      ✓ Título: El impacto de la inteligencia artificial...
      ✓ Autor: Dr. Juan Pérez
      ✓ Total de palabras: 4,523
      ✓ Total de caracteres: 35,421

[3/7] 📚 Analizando citas y referencias...
      ✓ Total: 23 citas | 25 referencias

[4/7] 🏷️  Clasificando tipo de artículo...
      ✓ Categoría: Científico
      ✓ Confianza: 85.0%

[5/7] ⭐ Analizando calidad...
      ⏳ Analizando con Ollama...
      Generando análisis: 485 palabras
      ✅ Análisis completado
      ✓ Puntuación: 8.2/10.0
      ✓ Nivel: Bueno

[6/7] 📋 Validando estructura...
      ✓ VÁLIDA

[7/7] 🔗 Relacionando citas con referencias...
      ✓ Tasa de coincidencia: 92.0%

================================================================================
✅ Análisis completado exitosamente

📋 VERIFICACIÓN EUMIC:
================================================================================

🟡 ADVERTENCIAS (2):
   1. [Formato General] Margen izquierdo no cumple estándar EUMIC
      → Requerido: 2.5 cm, Actual: 3.00 cm
   2. [Figuras] Figuras sin título descriptivo
      → 3 imágenes detectadas, 2 títulos encontrados

================================================================================
```

---

## 📁 Project Structure
```
silvina_editorial_v07/
├── main.py                      # Entry point, orchestrates analysis
├── eumic_verifier.py            # NEW: EUMIC format compliance checker
├── domain/
│   ├── models.py                # Core data models (DocumentContent, Citation, etc.)
│   └── enums.py                 # Enumerations (ClassificationCategory, QualityLevel)
├── data_access/
│   ├── word_reader.py           # Reads .docx files
│   ├── word_counter.py          # Accurate counts via COM automation
│   ├── content_extractor.py     # Extracts title, authors, sections
│   ├── citation_parser.py       # Parses APA citations (IMPROVED)
│   └── reference_parser.py      # Parses bibliography
├── business_logic/
│   ├── article_classifier.py    # LLM-based classification
│   ├── quality_analyzer.py      # Multi-dimension quality scoring (IMPROVED)
│   ├── structure_validator.py   # EUMIC structure compliance
│   └── citation_matcher.py      # Citation-reference linking
└── presentation/
    ├── report_formatter.py      # Text report generation
    ├── word_exporter.py         # Word document export
    └── config.py                # Configuration settings
```

---

## 🎯 Validation Standards

### EUMIC Guidelines
- Document format requirements (.docx)
- Margin specifications (2.5 cm)
- Font standards (Times New Roman/Arial, 12pt)
- Text alignment (justified)
- Figure and table formatting
- Abstract length (150-250 words)
- Keywords requirement (3-5)
- Required sections per article type

### APA 7 (Spanish)
- Author format: `Apellido, N.`
- Conjunction: `y` (not `&`)
- Date format: `(2020)`, `(2020a)`, `(2020, 15 de enero)`
- Page references: `(p. 23)`, `(pp. 45-67)`

---

## 🔄 Version History

### v0.7 (January 2026) - Current
- ✨ **NEW:** EUMIC format compliance verification system
- ✨ **NEW:** Clean list output for validation errors
- 🔧 **FIXED:** LLM feedback parser now correctly extracts all dimension analysis
- 🔧 **FIXED:** Removed hardcoded quality scores - uses actual LLM evaluations
- 🔧 **FIXED:** Eliminated duplicate output messages
- 🔧 **FIXED:** Removed unnecessary paragraph count warnings
- 🚀 **IMPROVED:** Increased LLM context window from 3,500 to 8,000 characters
- 🚀 **IMPROVED:** Enhanced citation parser for Spanish APA format
- 🚀 **IMPROVED:** Accurate word/character counting via Windows COM automation
- 📊 **IMPROVED:** Cleaner console output with better formatting

### v0.6 (December 2025)
- Citation-reference integrity validation
- IMRyD structure detection
- Organizational author support
- Two-tier analysis strategy

### v0.5 and earlier
- Basic character counting
- Simple LLM review
- Text-only output

---

## 🗺️ Roadmap

### v0.8 (Planned - Q1 2026)
- 🎨 **Gradio web interface** for user-friendly access
- 📱 Drag-and-drop file upload
- 📊 Interactive result visualization
- 💾 Batch processing for multiple documents

### v0.9 (Planned - Q2 2026)
- 📧 Email integration for automatic notifications
- 🔄 Version comparison (track revisions)
- 📈 Analytics dashboard for editorial team
- 🌐 Multi-language support (English, Portuguese)

### v1.0 (Planned - Q3 2026)
- 🏢 Production deployment at Universidad de la Defensa
- 📚 Integration with journal submission system
- 👥 Multi-user authentication
- 📊 Editorial workflow management

---

## 🧪 Testing

### Current Status (v0.7)
- ✅ Citation extraction: Working (improved regex)
- ✅ Reference parsing: Working
- ✅ Document classification: 85%+ accuracy
- ✅ Structure validation: Functional
- ✅ EUMIC compliance: Fully implemented
- ✅ Quality analysis: LLM feedback properly extracted
- ✅ Word counting: Accurate via COM automation
- ⚠️ Citation matching: Needs refinement for edge cases

### Test Documents
Located in `test_documents/` (not in repo for privacy):
- Scientific articles with IMRyD structure
- Review articles with flexible structure
- Documents with citation issues
- Documents with structural gaps
- Documents with EUMIC format violations

---

## 🐛 Known Issues

### v0.7
- **Citation matching** may fail for non-standard citation formats
- **EUMIC verification** requires document to be opened in Word (COM dependency)
- **LLM analysis** quality depends on Ollama model availability and performance
- **Windows-only** COM automation for accurate statistics (falls back to python-docx on other platforms)

---

## 🤝 Contributing

This project is currently in active development for internal use at Facultad Militar Conjunta, Universidad de la Defensa Nacional, Argentina.

**Contact:** Pablo Salonio (P-SAL) - plsalonio@gmail.com  
**Repository:** https://github.com/P-SAL/silvina-editorial

---

## 📄 License

MIT License - See [LICENSE](LICENSE) file for details

---

## 🙏 Acknowledgments

- **Revista Visión Conjunta** - Editorial team for requirements and testing
- **Facultad Militar Conjunta** - Universidad de la Defensa Nacional
- **Ollama Team** - Local LLM infrastructure
- **Claude (Anthropic)** - Development assistance and code review

---

## 📞 Support

For issues, questions, or suggestions:
- 🐛 [Open an issue](https://github.com/P-SAL/silvina-editorial/issues)
- 📧 Contact: plsalonio@gmail.com

---

**Last Updated:** January 29, 2026  
**Version:** 0.7  
**Status:** Active Development 🚀
