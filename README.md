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
5. **⭐ Quality Analysis** - 6 dimensions: clarity, coherence, argumentation, methodology, conclusions, format
6. **📋 Structure Validation** - Verifies required sections per article type (EUMIC standards)
7. **🔗 Citation Matching** - Links citations to references, identifies orphaned entries

---

### 🎯 **Classification System**

**Article Types (EUMIC-compliant):**
- **Artículo Científico** - Research with IMRyD structure, 3000-6000 words
- **Artículo de Divulgación** - Literature review, flexible structure
- **Artículo de Opinión** - Opinion/analysis piece
- **Artículo Corto** - Brief communication, 1000-2000 words

**Classification Method:** LLM-based with confidence scoring (0-100%)

---

### ⭐ **Quality Analysis**

**6 Evaluation Dimensions:**
- **Claridad** - Writing clarity and readability
- **Coherencia** - Logical flow and consistency
- **Argumentación** - Strength of arguments and evidence
- **Metodología** - Research methods rigor (if applicable)
- **Conclusiones** - Quality and relevance of conclusions
- **Formato** - Compliance with EUMIC/APA 7 standards

**Output:** Overall score (0-10), quality level (Excellent/Good/Acceptable/Needs Improvement/Poor)

---

### 📋 **Structure Validation**

**Required Sections by Type:**

| Article Type | Required Sections |
|--------------|-------------------|
| Científico | Resumen, Introducción, Metodología, Resultados, Discusión, Conclusiones, Referencias |
| Divulgación | Resumen, Introducción, Desarrollo, Conclusiones, Referencias |
| Opinión | Introducción, Desarrollo, Conclusiones, Referencias |

**Validation:** Detects missing sections, provides actionable feedback

---

### 🔗 **Citation-Reference Matching**

**Features:**
- Extracts APA 7 Spanish citations: `(Autor, 2020)`, `Autor (2020)`
- Handles `et al.`, organizational authors, multiple citations
- Matches citations to bibliography entries
- Identifies unmatched citations (red flag for publication)
- Calculates match rate percentage

---

### 📊 **Multi-Format Reports**

**3 Output Formats:**
1. **📄 Text Report** (`.txt`) - Complete analysis in plain text
2. **📘 Word Report** (`.docx`) - Formatted document with color-coded sections
3. **📊 JSON Data** (`.json`) - Structured data for further processing

**Report Includes:**
- ✅/❌ **Publishability decision** (NEW in v0.7)
- Document metadata (title, authors, word count)
- Classification results with confidence
- Quality scores by dimension
- Structure validation status
- Citation analysis with match rate
- Prioritized recommendations (High/Medium/Low)

---

## 🛠️ Technology Stack

| Component | Technology |
|-----------|------------|
| **Language** | Python 3.12 |
| **Document Parsing** | python-docx (Word .docx files) |
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

### Setup
```bash
# 1. Clone repository
git clone https://github.com/P-SAL/silvina-editorial.git
cd silvina-editorial/silvina_editorial_v07

# 2. Create virtual environment
python -m venv venv312
source venv312/bin/activate  # Windows: venv312\Scripts\activate

# 3. Install dependencies
pip install python-docx ollama

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
      ✓ 45 párrafos leídos

[2/7] 🔍 Extrayendo contenido estructurado...
      ✓ Título: El impacto de la inteligencia artificial...
      ✓ Total de palabras: 4,523

[3/7] 📚 Analizando citas y referencias...
      ✓ 23 citas encontradas
      ✓ 25 referencias encontradas

[4/7] 🏷️  Clasificando tipo de artículo...
      ✓ Categoría: Artículo Científico
      ✓ Confianza: 85.0%

[5/7] ⭐ Analizando calidad...
      ✓ Puntuación: 8.2/10.0
      ✓ Nivel: Bueno

[6/7] 📋 Validando estructura...
      ✓ VÁLIDA

[7/7] 🔗 Relacionando citas con referencias...
      ✓ Tasa de coincidencia: 95.7%

================================================================================
✅ Análisis completado exitosamente
================================================================================
```

---

## 📁 Project Structure
```
silvina_editorial_v07/
├── main.py                      # Entry point, orchestrates analysis
├── domain/
│   ├── models.py                # Core data models (DocumentContent, Citation, etc.)
│   └── enums.py                 # Enumerations (ClassificationCategory, QualityLevel)
├── data_access/
│   ├── word_reader.py           # Reads .docx files
│   ├── content_extractor.py     # Extracts title, authors, sections
│   ├── citation_parser.py       # Parses APA citations
│   └── reference_parser.py      # Parses bibliography
├── business_logic/
│   ├── article_classifier.py    # LLM-based classification
│   ├── quality_analyzer.py      # Multi-dimension quality scoring
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
- Character count requirements (Científico: 30,000-50,000)
- Required sections per article type
- Spanish APA 7 citation format

### APA 7 (Spanish)
- Author format: `Apellido, N.`
- Conjunction: `y` (not `&`)
- Date format: `(2020)`, `(2020a)`, `(2020, 15 de enero)`
- Page references: `(p. 23)`, `(pp. 45-67)`

---

## 🔄 Version History

### v0.7 (January 2026) - Current
- ✨ **NEW:** Modular 4-layer architecture
- ✨ **NEW:** Publishability decision in reports
- ✨ **NEW:** JSON export for data integration
- 🚀 Faster processing with batched LLM calls
- 🐛 Fixed citation parser for Spanish APA
- 📊 Enhanced Word report formatting

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

### Current Status
- ✅ Citation extraction: Working
- ✅ Reference parsing: Working
- ✅ Document classification: 85%+ accuracy
- ✅ Structure validation: Functional
- ⚠️ Citation matching: Needs refinement
- ⚠️ Quality analysis: Testing with diverse documents

### Test Documents
Located in `test_documents/` (not in repo for privacy):
- Scientific articles with IMRyD structure
- Review articles with flexible structure
- Documents with citation issues
- Documents with structural gaps

---

## 🤝 Contributing

This project is currently in active development for internal use at Universidad de la Defensa Nacional, Argentina.

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
- 📧 Contact: [via GitHub](https://github.com/P-SAL)

---

**Last Updated:** January 17, 2026  
**Version:** 0.7  
**Status:** Active Development 🚀