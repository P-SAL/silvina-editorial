# Silvina Editorial Assistant v0.7

![Version](https://img.shields.io/badge/version-0.7-blue) ![Python](https://img.shields.io/badge/python-3.8+-green) ![License](https://img.shields.io/badge/license-MIT-orange)

AI-powered manuscript review for Spanish academic journals | EUMIC compliance • APA 7 validation • Professional reports • Two-tier quality analysis

---

## 📖 Overview

Silvina is an intelligent editorial assistant for **Revista Visión Conjunta** (Facultad Militar Conjunta - Universidad de la Defensa Nacional, Argentina). It automates academic manuscript review using deterministic validation and selective AI-powered analysis.

**Current Version:** v0.7 (February 2026)  
**Architecture:** Modular 3-layer design (Data Access → Business Logic → Presentation)  
**LLM Integration:** Ollama (llama3.2:3b)  
**Grammar Engine:** LanguageTool (Spanish)

---

## ✨ What's New in v0.7

### 🎨 Professional Word Reports
- **Institutional branding**: EUMIC logo in header (all pages)
- **Executive summary**: Key metrics table with publishability decision
- **Detailed sections**: Color-coded, properly formatted with tables
- **Error details**: Exact locations, context, and suggestions for grammar/APA errors
- **Page numbering**: "X de Y" format in footer
- **Professional styling**: Calibri 12pt, proper spacing, aligned layouts

### 🔍 Enhanced Analysis
- **Two-tier quality system**:
  - **Tier 1 (Deterministic)**: Grammar & spelling with LanguageTool
  - **Tier 2 (LLM)**: Semantic dimensions (clarity, coherence, argumentation, conclusions)
- **Grammar error details**: Shows context, exact location, and correction suggestions
- **APA 7 validation**: Detects conjunction errors (& vs y) with specific citations
- **Robust citation parser**: Correctly detects narrative citations (e.g., "Coleman (2023)")

### 🛠️ Technical Improvements
- **XML-based extraction**: Direct parsing from Word XML for accurate citation detection
- **Flexible LLM parser**: Handles format variations in Ollama output
- **Cleaner architecture**: Separated Data/Business/Presentation layers
- **Better error handling**: Graceful fallbacks, informative messages

---

## 🎯 Core Features

### 📚 Comprehensive Analysis Pipeline

**7-Step Process:**
1. **Document Reading** - Extracts paragraphs from `.docx` files
2. **Content Extraction** - Identifies title, authors, sections, word count
3. **Citation & Reference Parsing** - Extracts in-text citations and bibliography (XML-based)
4. **Article Classification** - Científico vs Divulgación (LLM-powered)
5. **Quality Analysis** - Two-tier: deterministic grammar + LLM semantics
6. **Structure Validation** - Verifies required sections per EUMIC standards
7. **Citation Matching** - Links citations to references, calculates match rate

### ⭐ Two-Tier Quality Analysis

**Tier 1 - Deterministic (LanguageTool):**
- Grammar and spelling errors
- Context and suggestions provided
- Score: 0-10 based on error count

**Tier 2 - Semantic (LLM):**
- **Claridad**: Is the argument comprehensible?
- **Coherencia**: Are ideas logically connected?
- **Argumentación**: Is there evidence supporting claims?
- **Conclusiones**: Do conclusions derive from content?
- Overall score: 0-10 (average of 4 dimensions)

### 📋 EUMIC Structure Validation

**Required Sections by Article Type:**

| Article Type | Required Sections |
|-------------|-------------------|
| **Científico** | Resumen, Introducción, Metodología, Resultados, Discusión, Conclusiones, Referencias |
| **Divulgación** | Resumen, Introducción, Desarrollo, Conclusiones, Referencias |
| **Opinión** | Introducción, Desarrollo, Conclusiones, Referencias |

### 📖 APA 7 (Spanish) Validation

**Checks:**
- Conjunction errors: `&` → `y` for Spanish
- Parenthetical format: `(Autor, 2020)`
- Narrative format: `Autor (2020)`
- Multiple authors: `Autor1 y Autor2 (2020)`
- Et al. format: `Autor et al. (2020)`

**Output:** Specific errors with location, explanation, and correction

### 🔗 Citation-Reference Matching

**Features:**
- Detects 13 types of citations (parenthetical, narrative, single/multi-author)
- Matches citations to bibliography entries
- Identifies unmatched citations
- Calculates match rate percentage
- Flags missing references

---

## 📊 Report Formats

### 📄 Professional Word Report (.docx)

**Sections:**
1. **Title Page**: Document name, institutional logo
2. **Executive Summary**: Key metrics table, publishability decision
3. **Document Information**: Title, author, word count, pages
4. **Classification**: Article type with confidence
5. **Quality Analysis**: 
   - Grammar (Tier 1) with error details
   - Semantic dimensions (Tier 2) with feedback
6. **APA 7 Validation**: Specific errors with corrections
7. **Structure Validation**: Missing sections (if any)
8. **Citations Analysis**: Match rate, unmatched citations
9. **Recommendations**: Prioritized action items

**Styling:**
- EUMIC logo in header (all pages)
- Color-coded status indicators (red/yellow/green)
- Tables for structured data
- Page numbers: "X de Y" format
- Professional fonts: Calibri 12pt

### 📊 JSON Export (.json)

Complete structured data including:
- Document metadata
- All analysis results
- Raw scores and feedback
- Citation/reference lists
- Recommendations

---

## 🛠️ Technology Stack

| Component | Technology |
|-----------|-----------|
| **Language** | Python 3.8+ |
| **Document Parsing** | python-docx (Word .docx) |
| **Grammar Check** | LanguageTool (Spanish) |
| **LLM** | Ollama (llama3.2:3b) |
| **Architecture** | Modular 3-layer design |
| **Output** | Word reports with formatting |

---

## 📦 Installation

### Prerequisites
- Python 3.8+
- Ollama installed and running
- 8GB+ RAM recommended

### Setup

```bash
# 1. Clone repository
git clone https://github.com/P-SAL/silvina-editorial.git
cd silvina-editorial/silvina_editorial_v07

# 2. Create virtual environment
python -m venv venv312
source venv312/bin/activate  # Windows: venv312\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Pull LLM model (one-time)
ollama pull llama3.2:3b

# 5. Add institutional logo (optional)
# Place logo as: assets/logo.jpg
```

---

## 🚀 Usage

### Interactive Mode
```bash
python main.py
```
System will prompt for document path.

### Command Line
```bash
python main.py "path/to/document.docx"
```

### Output Files
Generated in same directory as input:
- `document_analisis.docx` - Professional Word report
- `document_analisis.json` - Structured JSON data

### Example Console Output
```
================================================================================
   SILVINA EDITORIAL ASSISTANT v0.7
================================================================================

📄 Analizando documento: mi_articulo.docx

[1/7] 📖 Leyendo documento...
      ✓ Documento leído correctamente

[3/7] 📚 Analizando citas y referencias...
      ✓ 13 citas detectadas
      ✓ 17 referencias detectadas

      🔍 Validando formato APA 7...
      ⚠️  3 errores de formato APA 7 detectados

      🔍 Validando gramática y ortografía...
      ✓ Gramática: 8.5/10.0 - ⚠️ 2 errores menores detectados

[5/7] ⭐ Analizando calidad...
      ✓ Puntuación: 8.2/10.0
      ✓ Nivel: Bueno

✅ Análisis completado exitosamente
```

---

## 📁 Project Structure

```
silvina_editorial_v07/
├── assets/
│   └── logo.jpg                    # Institutional logo (EUMIC)
├── domain/
│   ├── models.py                   # Data models (DocumentContent, Citation)
│   └── enums.py                    # Enumerations (ArticleType, QualityLevel)
├── data_access/
│   ├── word_reader.py              # .docx file reader
│   ├── content_extractor.py        # Content extraction
│   ├── citation_parser.py          # XML-based citation parser
│   └── reference_parser.py         # Bibliography parser
├── business_logic/
│   ├── article_classifier.py       # LLM classification
│   ├── quality_analyzer.py         # Semantic analysis (Tier 2)
│   ├── gramatica_checker.py        # Grammar check (Tier 1)
│   ├── structure_validator.py      # EUMIC validation
│   └── citation_matcher.py         # Citation-reference matching
├── presentation/
│   ├── word_exporter.py            # Professional Word reports
│   ├── report_formatter.py         # Text formatting
│   └── config.py                   # Configuration
├── apa_validator.py                # APA 7 Spanish validator
├── eumic_verifier.py               # EUMIC standards checker
└── main.py                         # Entry point
```

---

## 🎯 Publishability Decision Criteria

| Status | Conditions |
|--------|-----------|
| ✅ **APTO** | Quality ≥7, Grammar ≥7, Structure valid, No APA errors |
| ⚠️ **REQUIERE REVISIÓN** | Quality 6-7, or 1-3 APA errors |
| ❌ **NO APTO** | Quality <5, incomplete structure, or >5 critical errors |

---

## 🔄 Version History

### v0.7 (February 2026) - Current
- ✨ Professional Word reports with EUMIC branding
- ✨ Two-tier quality analysis (Grammar + Semantics)
- ✨ Grammar error details with context and suggestions
- ✨ APA 7 validation with specific error locations
- 🔧 Robust citation parser (handles narrative formats)
- 🔧 Flexible LLM parser (handles output variations)
- 🎨 Improved report formatting (tables, colors, spacing)
- 📊 JSON export for data integration

### v0.6 (December 2025)
- Citation-reference integrity validation
- IMRyD structure detection
- Basic LLM quality analysis
- Text-only reports

### v0.5 and earlier
- Basic document analysis
- Simple character counting
- Minimal validation

---

## 🗺️ Roadmap

### v0.8 (Planned - Q1 2026)
- 🎨 **Gradio web interface** for non-technical users
- 📱 Drag-and-drop file upload
- 📊 Interactive visualization
- 💾 Batch processing

### v0.9 (Planned - Q2 2026)
- 📧 Email notifications
- 🔄 Version comparison
- 📈 Analytics dashboard
- 🌐 English support

### v1.0 (Planned - Q3 2026)
- 🏢 Production deployment at UNDEF
- 📚 Journal submission integration
- 👥 Multi-user authentication

---

## 🤝 Contributing

Currently in active development for Universidad de la Defensa Nacional, Argentina.

**Contact:** Pablo Salonio (P-SAL)  
**Email:** plsalonio@gmail.com  
**Repository:** https://github.com/P-SAL/silvina-editorial

---

## 📄 License

MIT License - See LICENSE file

---

## 🙏 Acknowledgments

- **Revista Visión Conjunta** - Editorial requirements and testing
- **Facultad Militar Conjunta** - Universidad de la Defensa Nacional
- **Ollama Team** - Local LLM infrastructure
- **LanguageTool** - Open-source grammar checking
- **Claude (Anthropic)** - Development assistance

---

**Last Updated:** February 8, 2026  
**Version:** 0.7  
**Status:** Production Ready 🚀
