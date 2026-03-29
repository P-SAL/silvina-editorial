# Silvina Editorial Assistant v0.9

[![Version](https://img.shields.io/badge/version-v0.9-blue)](https://github.com/P-SAL/silvina-editorial)
[![Python](https://img.shields.io/badge/python-3.12-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Status](https://img.shields.io/badge/status-Active%20Development-yellow)](https://github.com/P-SAL/silvina-editorial)

**AI-powered manuscript review for Spanish academic journals** | EUMIC compliance • APA 7 validation • Modular architecture • LLM-powered quality analysis • Gradio web interface

---

## 📖 Overview

Silvina is an intelligent editorial assistant for **Revista Visión Conjunta** (Facultad Militar Conjunta - Universidad de la Defensa Nacional, Argentina). It automates academic manuscript review using **deterministic structural validation** and **selective AI-powered analysis**.

**Current Version:** v0.9 (Q2 2026)  
**Based on:** v0.8 (stable, deployed to editorial team)  
**Architecture:** Modular 4-layer design (Domain → Data Access → Business Logic → Presentation)  
**LLM Integration:** Ollama (llama3-gradient:8b-instruct-1048k-q4_K_M)  
**Interface:** Gradio web UI  
**Output Location:** `Documents\Silvina\reports\` (Word report, JSON data, feedback file)

---

## ✨ Key Features

### 🎨 **Gradio Web Interface**

- 📱 **Drag-and-drop** file upload — no technical knowledge required
- 📊 **Interactive result visualization** — color-coded scores and error summaries
- 💾 **One-click download** — Word report and JSON data
- 💬 **Structured expert feedback panel** — 8-field evaluation form capturing classification accuracy, quality score fairness, grammar false positives, structure validation, citation detection, weakest section, and publication recommendation
- 🔴 **Clean shutdown button** — closes the server safely from the browser
- 🚀 **Auto-launches in Chrome** on startup

### 🏗️ **Modular Architecture**

**4-Layer Design:**
```
domain/           # Core models & enums (DocumentContent, Citation, Reference)
data_access/      # Document parsing (Word reader, citation/reference extraction)
business_logic/   # Analysis engines (classifier, quality analyzer, structure validator)
presentation/     # Output formatting (Word/JSON reports)
```

---

### 📚 **Document Analysis Pipeline**

**7-Step Process:**

1. **📖 Document Reading** - Extracts paragraphs from .docx files
2. **🔍 Content Extraction** - Identifies title, authors, sections, word count
3. **📚 Citation & Reference Parsing** - Extracts in-text citations and bibliography
4. **🏷️ Article Classification** - Científico vs Divulgación (LLM-powered)
5. **⭐ Quality Analysis** - 4 semantic dimensions: claridad, coherencia, argumentación, conclusiones
6. **📋 Structure Validation** - Verifies required sections per article type (EUMIC standards)
7. **🔗 Citation Matching** - Links citations to references, identifies orphaned entries

---

### 🆕 **EUMIC Compliance Verification**

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

**4 Semantic Evaluation Dimensions:**
- **Claridad** - Writing clarity and readability
- **Coherencia** - Logical flow and consistency
- **Argumentación** - Strength of arguments and evidence
- **Conclusiones** - Quality and relevance of conclusions

**Architecture:** Two-call LLM approach (Call 1: Claridad + Coherencia / Call 2: Argumentación + Conclusiones) with split-based parser and explicit conclusion section detection.

**Output:** Overall score (0-10), quality level (Excelente/Bueno/Aceptable/Necesita Mejora/Deficiente)

---

### 📋 **Structure Validation**

**Required Sections by Type:**

| Article Type | Required Sections |
|--------------|-------------------|
| Científico | Resumen, Introducción, Metodología, Resultados, Discusión, Conclusiones, Referencias |
| Divulgación | Resumen, Introducción, Desarrollo, Conclusiones, Referencias |
| Opinión | Introducción, Argumentación, Conclusiones, Referencias |

---

### 📊 **Multi-Format Reports**

**3 Output Files** saved automatically to `C:\Users\[user]\Documents\Silvina\reports\`:
1. **📘 Word Report** (`_analisis.docx`) - Formatted document with color-coded sections
2. **📊 JSON Data** (`_analisis.json`) - Structured data for further processing
3. **💬 Feedback File** (`_feedback.json`) - Expert evaluation submitted via Gradio

---

### 🔄 **Feedback Processing Pipeline**

Collect `_feedback.json` files from editorial team → run `process_feedback.py` → get data-driven development priorities for next version.

```bash
python process_feedback.py --folder feedback_received/
```

**Output:**
- `feedback_summary_YYYYMMDD.md` — ranked issues by frequency for editorial review
- `v09_dev_prompt_YYYYMMDD.md` — structured development prompt for next session

---

## 🛠️ Technology Stack

| Component | Technology |
|-----------|------------|
| **Language** | Python 3.12 |
| **Web Interface** | Gradio |
| **Document Parsing** | python-docx |
| **Word Automation** | win32com (Windows COM for accurate counts) |
| **LLM Integration** | Ollama (local inference) |
| **Grammar Checking** | LanguageTool |
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
cd silvina-editorial/silvina_editorial_v09

# 2. Create virtual environment
python -m venv venv312
source venv312/Scripts/activate  # Windows Git Bash

# 3. Install dependencies
pip install python-docx ollama pywin32 gradio language-tool-python

# 4. Pull LLM model (one-time)
ollama pull llama3-gradient:8b-instruct-1048k-q4_K_M

# 5. Verify installation
python -c "import gradio; print('Gradio ready')"
```

---

## 🚀 Usage

### Launch Web Interface
```bash
python gradio_app.py
```

Chrome will open automatically at `http://127.0.0.1:7861`

**Workflow:**
1. Drag and drop your `.docx` manuscript
2. Click **Analizar Documento**
3. Review results on screen
4. Download the **Word report** or **JSON data** (also saved automatically to `Documents\Silvina\reports\`)
5. Complete the **expert evaluation form** and click **Enviar Evaluación**
6. Click **Cerrar Silvina** when done

### Process Feedback (Coordinator)
```bash
# 1. Collect _feedback.json files from editorial team into feedback_received/
# 2. Run processing script
python process_feedback.py --folder feedback_received/
# 3. Review feedback_summary_YYYYMMDD.md
# 4. Use dev prompt for next development session
```

### Command Line (legacy)
```bash
python main.py "path/to/document.docx"
```

---

## 📁 Project Structure
```
silvina_editorial_v09/
├── gradio_app.py                # Gradio web interface entry point
├── main.py                      # CLI entry point
├── process_feedback.py          # Feedback processing pipeline
├── eumic_verifier.py            # EUMIC format compliance checker
├── apa_validator.py             # APA 7 citation format validator
├── feedback_received/           # Drop feedback JSONs here for processing
├── assets/
│   └── SILVINA V08.png          # Logo
├── domain/
│   ├── models.py                # Core data models
│   └── enums.py                 # Enumerations
├── data_access/
│   ├── word_reader.py
│   ├── content_extractor.py
│   ├── citation_parser.py
│   └── reference_parser.py
├── business_logic/
│   ├── article_classifier.py
│   ├── quality_analyzer.py
│   ├── gramatica_checker.py
│   ├── structure_validator.py
│   └── citation_matcher.py
└── presentation/
    ├── word_exporter.py
    └── config.py
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

### v0.9 (Q2 2026) - Current
- 🔧 **FIXED:** Grammar checker false positives — filter `rule_issue_type == misspelling` eliminates proper noun, surname, and military acronym false positives
- 🔧 **FIXED:** Title extraction — author name no longer included in title
- 🔧 **FIXED:** Multi-author detection — `AUTORES` pattern now correctly collects authors listed on separate lines

### v0.8 (Q1 2026)
- Gradio web interface for editorial staff
- Structured expert feedback panel (8 evaluation fields)
- All reports save to `Documents\Silvina\reports\`
- Feedback JSON saved alongside analysis reports
- `process_feedback.py` — automated feedback processing pipeline
- Two-call LLM architecture for quality analysis
- Split-based parser replacing fragile regex
- Explicit conclusion section detection

### v0.7 (January 2026)
- EUMIC format compliance verification system
- Grammar checker (Tier 1 deterministic)
- Accurate word/character counting via Windows COM automation
- APA 7 Spanish format validation

### v0.6 (December 2025)
- Citation-reference integrity validation
- IMRyD structure detection
- Two-tier analysis strategy

---

## 🗺️ Roadmap

### v0.9 (Q2 2026) ✅ Current
- ✅ Grammar false positive fix
- ✅ Title extraction fix
- ✅ Multi-author detection fix
- ⬜ Batch processing for multiple documents
- ⬜ Inline `Resumen:` detection improvement
- ⬜ Alternative reference headings (`Fuentes bibliográficas consultadas`)
- ⬜ Footnote citation detection and reporting
- ⬜ Security measures (file validation, authentication, rate limiting)
- ⬜ Deployment preparation for institutional web server
- ⬜ Analytics dashboard for editorial team

### v1.0 (Planned - Q3 2026)
- 🏢 Production deployment at Universidad de la Defensa
- 🌐 Accessible via institutional webpage
- 👥 Multi-user authentication
- 📊 Editorial workflow management

### v2.0 (Future)
- 🧠 Editorial memory — learns from accumulated feedback (RAG-based)
- 🤖 Agentic workflows

---

## 🐛 Known Issues

### v0.9
- **Structure validator** misses inline `Resumen:` format
- **Citation parser** misses `Fuente:` footnote format
- **Reference parser** misses `Fuentes bibliográficas consultadas` heading
- **Windows-only** COM automation for accurate statistics (falls back to python-docx on other platforms)

---

## 🤝 Contributing

This project is in active development for internal use at Facultad Militar Conjunta, Universidad de la Defensa Nacional, Argentina.

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

**Last Updated:** March 2026  
**Version:** 0.9  
**Status:** Active Development 🚀
