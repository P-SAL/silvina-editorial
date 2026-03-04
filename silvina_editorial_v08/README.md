# Silvina Editorial Assistant v0.8

[![Version](https://img.shields.io/badge/version-v0.8-blue)](https://github.com/P-SAL/silvina-editorial)
[![Python](https://img.shields.io/badge/python-3.12-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Status](https://img.shields.io/badge/status-Active%20Development-yellow)](https://github.com/P-SAL/silvina-editorial)

**AI-powered manuscript review for Spanish academic journals** | EUMIC compliance • APA 7 validation • Modular architecture • LLM-powered quality analysis • Gradio web interface

---

## 📖 Overview

Silvina is an intelligent editorial assistant for **Revista Visión Conjunta** (Facultad Militar Conjunta - Universidad de la Defensa Nacional, Argentina). It automates academic manuscript review using **deterministic structural validation** and **selective AI-powered analysis**.

**Current Version:** v0.8 (Q1 2026)  
**Architecture:** Modular 4-layer design (Domain → Data Access → Business Logic → Presentation)  
**LLM Integration:** Ollama (llama3-gradient:8b-instruct-1048k-q4_K_M)  
**Interface:** Gradio web UI (new in v0.8)

---

## ✨ Key Features

### 🎨 **Gradio Web Interface (NEW in v0.8)**

- 📱 **Drag-and-drop** file upload — no technical knowledge required
- 📊 **Interactive result visualization** — color-coded scores and error summaries
- 💾 **One-click download** — Word report and JSON data
- 💬 **Expert feedback panel** — editorial staff can rate analysis accuracy
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
5. **⭐ Quality Analysis** - 5 dimensions: normativa, claridad, coherencia, argumentación, conclusiones
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

**5 Evaluation Dimensions:**
- **Normativa** - Orthographic and grammatical correctness
- **Claridad** - Writing clarity and readability
- **Coherencia** - Logical flow and consistency
- **Argumentación** - Strength of arguments and evidence
- **Conclusiones** - Quality and relevance of conclusions

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

**2 Output Formats:**
1. **📘 Word Report** (`.docx`) - Formatted document with color-coded sections
2. **📊 JSON Data** (`.json`) - Structured data for further processing

---

## 🛠️ Technology Stack

| Component | Technology |
|-----------|------------|
| **Language** | Python 3.12 |
| **Web Interface** | Gradio |
| **Document Parsing** | python-docx |
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
cd silvina-editorial/silvina_editorial_v08

# 2. Create virtual environment
python -m venv venv312
source venv312/Scripts/activate  # Windows Git Bash

# 3. Install dependencies
pip install python-docx ollama pywin32 gradio

# 4. Pull LLM model (one-time)
ollama pull llama3-gradient:8b-instruct-1048k-q4_K_M

# 5. Verify installation
python -c "import gradio; print('✅ Gradio ready')"
```

---

## 🚀 Usage

### Launch Web Interface (v0.8)
```bash
python gradio_app.py
```

Chrome will open automatically at `http://127.0.0.1:7861`

**Workflow:**
1. Drag and drop your `.docx` manuscript
2. Click **Analizar Documento**
3. Review results on screen
4. Download the **Word report** or **JSON data**
5. Submit your expert evaluation (optional)
6. Click **Cerrar Silvina** when done

### Command Line (legacy)
```bash
python main.py "path/to/document.docx"
```

---

## 📁 Project Structure
```
silvina_editorial_v08/
├── gradio_app.py                # NEW: Gradio web interface entry point
├── main.py                      # CLI entry point, orchestrates analysis
├── eumic_verifier.py            # EUMIC format compliance checker
├── apa_validator.py             # APA 7 citation format validator
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

### v0.8 (Q1 2026) - Current
- ✨ **NEW:** Gradio web interface for editorial staff
- ✨ **NEW:** Drag-and-drop file upload
- ✨ **NEW:** Interactive result visualization (scores, error counters, progress)
- ✨ **NEW:** Expert feedback panel for continuous improvement
- ✨ **NEW:** One-click Word/JSON report download
- ✨ **NEW:** Auto-launch in Chrome on startup
- 🔧 **FIXED:** Duplicate click handler bug (triple-fire on analyze button)

### v0.7 (January 2026)
- EUMIC format compliance verification system
- Grammar checker (Tier 1 deterministic)
- Enhanced LLM feedback parsing
- Accurate word/character counting via Windows COM automation
- APA 7 Spanish format validation

### v0.6 (December 2025)
- Citation-reference integrity validation
- IMRyD structure detection
- Two-tier analysis strategy

### v0.5 and earlier
- Basic character counting
- Simple LLM review
- Text-only output

---

## 🗺️ Roadmap

### v0.8 (Q1 2026) ✅ Current
- ✅ Gradio web interface
- ✅ Drag-and-drop file upload
- ✅ Interactive result visualization
- ⬜ Batch processing for multiple documents *(moved to v0.9)*

### v0.9 (Planned - Q2 2026)
- 💾 Batch processing for multiple documents
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

## 🐛 Known Issues

### v0.8
- **Citation matching** may fail for non-standard citation formats
- **EUMIC verification** requires Windows (COM dependency)
- **LLM analysis** quality depends on Ollama model availability
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
**Version:** 0.8  
**Status:** Active Development 🚀