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
4. **🏷️ Article Classification** - Científico / Divulgación / Opinión (hybrid signal system)
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

### 🎯 **Classification System — Hybrid Signal Approach**

v0.9 implements a hybrid signal system combining deterministic and LLM-based signals, aligned with EUMIC norms and social science article conventions.

**Signal evaluation:**

| Signal | Type | Criterion | Role |
|--------|------|-----------|------|
| S1 — IMRyD override | Deterministic | Complete IMRyD structure → immediate CIENTÍFICO (0.95) | Override |
| S2a — Reference count | Deterministic | ≥ 15 references (EUMIC minimum) | Supporting |
| S2b — Reference recency | Deterministic | ≥ 50% of references within last 4 years | Supporting |
| S3 — Methodological vocab | Deterministic | ≥ 4 terms AND ≥ 1 hard term | Tiebreaker |
| S4 — Research intent | LLM (targeted) | Explicit research intent via linguistic/structural patterns | **Primary** |
| S5 — Conclusive contribution | LLM (targeted) | Conclusions that exteriorize a contribution via systematic process | **Primary** |

**S4 detects any of:**
- Verbs of intent: examinar, analizar, identificar, determinar, explorar, comprender, evaluar, investigar, revisar, sintetizar
- Scope markers: "el presente estudio", "esta investigación", "la presente revisión"
- Problem markers: "el problema central", "el objetivo es", "la pregunta que guía"
- Research questions or hypotheses — single or multiple, numbered or sequential

**S5 detects any of:**
- Findings from systematic, replicable or verifiable process
- Framework, model, taxonomy or classification proposed
- Evidence-based recommendations derived from analysis
- Knowledge gap identified and addressed
- Synthesis beyond description integrating multiple sources

**Classification rule:**

| Condition | Result | Confidence |
|-----------|--------|------------|
| S4 + S5 + (S2a OR S2b) | CIENTÍFICO | 0.85 |
| S4 + S5 (no S2) | DIVULGACIÓN | 0.75 |
| S2a + S2b + S3 (no S4/S5) | DIVULGACIÓN | 0.70 |
| S4 OR S5 alone | DIVULGACIÓN | 0.65 |
| No signals | OPINIÓN | 0.65 |

**Design principles:**
- S4 and S5 are the primary discriminators — necessary and sufficient for CIENTÍFICO when combined with reference support
- S2a/S2b provide reference quality evidence — supporting but not sufficient alone
- S3 acts as tiebreaker only in ambiguous cases
- LLM prompts are precision-engineered to detect specific linguistic and structural forms, not judge quality

**Article Types (EUMIC-compliant):**
- **Artículo Científico** - Research intent + conclusive contribution + reference support
- **Artículo de Divulgación** - Academic synthesis, flexible structure, partial scientific signals
- **Artículo de Opinión** - Argumentative text, no scientific signals

---

### ⭐ **Quality Analysis**

**4 Semantic Evaluation Dimensions:**
- **Claridad** - Writing clarity and readability
- **Coherencia** - Logical flow and consistency
- **Argumentación** - Strength of arguments and evidence
- **Conclusiones** - Quality and relevance of conclusions

**Architecture:** Two-call LLM approach (Call 1: Claridad + Coherencia / Call 2: Argumentación + Conclusiones). Parser handles both numbered (`**1. Argumentación**`) and unnumbered (`**Argumentación**`) LLM response formats for robustness.

**Text sampling:** First 3500 chars (introduction/research questions) + last 2500 chars (conclusions), with bibliography section excluded via paragraph-level detection.

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
- 8GB+ RAM (32GB recommended for optimal LLM performance)
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

### Command Line
```bash
python main.py "path/to/document.docx"
```

---

## 📁 Project Structure
```
silvina_editorial_v09/
├── gradio_app.py
├── main.py
├── process_feedback.py
├── eumic_verifier.py
├── apa_validator.py
├── assets/
│   └── SILVINA V08.png
├── domain/
│   ├── models.py
│   └── enums.py
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
│   ├── citation_matcher.py
│   └── vocab/
│       ├── __init__.py
│       └── methodological_terms.py
└── presentation/
    ├── word_exporter.py
    └── config.py
```

---

## 🎯 Validation Standards

### EUMIC Guidelines
- Científico: 15-30 references recommended (APA 7, mandatory)
- Abstract: 150-250 words
- Keywords: 3-5
- Margins: 2.5 cm
- Font: Times New Roman 12pt

### APA 7 (Spanish)
- Author: `Apellido, N.`
- Conjunction: `y` (not `&`)
- Date: `(2020)`, `(2020a)`

---

## 🔄 Version History

### v0.9 (Q2 2026) - Current

**Classification System — Option B redesign:**
- ✨ **NEW:** S4 prompt redesigned — detects explicit research intent via linguistic/structural patterns (verbs of intent, scope markers, problem markers, sequential research questions)
- ✨ **NEW:** S5 prompt redesigned — detects conclusive contribution forms (systematic findings, frameworks, evidence-based recommendations, knowledge gap addressed, synthesis beyond description)
- ✨ **NEW:** Classification rule based on signal semantics: S4+S5+(S2a OR S2b) → CIENTÍFICO; partial signals → DIVULGACIÓN; no signals → OPINIÓN
- ✨ **NEW:** LLM response parsing updated — extracts SI/NO from full response (100 tokens) not just first word
- ✨ **NEW:** Bibliography section excluded from text sample via paragraph-level header detection (≤30 chars) — eliminates false cuts from mid-prose mentions of "referencias"
- ✨ **NEW:** Text sampling redesigned — first 3500 chars (intro) + last 2500 chars (conclusion) replacing sequential 7000-char cut
- 🔧 **FIXED:** S4/S5 failing on systematic review articles — sample now always includes introduction and conclusion sections

**Earlier v0.9 fixes:**
- ✨ **NEW:** `references` field added to `DocumentContent` model
- ✨ **NEW:** `content_extractor.py` populates references via `ReferenceParser`
- ✨ **NEW:** `business_logic/vocab/methodological_terms.py` — vocabulary file for future expansion
- 🔧 **FIXED:** IMRyD false positives — structure analyzer scans only short paragraphs (≤5 words)
- 🔧 **FIXED:** S3 calibration — threshold 4 terms + mandatory hard term
- 🔧 **FIXED:** Quality analyzer call 2 — parser handles numbered and unnumbered headers
- 🔧 **FIXED:** Grammar label — "ortografía" removed (spelling not currently checked)
- 🔧 **FIXED:** Grammar false positives — misspelling filter
- 🔧 **FIXED:** Title/author extraction improvements
- 🔧 **FIXED:** Reference parser — `Fuentes bibliográficas consultadas` support
- 🔧 **FIXED:** Version header updated to v0.9
- 🔧 **FIXED:** Citation matcher `_normalize_author()` — now extracts first author surname only, match rate improved from 13.8% to 93.1%
- 🔧 **FIXED:** Title extraction — trailing colon removed from first title part before combining with subtitle
- 🔧 **FIXED:** Author extraction — title line count detection prevents subtitle from being misread as author

### v0.8 (Q1 2026)
- Gradio web interface, feedback pipeline, two-call LLM quality analysis

### v0.7 (January 2026)
- EUMIC compliance, grammar checker, APA 7 validation

### v0.6 (December 2025)
- Citation-reference validation, IMRyD detection

---

## 🗺️ Roadmap

### v0.9 (Q2 2026) ✅ Current
- ✅ Classification system redesign (Option B)
- ✅ S4/S5 precision prompt engineering
- ✅ Bibliography-aware text sampling
- ✅ IMRyD false positive fix
- ✅ Quality analyzer improvements
- ⬜ Citation matching investigation (low match rate)
- ⬜ Title/author extraction (subtitle misread as author)
- ⬜ Batch processing

### v0.9 → v1.0 (Security & Deployment)
- 🔒 File validation, authentication, rate limiting
- 🌐 Institutional web server deployment (supervised availability model)
- 👥 Multi-user support

### v1.0 (Q3 2026)
- 🏢 Production deployment at Universidad de la Defensa
- 📊 Editorial workflow management

### v2.0 (Future)
- 🧠 Editorial memory (RAG-based)
- 🤖 Agentic workflows
- 🔄 Batch processing

---

## 🐛 Known Issues

### v0.9
- **Citation matching** low rate (13.8% on test article) — matcher may not handle all APA 7 citation formats
- **Title/author extraction** fails when subtitle appears on line 2 — misread as author
- **Quality analyzer** occasionally returns `No disponible` on one dimension — LLM non-determinism on CPU
- **Misspelling** not detected — LanguageTool misspelling filter excluded to avoid false positives on proper nouns
- **Pleonasm/wordiness** not detected — no free Spanish deterministic engine available
- **Windows-only** COM automation for accurate statistics

---

## 🤝 Contributing

**Contact:** Pablo Salonio (P-SAL) - plsalonio@gmail.com  
**Repository:** https://github.com/P-SAL/silvina-editorial

---

## 📄 License

MIT License

---

## 🙏 Acknowledgments

- **Revista Visión Conjunta** - Editorial team
- **Facultad Militar Conjunta** - Universidad de la Defensa Nacional
- **Ollama Team** - Local LLM infrastructure
- **Claude (Anthropic)** - Development assistance

---

**Last Updated:** April 2026  
**Version:** 0.9  
**Status:** Active Development 🚀
