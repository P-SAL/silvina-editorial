# Silvina Editorial Assistant v0.95

[![Version](https://img.shields.io/badge/version-v0.95-blue)](https://github.com/P-SAL/silvina-editorial)
[![Python](https://img.shields.io/badge/python-3.12-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Status](https://img.shields.io/badge/status-Active%20Development-yellow)](https://github.com/P-SAL/silvina-editorial)
[![Branch](https://img.shields.io/badge/dev%20branch-silvina__editorial__v095-orange)](https://github.com/P-SAL/silvina-editorial/tree/silvina_editorial_v095)

**AI-powered manuscript review for Spanish academic journals** | EUMIC compliance • APA 7 validation • Modular architecture • LLM-powered quality analysis • Gradio web interface

---

## 📖 Overview

Silvina is an intelligent editorial assistant for **Revista Visión Conjunta** (Facultad Militar Conjunta - Universidad de la Defensa Nacional, Argentina). It automates academic manuscript review using **deterministic structural validation** and **selective AI-powered analysis**.

**Current Version:** v0.95 (Q2 2026)
**Architecture:** Hexagonal Architecture (Domain → Application → Infrastructure)
**LLM Integration:** Ollama (hf.co/unsloth/gemma-4-26B-A4B-it-GGUF:UD-IQ4_XS)
**Interface:** Gradio web UI + CLI
**Output Location:** `Documents\Silvina\reports\` (Word report, JSON data)

---

## 🌿 Branch Structure & Development Workflow

```
main                    ← Production branch — stable, merged from dev
silvina_editorial_v095  ← Active development branch ← ALL work goes here
silvina_editorial_v09   ← Historical reference (read-only)
silvina_editorial_v08   ← Historical reference (read-only)
```

**Development workflow:**
1. All code changes happen on `silvina_editorial_v09`
2. Push to `silvina_editorial_v09` after each session
3. When a stable milestone is reached → merge to `main`
4. Never develop directly on `main`

**Setup on development machine:**
```bash
git checkout silvina_editorial_v095
cd silvina_editorial
source ../venv312/Scripts/activate  # Windows Git Bash
```

**Primary development machine:** DESKTOP-OE2SEGH (32GB RAM, Ryzen 7 8700G)
**Secondary machine:** DESKTOP-LN7Q8I6 (8GB RAM) — code editing only, no LLM inference

---

## ✨ Key Features

### 🎨 **Gradio Web Interface**
- Drag-and-drop file upload
- Interactive result visualization
- One-click Word/JSON download
- Structured expert feedback panel (8 evaluation fields)
- Clean shutdown button

### 🏗️ **Hexagonal Architecture**
```
src/domain/         # Entities, DTOs, enums, exceptions, and ports
src/application/    # Use cases (orchestrate domain behavior)
src/infrastructure/ # Adapters (docx, ollama, win32com, language_tool), wiring, and configuration
```

---

## 📚 Document Analysis Pipeline

**7-Step Process:**

1. **📖 Document Reading** — Extracts paragraphs from .docx
2. **🔍 Content Extraction** — Title, authors, sections, word count
3. **📚 Citation & Reference Parsing** — In-text citations, bibliography, APA 7 validation
4. **🏷️ Article Classification** — Hybrid signal system (S1–S5)
5. **⭐ Quality Analysis** — 4 semantic dimensions via LLM
6. **📋 Structure Validation** — Required sections per article type
7. **🔗 Citation Matching** — Links citations to references

---

## 🎯 Classification System — Hybrid Signal Approach

### Signal Architecture

| Signal | Type | Criterion | Role |
|--------|------|-----------|------|
| S1 — IMRyD override | Deterministic | Complete IMRyD structure | Immediate CIENTÍFICO (0.95) |
| S2a — Reference count | Deterministic | ≥ 12 references | Confidence modulator |
| S2b — Reference recency | Deterministic | ≥ 50% within last 4 years | Confidence modulator |
| S3 — Methodological vocab | Deterministic | ≥ 4 terms + ≥ 1 hard term | Mandatory gate for CIENTÍFICO |
| S4 — Research intent | LLM (combined) | Explicit research intent detected | Mandatory gate for CIENTÍFICO |
| S5 — Conclusive contribution | LLM (combined) | Evidence-based contribution detected | Mandatory gate for CIENTÍFICO |

**S3 vocabulary covers:** quantitative methodology, qualitative/social science, experimental/simulation design, systematic review/meta-analysis, validation and results markers, metrics and measurement — in Spanish and English.

**S4 detects:** verbs of intent, scope markers, problem markers, experimental objective expressions, research questions/hypotheses (single or multiple, numbered or sequential).

**S5 detects:** findings from systematic process, frameworks/models/taxonomies proposed, evidence-based recommendations, knowledge gaps addressed, synthesis beyond description, quantitative experimental results, hypothesis confirmation.

**S4 and S5 are evaluated in a single combined LLM call** for efficiency and consistency.

### Classification Rule

| Condition | Result | Confidence |
|-----------|--------|------------|
| S3 + S4 + S5 + S2a + S2b | CIENTÍFICO | 0.95 |
| S3 + S4 + S5 + S2b | CIENTÍFICO | 0.88 |
| S3 + S4 + S5 + S2a | CIENTÍFICO | 0.80 |
| S3 + S4 + S5 (no S2) | CIENTÍFICO | 0.72 |
| S3 + S5 (no S4) | CIENTÍFICO | 0.72 — reduced confidence, manual review recommended |
| S3 + S4 (no S5) | CIENTÍFICO | 0.72 — reduced confidence, manual review recommended |
| S3 + (S2a or S2b) (no S4/S5) | CIENTÍFICO | 0.70 — reduced confidence, manual review recommended |
| S3 alone (no S2/S4/S5) | CIENTÍFICO | 0.60 — very reduced confidence, manual review required |
| S4 + S5 (no S3) | DIVULGACIÓN | 0.75 |
| S4 or S5 alone | DIVULGACIÓN | 0.65 |
| No signals | OPINIÓN | 0.65 |

### Design Principles
- **S3 is the mandatory methodological gate** — its presence means the article is NOT OPINIÓN
- **S4 and S5 are the primary scientific discriminators** — together with S3 they confirm CIENTÍFICO
- **S2a/S2b modulate confidence** — references support but do not determine classification
- **Reduced confidence classifications always include a manual review recommendation** — the editorial team makes the final call
- **S4/S5 non-determinism on CPU is a known hardware limitation** — GPU inference (v1.0) will resolve this

### Structure Validation by Classification Path
- **S1 classified (IMRyD):** requires full IMRyD sections (Metodología, Resultados, Discusión)
- **S2-S5 classified:** requires DIVULGACIÓN sections (Resumen, Introducción, Desarrollo, Conclusiones, Referencias)

---

## ⭐ Quality Analysis

**4 Semantic Dimensions:**
- **Claridad** — Writing clarity and readability
- **Coherencia** — Logical flow and consistency
- **Argumentación** — Strength of arguments and evidence
- **Conclusiones** — Quality and relevance of conclusions

**Architecture:** Single combined LLM call per dimension pair (Call 1: Claridad + Coherencia / Call 2: Argumentación + Conclusiones). Score inference from narrative when LLM omits score format.

**Text sampling:** First 3500 chars (intro) + last 2500 chars (conclusions), bibliography excluded via paragraph-level header detection.

---

## 📊 Multi-Format Reports

**3 Output Files** saved to `C:\Users\[user]\Documents\Silvina\reports\`:
1. **📘 Word Report** (`_analisis.docx`)
2. **📊 JSON Data** (`_analisis.json`)
3. **💬 Feedback File** (`_feedback.json`) — via Gradio

---

## 🛠️ Technology Stack

| Component | Technology |
|-----------|------------|
| Language | Python 3.12 |
| Web Interface | Gradio |
| Document Parsing | python-docx |
| Word Automation | win32com (Windows COM) |
| LLM Integration | Ollama (local inference) |
| Grammar Checking | LanguageTool |
| Architecture | Hexagonal Architecture |

---

## 📦 Installation

```bash
# 1. Clone repository
git clone https://github.com/P-SAL/silvina-editorial.git
cd silvina-editorial

# 2. Switch to development branch
git checkout silvina_editorial_v095
cd silvina_editorial_v095

# 3. Create virtual environment
python -m venv ../venv312
source ../venv312/Scripts/activate  # Windows Git Bash

# 4. Install dependencies
pip install -r requirements.txt

# 5. Pull LLM model
ollama pull hf.co/unsloth/gemma-4-26B-A4B-it-GGUF:UD-IQ4_XS
```

---

## 🚀 Usage

### Web Interface
```bash
python gradio_app.py
```

### Command Line
```bash
python main.py
```

---

## 📁 Project Structure
```
silvina_editorial_v095/
├── main.py
├── gradio_app.py
├── process_feedback.py
├── version.txt
├── requirements.txt
├── src/
│   ├── domain/                  # Entities, DTOs, enums, exceptions, and ports
│   │   ├── citation/
│   │   ├── classification/
│   │   ├── document/
│   │   ├── dtos/
│   │   ├── entities/
│   │   ├── enums/
│   │   ├── exceptions/
│   │   ├── gateway/
│   │   ├── grammar/
│   │   ├── ports/
│   │   ├── quality/
│   │   ├── recommendation/
│   │   ├── report/
│   │   └── structure/
│   ├── application/             # Use cases (orchestrate domain behavior)
│   │   ├── analyze_document_use_case.py
│   │   └── export_report_use_case.py
│   └── infrastructure/          # Adapters, config, and wirings
│       ├── adapters/
│       │   ├── document/
│       │   ├── gateway/
│       │   ├── grammar/
│       │   ├── llm_generator/
│       │   └── report/
│       ├── config/
│       ├── env_config.py
│       └── wirings/
└── tests/                       # Integration and unit tests
    ├── e2e/
    ├── fixtures/
    ├── smoke/
    └── test_main_cli_args.py
```

---

## 🔄 Version History

### v0.95 (Q2 2026) — Current

**Classification System — Major Revision:**
- ✨ **NEW:** S3 expanded vocabulary — now covers quantitative, qualitative/social science, experimental/simulation, systematic review, validation/results, metrics categories in Spanish and English
- ✨ **NEW:** S4+S5 combined into single LLM call — reduces non-determinism, improves efficiency
- ✨ **NEW:** Granular confidence calibration — 0.60/0.70/0.72/0.80/0.88/0.95 based on signal combination
- ✨ **NEW:** Partial signal cases (S3+S4, S3+S5) → CIENTÍFICO 0.72 with manual review recommendation
- ✨ **NEW:** S3 alone → CIENTÍFICO with reduced confidence rather than DIVULGACIÓN — methodological substance acknowledged
- ✨ **NEW:** Structure validator path-aware — S2-S5 classified CIENTÍFICO uses DIVULGACIÓN structure requirements (not IMRyD)
- 🔧 **FIXED:** APA validator false positives — institutional acronyms (PLANCAMIL, UNESCO), identifiers (arXiv:), date ranges no longer flagged as author surname errors
- 🔧 **FIXED:** Citation matcher — non-author citations skipped in matching logic
- 🔧 **FIXED:** Author extraction — multi-line author lists (semicolon-separated) now collected correctly across up to 3 continuation paragraphs
- 🔧 **FIXED:** Title extraction — lines with semicolons correctly identified as author lists, not subtitles
- 🔧 **FIXED:** Quality analyzer score inference — when LLM omits score format, score inferred from narrative sentiment
- 🔧 **FIXED:** Grammar label — "ortografía" removed (spelling not currently checked)
- 🔧 **FIXED:** `_apply_rule()` comment renamed from "SIGNAL 6" to "CLASSIFICATION RULE"
- ✨ **NEW:** Unmatched citations listed by name in ANÁLISIS FINAL (not just count)
- ✨ **NEW:** Branch-based development workflow established

**Earlier v0.95 fixes:**
- ✨ **NEW:** 5-signal hybrid classification engine
- ✨ **NEW:** Bibliography-aware text sampling (3500+2500 chars)
- ✨ **NEW:** `references` field added to `DocumentContent`
- 🔧 **FIXED:** IMRyD false positives in structure analyzer
- 🔧 **FIXED:** Citation matcher `_normalize_author()` — 13.8% → 93.1% match rate
- 🔧 **FIXED:** Quality analyzer call 2 parser — handles numbered and unnumbered headers
- 🔧 **FIXED:** Reference parser — `Fuentes bibliográficas consultadas` support
- 🔧 **FIXED:** Publishability verdict logic

### v0.8 (Q1 2026)
- Gradio web interface, feedback pipeline, two-call LLM quality analysis

### v0.7 (January 2026)
- EUMIC compliance, grammar checker, APA 7 validation

### v0.6 (December 2025)
- Citation-reference validation, IMRyD detection

---

## 🗺️ Roadmap

### v0.95 (Q2 2026) — Active Development
- ✅ Classification system major revision
- ✅ S3 vocabulary expansion
- ✅ Confidence calibration
- ✅ Structure validator path-aware fix
- ✅ APA validator false positive fix
- ✅ Author extraction multi-line fix
- ✅ Citation matcher improvements
- ✅ Branch-based workflow established
- ⬜ Security measures (file validation, authentication, rate limiting)
- ⬜ Web deployment preparation

### v0.95 → v1.0 (Security & Deployment)
- 🔒 File validation, defusedxml, path traversal protection
- 🔒 Authentication and rate limiting
- 🔒 Prompt injection detection
- 🌐 nginx + HTTPS deployment
- 🌐 Supervised availability model (institutional server)

### v1.0 (Q3 2026)
- 🏢 Production deployment at Universidad de la Defensa
- 🖥️ Hardware upgrade: 64GB RAM + GPU (RTX 4060/4070)
- 🤖 Model upgrade: llama3.1:70b q4_K_M — deterministic S4/S5 inference
- 📊 Editorial analytics dashboard

### v2.0 (Future)
- 🧠 Editorial memory (RAG-based)
- 🤖 Multi-agent R+D orchestration
- 🔄 Batch processing

---

## 🐛 Known Issues & Limitations

| Issue | Status | Notes |
|-------|--------|-------|
| S4/S5 non-determinism on CPU | Known hardware limitation | GPU inference (v1.0) resolves this |
| Misspelling not detected | Deferred | LanguageTool filter excludes misspellings to avoid false positives on proper nouns |
| Pleonasm/wordiness not detected | Deferred | No free Spanish deterministic engine available |
| Citation matching on institutional refs | Partial | NIST, MITRE, arXiv-style citations have low match rates by design |
| Windows-only COM automation | By design | Falls back to python-docx on other platforms |
| Quality analyzer "No disponible" | Intermittent | LLM response format variance on very long documents |

---

## 🤝 Contributing

**Contact:** Pablo Salonio (P-SAL) — plsalonio@gmail.com
**Repository:** https://github.com/P-SAL/silvina-editorial
**Active branch:** `silvina_editorial_v095`

---

## 📄 License

MIT License

---

## 🙏 Acknowledgments

- **Revista Visión Conjunta** — Editorial team for requirements and testing
- **Facultad Militar Conjunta** — Universidad de la Defensa Nacional
- **Ollama Team** — Local LLM infrastructure
- **Claude (Anthropic)** — Development assistance

---

**Last Updated:** July 2026
**Version:** 0.95
**Active Branch:** silvina_editorial_v095
**Status:** Active Development 🚀
