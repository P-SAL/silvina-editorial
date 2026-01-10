# Silvina - AI-Powered Editorial Assistant

[![Version](https://img.shields.io/badge/version-v0.6-blue)](https://github.com/P-SAL/silvina-editorial)
[![Python](https://img.shields.io/badge/python-3.12-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Status](https://img.shields.io/badge/status-Production%20Ready-success)](https://github.com/P-SAL/silvina-editorial)

**Automated editorial validation for Spanish academic journals** | APA 7 compliance • EUMIC guidelines • Citation integrity • IMRyD validation • Local LLM integration

---

## 📖 Overview

Silvina is an intelligent editorial assistant developed for **Revista Visión Conjunta** at Facultad Militar Conjunta - Universidad de la Defensa Nacional, Argentina. It automates manuscript review with deterministic structural validation and AI-powered quality analysis.

**Current Version:** v0.6 (January 2026)  
**Architecture:** Two-tier analysis (Structural + Selective LLM)  
**Accuracy:** 100% citation matching • 100% reference validation • Zero false positives

---

## ✨ Features

### 🆕 v0.6 - Enhanced Analysis & Citation Integrity

#### **Deterministic Article Classification**
- **Type Detection:** Científica (30-50K chars + IMRyD) vs Divulgación (~30K chars, flexible structure)
- **Rule-Based Logic:** No LLM guessing—pure Python validation using EUMIC thresholds
- **Confidence Scoring:** 0-10 scale with detailed justification

#### **Citation-Reference Integrity**
- **Smart Matching:** Handles organizational authors (`Ministerio de Economía`, `CIA`)
- **Year Variants:** Recognizes `2020a`, `1983-2003`, `2004, diciembre 15`
- **Orphaned Detection:** 
  - 🔴 CRITICAL: Citations without references
  - 🟡 WARNING: Uncited references (severity depends on section type)
- **Abbreviation Support:** Matches `CIA-` citations with full organizational names

#### **IMRyD Structure Validation**
- **Section Detection:** Introducción, Métodos, Resultados, Discusión, Conclusiones
- **Order Verification:** Flags out-of-sequence sections
- **Length Check:** Validates minimum word counts per section
- **Missing Sections:** Identifies incomplete manuscripts

#### **Enhanced Reference Validation**
- **Spanish APA 7 Compliance:** Author format, year, conjunctions (`y` not `&`)
- **Organizational Authors:** Correctly validates institutional sources
- **DOI/URL Detection:** Identifies deprecated `Recuperado de` format
- **Alphabetical Order:** Verifies reference list sorting

#### **Two-Tier Analysis Strategy**
- **Tier 1 (Always):** Full structural validation (citations, references, IMRyD)
- **Tier 2 (Selective):**
  - Documents ≤5,000 words → Full LLM analysis
  - Documents >5,000 words → Strategic sampling (key sections only)
- **Ollama Integration:** Local LLM (llama3-gradient:8b-instruct-1048k) for grammar/style review

#### **Professional Reporting**
- **Progress Bars:** Real-time feedback via `tqdm` (paragraphs, citations, LLM streaming)
- **Word Export:** Formatted `.docx` reports with color-coded findings
- **Text Reports:** Timestamped `.txt` files for archival

---

## 📊 Validation Metrics

| Feature | Status | Accuracy |
|---------|--------|----------|
| Character Counting | ✅ | 99.7% vs MS Word |
| Citation Extraction | ✅ | 100% detection |
| Reference Validation | ✅ | 100% APA 7 Spanish |
| Citation-Reference Matching | ✅ | 100% (46/46 test) |
| IMRyD Detection | ✅ | 5/5 sections |
| Organizational Authors | ✅ | 100% (3/3 test) |
| False Positives | ✅ | 0% |

**v0.6 Test Results:**
- Document: 46,218 characters, 7,847 words
- Citations: 46 extracted, 46/46 matched
- References: 47 found, 47/47 valid
- IMRyD: 5/5 sections detected
- Classification: Científica (confidence: alta, score: 10/10)

---

## 🛠️ Technical Architecture

### **Object-Oriented Design**

**Core Classes:**
- `Document`: Document loading, analysis orchestration, report generation
- `Reference`: APA 7 validation, similarity detection
- `Citation`: In-text citation parsing (narrativa/parentética)
- `CitationMatcher`: Integrity validation with smart key matching
- `StructureValidator`: IMRyD detection and verification
- `ArticleClassifier`: Deterministic type classification
- `HybridAnalysisStrategy`: Tier 1/2 analysis planning

### **Technology Stack**
- **Language:** Python 3.12
- **Document Processing:** pywin32 (COM automation)
- **Progress Bars:** tqdm
- **Word Export:** python-docx
- **AI/LLM:** Ollama with llama3-gradient:8b-instruct-1048k-q5_K_M
- **Pattern Matching:** Advanced regex for Spanish APA citations
- **Version Control:** Git with semantic versioning

### **Design Patterns**
- Single Responsibility Principle
- Composition over inheritance
- Deterministic validation before AI inference
- Defensive programming with comprehensive error handling

---

## 📦 Installation

### Prerequisites
- **Python 3.12+**
- **Microsoft Word** (2016 or later)
- **Windows 10/11** (for COM automation)
- **RAM:** 16GB minimum, 32GB recommended for full LLM features
- **[Ollama](https://ollama.ai/)** (optional, for Tier 2 analysis)

### Setup
```bash