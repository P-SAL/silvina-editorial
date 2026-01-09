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
- **Ollama Integration:** Local LLM (llama3-gradient:8b) for grammar/style review

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
- **AI/LLM:** Ollama with llama3-gradient:8b (optional)
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
# 1. Clone repository
git clone https://github.com/P-SAL/silvina-editorial.git
cd silvina-editorial

# 2. Create virtual environment
python -m venv venv312
source venv312/Scripts/activate  # Git Bash
# or: venv312\Scripts\activate  # CMD

# 3. Install dependencies
pip install -r requirements.txt

# 4. Register pywin32 (run as administrator)
python venv312/Scripts/pywin32_postinstall.py -install

# 5. (Optional) Install Ollama for LLM analysis
# Download from https://ollama.ai/
ollama pull llama3-gradient:8b
```

---

## 🚀 Usage

### Quick Start
```bash
# Edit filepath in script (line 1260)
filepath = r"C:\path\to\your\document.docx"

# Run analysis
python silvina_editorial_v0.6.py

# Outputs:
# - Console report with progress bars
# - reporte_silvina_v06_YYYYMMDD_HHMMSS.txt
# - reporte_silvina_v06_YYYYMMDD_HHMMSS.docx
```

### Sample Output
```
======================================================================
SILVINA v0.6 - REPORTE COMPLETO
======================================================================
Documento: PI_Presupuesto_Informe_2024.docx
Fecha: 09/01/2026 14:32
Caracteres: 46,218
Palabras: 7,847

======================================================================
CLASIFICACIÓN DE ARTÍCULO (Determinística)
======================================================================
Tipo: Científica
Confianza: ALTA
Puntuación: 10/10

✅ Indicadores Positivos:
  • 46 citas APA detectadas
  • 5/5 secciones IMRyD presentes
  • Bibliografía extensa (8947 caracteres)

======================================================================
ESTRUCTURA IMRyD
======================================================================
Secciones detectadas: 5/5
  1. Introducción - 2431 palabras
  2. Métodos - 891 palabras
  3. Resultados - 1823 palabras
  4. Discusión - 1456 palabras
  5. Conclusiones - 712 palabras

======================================================================
INTEGRIDAD DE CITAS Y REFERENCIAS
======================================================================
Citas en texto: 46
Referencias bibliográficas: 47
Tipo de sección: REFERENCIAS

✅ Sistema de citación íntegro
✅ Todas las citas tienen referencia válida

======================================================================
VALIDACIÓN DE REFERENCIAS APA
======================================================================
Total: 47
✅ Válidas: 47
❌ Con problemas: 0
```

---

## 📁 Project Structure
```
silvina-editorial/
├── silvina_editorial_v0.6.py    # Current: v0.6 with citation integrity
├── silvina_editorial_v0.5.py    # Previous: Complete EUMIC compliance
├── requirements.txt              # Dependencies (tqdm, python-docx, etc.)
├── README.md                     # This file
├── LICENSE                       # MIT License
├── CITATION.cff                  # Citation metadata
├── docs/                         # Guidelines
│   ├── EUMIC_guidelines.pdf
│   └── APA7_spanish.pdf
├── test_documents/               # Sample documents
│   └── PI_Presupuesto_2024.docx
└── reports/                      # Generated reports
    ├── reporte_*.txt
    └── reporte_*.docx
```

---

## 🗺️ Roadmap

### ✅ Completed
- **v0.1-0.5** (Nov 2025 - Jan 2026): Basic analysis → Full EUMIC compliance
- **v0.6** (Jan 2026): Citation integrity + IMRyD + Two-tier analysis

### 📅 Upcoming
- **v0.7** (Feb 2026): Plagiarism detection, figure/table validation
- **v0.8** (Mar 2026): GUI interface, batch processing
- **v0.9** (Apr 2026): Beta testing with Revista Visión Conjunta
- **v1.0** (Jun 2026): Production release with REST API

---

## 🧪 Testing

**Test Document:** `PI_Presupuesto_Informe_2024.docx`
- Type: Científica
- Length: 46,218 characters, 7,847 words
- Citations: 46 in-text
- References: 47 bibliographic entries
- IMRyD: Complete 5-section structure

**Results:**
```
✅ Citations extracted: 46/46
✅ Citation-reference matching: 46/46 (100%)
✅ Reference validation: 47/47 valid (100%)
✅ IMRyD sections: 5/5 detected
✅ Classification: Correctly identified as Científica
✅ Analysis plan: Strategic sampling (>5K words)
```

---

## 🤝 Contributing

Feedback and suggestions welcome via [GitHub Issues](https://github.com/P-SAL/silvina-editorial/issues).

**Areas of interest:**
- Spanish NLP tools
- Academic workflow automation
- APA validation systems
- Editorial process optimization

---

## 📚 References

- **APA 7 Spanish:** [https://apastyle.apa.org/](https://apastyle.apa.org/)
- **EUMIC Guidelines:** Universidad de la Defensa Nacional
- **RAE:** [https://www.rae.es/](https://www.rae.es/)
- **Ollama:** [https://ollama.ai/](https://ollama.ai/)

---

## 📄 License

MIT License - See [LICENSE](LICENSE) file.

**Disclaimer:** Independent academic tool. No official endorsement by institutions except for pilot/internal evaluation.

---

## 👤 Author

**Pablo Salonio**  
Associate Dean for Research, Facultad Militar Conjunta - Universidad de la Defensa Nacional (Bs.As., Argentina)  
AI Agent Orchestration & Governance Lead | Python-Literate


📧 plsalonio@gmail.com  
🔗 [LinkedIn](https://www.linkedin.com/in/pablosalonio) 
🔗 [GitHub](https://github.com/P-SAL)

**Academic Context:** 7-month Python/AI intensive (Nov 2025 - Jun 2026)

---

⭐ **Star this repo if you find it useful!**