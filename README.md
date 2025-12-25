# Silvina - Asistente Editorial

[![Status](https://img.shields.io/badge/status-v0.4%20complete-brightgreen)](https://github.com/P-SAL/silvina-editorial)
[![Python](https://img.shields.io/badge/python-3.12-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

**AI-powered editorial assistant for academic journals** | Automated APA validation with local LLM integration

---

## 📖 Overview

Silvina is an intelligent editorial assistant developed for **Revista Visión Conjunta**, designed to streamline the manuscript review process for academic journals. It combines traditional document analysis with modern AI capabilities to provide comprehensive editorial feedback in Spanish.

**Current Version:** v0.4 (Complete)  
**Target Release:** v1.0 by June 2026

---

## 🎯 Development Status

This project is actively under development and focuses on enhancing editorial practices with AI technology, supported by contemporary project management principles. I am collaborating with the editorial team and subject matter experts to ensure high-quality deliverables, refined through rigorous reviews and professional input.

---

## ✨ Features

### ✅ v0.2 - Core Document Analysis (Complete)
- **Accurate Character Counting:** Includes body text, footnotes, and endnotes (0.5% error margin)
- **Format Compliance Checking:** Validates article length against journal guidelines (16k-24k short, 36k-40k long)
- **Microsoft Word Integration:** Direct COM automation for seamless document analysis
- **LLM-Powered Grammar Review:** Integration with Ollama for AI-driven analysis
- **Local Processing:** Privacy-focused design using local LLM models (llama3.2:1b / llama3.1:8b)
- **Professional Spanish Reports:** Generates formatted editorial review reports

### ✅ v0.3 - APA Citation Extraction (Complete)
- **Referencias Section Detection:** Finds Referencias/Bibliografía sections in documents
- **Pattern-Based Extraction:** Uses proven regex patterns for reliable text extraction
- **Multiple Heading Support:** Handles variations in section naming

### ✅ v0.4 - Object-Oriented Architecture with Validation (Complete)
- **Document Class:** 
  - Word document loading and management
  - Referencias section extraction (handles any heading variant)
  - Accurate character counting (matches Word within 0.5%)
  - LLM integration for grammar/style review
  - Professional Spanish report generation
  
- **Reference Class:**
  - Individual citation encapsulation
  - APA author format validation (`Apellido, I.` pattern)
  - Year format validation (`(YYYY)` pattern)
  - Detailed validation reporting
  
- **Features:**
  - Handles merged references in single paragraphs
  - Detects 8/8 references in test documents
  - Identifies formatting issues (organization names, missing years)
  - Generates timestamped report files
  - Clean error handling and user feedback

---

## 🛠️ Technical Stack

- **Language:** Python 3.12
- **Architecture:** Object-Oriented (Document & Reference classes)
- **AI/LLM:** Ollama (llama3.1:8b recommended for quality analysis)
- **Document Processing:** pywin32 (COM automation for Microsoft Word)
- **Pattern Matching:** Regular expressions for APA validation
- **Development:** VS Code, Git, virtual environments

---

## 📦 Installation

### Prerequisites
- Python 3.12+
- Microsoft Word (2016 or later)
- Windows 10/11
- [Ollama](https://ollama.ai/) installed and running

### Setup

1. **Clone the repository:**
```bash
git clone https://github.com/P-SAL/silvina-editorial.git
cd silvina-editorial
```

2. **Create virtual environment:**
```bash
python -m venv venv312
source venv312/Scripts/activate  # On Windows Git Bash
```

3. **Install dependencies:**
```bash
pip install -r requirements.txt
```

4. **Register pywin32 (requires administrator rights):**
```bash
python venv312/Scripts/pywin32_postinstall.py -install
```

5. **Start Ollama server:**
```bash
ollama serve
```

6. **Pull LLM model (recommended):**
```bash
# For best quality (32GB RAM):
ollama pull llama3.1:8b

# For lightweight systems (8GB RAM):
ollama pull llama3.2:1b
```

---

## 🚀 Usage

### v0.4 Workflow

1. **Prepare your document** (must have Referencias/Bibliografía section)
2. **Run Silvina:**
```bash
python silvina_editorial_v0.4.py
```

3. **Review the generated report** (console + saved file)

### Example Output
```bash
$ python silvina_editorial_v0.4.py

======================================================================
SILVINA v0.4 - ASISTENTE EDITORIAL
======================================================================

✅ Connected: C:\Users\usuario\Desktop\article.docx
🔍 Characters: 22,188

🤖 Analizando con LLM...

======================================================================
SILVINA - ASISTENTE EDITORIAL v0.4
======================================================================

Documento: article.docx
Fecha: 25/12/2025 12:51
Caracteres totales: 22,188

======================================================================
REVISIÓN DE GRAMÁTICA Y ESTILO (LLM)
======================================================================

[Grammar and style analysis in Spanish]

======================================================================
VALIDACIÓN DE REFERENCIAS APA
======================================================================

Referencias encontradas: 8
✅ Válidas: 4
❌ Con problemas: 4

[Detailed validation results]

💾 Reporte guardado: reporte_silvina_v04_20251225_125256.txt
```

---

## 📊 Project Roadmap

**Completed Milestones:**
- ✅ **Nov 2025:** v0.1 - Basic document analysis
- ✅ **Nov 2025:** v0.2 - LLM integration for grammar/style review
- ✅ **Dec 2025:** v0.3 - Referencias extraction with proven patterns
- ✅ **Dec 2025:** v0.4 - OOP refactor with APA validation

**Upcoming:**
- 📅 **Jan 2026:** v0.5 - Enhanced validation (organization names, improved patterns)
- 📅 **Feb-Mar 2026:** v0.6 - Additional APA rules (URLs, DOIs, alphabetical order)
- 📅 **Apr 2026:** v0.7 - Duplicate detection and advanced checks
- 📅 **May 2026:** v0.9 - Beta testing with Revista Visión Conjunta
- 🎯 **June 2026:** v1.0 - Production release

---

## 📁 Project Structure
```
silvina-editorial/
├── silvina_editorial_v0.2.py    # LLM integration version
├── silvina_editorial_v0.3.py    # Referencias extraction
├── silvina_editorial_v0.4.py    # Current: OOP with validation
├── requirements.txt              # Python dependencies
├── daily_log.txt                 # Development journal
├── README.md                     # This file
└── venv312/                      # Virtual environment (not in git)
```

---

## 🧪 Testing

**Test Document:** Academic article with 8 APA references
**Results (v0.4):**
- Character count: 22,188 (accurate)
- References extracted: 8/8 (100%)
- Valid APA format: 4/8 (50% - identifies real formatting issues)
- LLM review: Complete grammar/style analysis
- Report generation: Timestamped file saved successfully

---

## 🤝 Contributing

Feedback and suggestions are welcome via GitHub Issues. If you're working on similar editorial workflows or academic journal automation, feel free to reach out!

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 👤 Author

**Pablo Salonio**  
Associate Dean for Research | AI Agent Developer  
📧 plsalonio@gmail.com  
🔗 [LinkedIn](https://www.linkedin.com/in/pablosalonio)  
💻 [GitHub](https://github.com/P-SAL)

---

## 🙏 Acknowledgments

- Built for **Revista Visión Conjunta** academic journal
- Designed for editorial teams requiring Spanish-language APA 7 compliance
- Powered by [Ollama](https://ollama.ai/) for local LLM processing

---

**⭐ If you find this project interesting, consider starring the repository!**