# Silvina - Asistente Editorial

[![Status](https://img.shields.io/badge/status-in%20development-yellow)](https://github.com/P-SAL/silvina-editorial)
[![Python](https://img.shields.io/badge/python-3.12-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

**AI-powered editorial assistant for academic journals** | Automated document analysis with local LLM integration

---

## 📖 Overview

Silvina is an intelligent editorial assistant developed for **Revista Visión Conjunta**, designed to streamline the manuscript review process for academic journals. It combines traditional document analysis with modern AI capabilities to provide comprehensive editorial feedback in Spanish.

This project is actively under development and focuses on enhancing editorial practices with AI technology, supported by contemporary project management principles. The editorial team and subject matter experts are collaborating to ensure high-quality deliverables, refined through rigorous reviews and professional input.

**Current Version:** v0.4 (In Development)  
**Target Release:** v1.0 by June 2026

---

## ✨ Features

### ✅ Implemented (v0.1)
- **Accurate Character Counting:** Includes body text, footnotes, and endnotes (matches Microsoft Word exactly)
- **Format Compliance Checking:** Validates article length against journal guidelines (16k-24k for short articles, 36k-40k for long articles)
- **Microsoft Word Integration:** Direct COM automation for seamless document analysis
- **Professional Spanish Reports:** Generates formatted editorial review reports
- **Multi-computer Support:** Works across different Windows environments

### 🚧 Implemented (v0.2)
- **LLM-Powered Grammar Review:** Integration with Ollama for AI-driven grammar and style analysis
- **Local Processing:** Privacy-focused design using local LLM models (llama3.2:1b / llama3.1:8b)
- **Intelligent Feedback:** Structured editorial recommendations in Spanish

### 📋 Implemented (v0.3)
- APA citation format validation
- Content structure analysis
- Reference list verification
- PDF report export
- Batch processing capabilities

### 📋 In progress (v0.4)
- Built Reference class with:
  * __init__ stores citation text
  * validate_author() checks format
  * validate_year() checks format
  * is_valid() combines all checks
  * get_validation_report() returns detailed results
- Tested with 7 different references
- All validations working correctly
---

## 🛠️ Technical Stack

- **Language:** Python 3.12
- **AI/LLM:** Ollama (llama3.2:1b for lightweight systems, llama3.1:8b for powerful machines)
- **Document Processing:** pywin32 (COM automation for Microsoft Word)
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
python -m venv venv
source venv/Scripts/activate  # On Windows Git Bash
```

3. **Install dependencies:**
```bash
pip install -r requirements.txt
```

4. **Register pywin32 (requires administrator rights):**
```bash
python venv/Scripts/pywin32_postinstall.py -install
```

5. **Start Ollama server:**
```bash
ollama serve
```

6. **Pull LLM model:**
```bash
# For systems with 8GB RAM:
ollama pull llama3.2:1b

# For systems with 32GB+ RAM:
ollama pull llama3.1:8b
```

---

## 🚀 Usage

### Basic Workflow

1. **Open your document in Microsoft Word**
2. **Keep the document active** (visible window)
3. **Run Silvina:**
```bash
python silvina_editorial.py
```

4. **Enter document path** when prompted
5. **Choose whether to use LLM review** (optional)
6. **Review the generated report**

### Example
```bash
$ python silvina_editorial_v02.py

╔════════════════════════════════════════════════════════════════╗
║              SILVINA - ASISTENTE EDITORIAL v0.2                ║
╚════════════════════════════════════════════════════════════════╝

📁 Ingrese la ruta del documento .docx: D:\Documentos\articulo.docx
🤖 ¿Usar revisión LLM? (s/n, Enter=sí): 

🔄 Conectando con Word abierto...
   ✓ Word encontrado
   ✓ Documento: articulo.docx
   ✓ Total caracteres: 19,650

🤖 Analizando gramática y estilo con LLM...
   ✓ Revisión LLM completada

[Report generated and saved]
```

---

## 📊 Project Status

**Development Timeline:**
- ✅ **Nov 2025:** v0.1 - Core functionality (character counting, format compliance)
- 🔄 **Dec 2025:** v0.2 - LLM integration (in progress)
- 📅 **Jan-Feb 2026:** v0.3 - APA citation checking
- 📅 **Mar-Apr 2026:** v0.4 - Content analysis features
- 📅 **May 2026:** v0.9 - Beta testing and refinement
- 🎯 **June 2026:** v1.0 - Production release

**Active Development:** Regular commits, ongoing feature additions

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