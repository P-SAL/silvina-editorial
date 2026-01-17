# Silvina Editorial Assistant v0.7

[![Version](https://img.shields.io/badge/version-v0.7-blue)](https://github.com/P-SAL/silvina-editorial)
[![Python](https://img.shields.io/badge/python-3.12-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Status](https://img.shields.io/badge/status-Active%20Development-yellow)](https://github.com/P-SAL/silvina-editorial)

**AI-powered manuscript review for Spanish academic journals** | EUMIC compliance ‚Ä¢ APA 7 validation ‚Ä¢ Modular architecture ‚Ä¢ LLM-powered quality analysis

---

## Ì≥ñ Overview

Silvina v0.7 is an intelligent editorial assistant for **Revista Visi√≥n Conjunta** (Facultad Militar Conjunta - Universidad de la Defensa Nacional, Argentina). It automates academic manuscript review using **deterministic structural validation** and **selective AI-powered analysis**.

**Current Version:** v0.7 (January 2026)  
**Architecture:** Modular 4-layer design (Domain ‚Üí Data Access ‚Üí Business Logic ‚Üí Presentation)  
**LLM Integration:** Ollama (llama3-gradient:8b-instruct-1048k-q4_K_M)

---

## ‚ú® Key Features

### ÌøóÔ∏è **Modular Architecture (NEW in v0.7)**

**4-Layer Design:**
```
domain/           # Core models & enums
data_access/      # Document parsing
business_logic/   # Analysis engines
presentation/     # Output formatting
```

### Ì≥ö **Document Analysis Pipeline**

**7-Step Process:**
1. Ì≥ñ Document Reading
2. Ì¥ç Content Extraction  
3. Ì≥ö Citation & Reference Parsing
4. Ìø∑Ô∏è Article Classification
5. ‚≠ê Quality Analysis
6. Ì≥ã Structure Validation
7. Ì¥ó Citation Matching

### Ì≥ä **Multi-Format Reports**

- Ì≥Ñ Text Report (`.txt`)
- Ì≥ò Word Report (`.docx`) 
- Ì≥ä JSON Data (`.json`)

**Includes publishability decision** ‚úÖ/‚ùå

---

## Ì∫Ä Quick Start
```bash
# Install
pip install python-docx ollama

# Run analysis
python main.py "document.docx"
```

---

## Ì∑∫Ô∏è Roadmap

- **v0.8** (Q1 2026): Gradio web interface
- **v0.9** (Q2 2026): Batch processing
- **v1.0** (Q3 2026): Production deployment

---

**Version:** 0.7 | **Last Updated:** January 17, 2026
