# Silvina - Asistente Editorial Impulsado por IA

[![Estado](https://img.shields.io/badge/estado-v0.5%20COMPLETO-success)](https://github.com/P-SAL/silvina-editorial)
[![Python](https://img.shields.io/badge/python-3.12-blue)](https://www.python.org/)
[![Licencia](https://img.shields.io/badge/licencia-MIT-green)](LICENSE)

**Validación editorial automatizada para revistas académicas en español** | Cumplimiento APA 7 • Directrices EUMIC • Integración LLM local

[🇬🇧 English Version](README.md) | 🇪🇸 Versión en Español

---

## 📖 Descripción General

Silvina es un asistente editorial inteligente desarrollado para **Revista Visión Conjunta** de la Universidad de la Defensa Nacional, Argentina. Automatiza el proceso de revisión de manuscritos combinando análisis tradicional de documentos con capacidades de IA moderna, proporcionando retroalimentación editorial completa enteramente en español.

**Versión Actual:** v0.5 COMPLETO (Enero 2026)  
**Lanzamiento Objetivo:** v1.0 para Junio 2026  
**Precisión:** 99.7% conteo de caracteres • 100% extracción de referencias • Cero falsos positivos

---

## 🎯 Estado de Desarrollo

**v0.5 está listo para producción** y valida exitosamente:
- Detección de tipo de artículo (Divulgación vs Científica)
- Formato completo de referencias APA 7 en español
- Cumplimiento de directrices editoriales EUMIC
- Gramática y estilo con LLM contextualizado con RAE

Este proyecto sigue prácticas profesionales de desarrollo de software con control de versiones, lanzamientos incrementales y pruebas exhaustivas. Desarrollado como parte de un curso de 7 meses en Desarrollo Python + Agentes de IA (Noviembre 2025 - Junio 2026).

---

## ✨ Características

### ✅ v0.5 COMPLETO - Cumplimiento Total EUMIC

#### **Análisis de Artículos**
- **Detección Automática de Tipo:** Distingue "Divulgación" (~30K caracteres) de "Científica" (30-50K caracteres) usando análisis de estructura IMRyD
- **Validación de Conteo de Caracteres:** Precisión del 99.7% incluyendo cuerpo, notas al pie y notas finales
- **Verificación de Estructura:** Detecta presencia de Introducción, Métodos, Resultados, Discusión, Conclusiones

#### **Validación de Referencias APA 7 en Español**
- **Validación de Formato de Autor:**
  - ✅ Autores personales: `Apellido, I.`
  - ✅ Autores organizacionales: `Google Quantum AI`, `IBM Research`
  - ✅ Formato et al.: `Chen, HZ. et al.`
  
- **Formato de Año:** Valida requisito de paréntesis `(AAAA)`

- **Regla de Conjunción Española:** Detecta uso incorrecto de `&` (debe ser `y` en APA español)
  - ❌ `García, M. & Pérez, J.` 
  - ✅ `García, M. y Pérez, J.`

- **Orden Alfabético:** Verifica que las referencias estén ordenadas por apellido del primer autor

- **Validación DOI/URL:**
  - Detecta presencia de DOI o URL
  - Señala formato obsoleto: `Recuperado de` (debe omitirse en APA 7)

- **Comillas Españolas:** Valida uso de `« »` en lugar de `" "`

- **Detección de Duplicados:** Identifica referencias similares usando umbral de similitud del 85%

- **Detección de Tipo de Sección:** Distingue entre:
  - **Referencias** (solo trabajos citados)
  - **Bibliografía** (todos los trabajos consultados)

#### **Revisión Gramatical Impulsada por IA**
- **Integración LLM Local:** Usa Ollama (llama3-gradient:8b) para análisis de texto en español
- **Contexto de Reglas RAE:** Revisión enfocada usando estándares de la Real Academia Española
- **Gestión de Tokens:** Manejo inteligente de ventana de contexto (8K tokens)
- **Cero Alucinaciones:** Prompts estrictos previenen generación de errores falsos

#### **Reportes Profesionales**
- **UX Limpio:** Referencias válidas mostradas en una línea, problemas detallados
- **Archivos con Marca de Tiempo:** Generación automática de reportes con fecha/hora
- **Transparencia Técnica:** Análisis de capacidad LLM incluido al final del reporte
- **Recomendaciones Accionables:** Orientación clara sobre cómo corregir problemas

---

## 📊 Métricas de Validación (v0.5)

| Tipo de Validación | Implementación | Precisión |
|-------------------|----------------|-----------|
| Conteo de Caracteres | ✅ Completo | 99.7% vs MS Word |
| Extracción de Referencias | ✅ Completo | 100% (8/8 doc prueba) |
| Formato de Autor | ✅ Completo | 100% detección |
| Formato de Año | ✅ Completo | 100% detección |
| Conjunción Española | ✅ Completo | 100% detección |
| Orden Alfabético | ✅ Completo | 100% verificación |
| Presencia DOI/URL | ✅ Completo | 100% detección |
| Detección de Duplicados | ✅ Completo | 85%+ similitud |
| Falsos Positivos | ✅ Eliminados | 0% |

**Resultados de Prueba:**
- Documento: 22,188 caracteres
- Referencias: 8 encontradas, 4 válidas, 4 señaladas (todos problemas legítimos)
- Errores de `&` español: 3 detectados correctamente
- Formato de año faltante: 1 detectado correctamente
- Autores organizacionales: 3 validados correctamente

---

## 🛠️ Arquitectura Técnica

### **Diseño Orientado a Objetos**

**Clase `Document`**
- Automatización COM para integración con Microsoft Word
- Extracción de sección Referencias/Bibliografía
- Cálculo de tokens para gestión de contexto LLM
- Generación de reportes con secciones personalizables
- Orquestación de validaciones

**Clase `Reference`**
- Encapsulación de citas individuales
- Validación de formato APA 7 español
- Detección de DOI/URL
- Comparación de similitud para duplicados

### **Stack Tecnológico**
- **Lenguaje:** Python 3.12
- **Procesamiento de Documentos:** pywin32 (automatización COM)
- **IA/LLM:** Ollama con llama3-gradient:8b
- **Coincidencia de Patrones:** Regex avanzado para texto en español
- **Detección de Similitud:** difflib.SequenceMatcher
- **Desarrollo:** VS Code, Git, entornos virtuales

### **Patrones de Diseño**
- Principio de Responsabilidad Única
- Composición sobre herencia (Document tiene-muchos References)
- Programación defensiva con manejo exhaustivo de errores

---

## 📦 Instalación

### Prerrequisitos
- **Python 3.12+**
- **Microsoft Word** (2016 o posterior)
- **Windows 10/11** (para automatización COM)
- **RAM:** 8GB mínimo, 32GB recomendado para funciones LLM completas
- **[Ollama](https://ollama.ai/)** (opcional, para revisión gramatical)

### Configuración
```bash
# 1. Clonar repositorio
git clone https://github.com/P-SAL/silvina-editorial.git
cd silvina-editorial

# 2. Crear entorno virtual
python -m venv venv312
source venv312/Scripts/activate  # Windows Git Bash
# o
venv312\Scripts\activate  # Windows CMD

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Registrar pywin32 (requiere administrador)
python venv312/Scripts/pywin32_postinstall.py -install

# 5. Instalar Ollama (opcional)
# Descargar de https://ollama.ai/
ollama pull llama3-gradient:8b
```

---

## 🚀 Uso

### Inicio Rápido
```bash
# Ejecutar con revisión gramatical LLM
python silvina_editorial_v0_5.py

# Salidas:
# - Reporte en consola
# - Archivo con marca de tiempo: reporte_silvina_v05_AAAAMMDD_HHMMSS.txt
```

### Uso Programático
```python
from silvina_editorial_v0_5 import Document

# Cargar documento
doc = Document("ruta/al/articulo.docx")
doc.load()

# Generar reporte (con revisión LLM opcional)
report = doc.generate_report(include_llm=True)
print(report)

# Guardar en archivo
with open("reporte.txt", "w", encoding="utf-8") as f:
    f.write(report)

# Limpiar
doc.close()
```

### Salida de Ejemplo
```
======================================================================
SILVINA - ASISTENTE EDITORIAL v0.5 COMPLETE
======================================================================

Documento: escudo_cuantico.docx
Fecha: 01/01/2026 17:19
Caracteres totales: 22,188

======================================================================
TIPO DE ARTÍCULO Y CUMPLIMIENTO EUMIC
======================================================================
Tipo detectado: Divulgación
Caracteres: 22,188
⚠️ Divulgación con 22,188 caracteres (objetivo: ~30,000 ± 5,000)

[... continúa ...]
```

---

## 🗺️ Hoja de Ruta del Proyecto

### ✅ Hitos Completados

- **v0.1** (Nov 2025): Análisis básico de documentos
- **v0.2** (Nov 2025): Integración LLM para revisión gramatical/estilo
- **v0.3** (Dic 2025): Extracción de Referencias con patrones probados
- **v0.4** (Dic 2025): Refactorización OOP con validación APA
- **v0.5** (Ene 2026): **Cumplimiento EUMIC COMPLETO + Todas las reglas APA 7 español**

### 📅 Próximos Lanzamientos

**v0.6** (Feb 2026) - Integridad de Citación
- Detección de citas en texto
- Cruce de citas con lista de referencias
- Detección de citas/referencias huérfanas
- Validación profunda de estructura IMRyD
- Validación de figuras y tablas

**v0.7** (Mar 2026) - Características Avanzadas
- Validación de figuras y tablas
- Verificación de formato de títulos/subtítulos
- Análisis de legibilidad (Flesch-Kincaid para español)
- GUI opcional (interfaz de arrastrar y soltar)

**v1.0** (Jun 2026) - Lanzamiento de Producción 🎯
- Motor de recomendaciones completo
- Integración de base de datos para seguimiento de historial
- Panel web para múltiples usuarios
- API REST para integración externa
- Documentación bilingüe completa (ES/EN)

---

## 📄 Licencia

Este proyecto está licenciado bajo la Licencia MIT.
Usted es libre de usar, modificar y distribuir este software, siempre que se incluya el aviso original de copyright y de licencia.

Este software se proporciona “tal cual”, sin garantía de ningún tipo.Consulte el archivo LICENSE para conocer los detalles completos.

## 📄 Descargo de responsabilidad institucional

Este proyecto es una herramienta de software académico independiente, desarrollada en un contexto educativo y de investigación.
Su uso no implica respaldo oficial, certificación ni responsabilidad institucional por parte de la Universidad de la Defensa Nacional ni de la Revista Visión Conjunta, salvo cuando se indique explícitamente en el marco de pruebas piloto o evaluaciones internas.

## 👤 Autor

**Pablo Salonio**  
Secretario de Investigación - Facultad Militar Conjunta, Universidad de la Defensa Nacional  
Orquestación y Gobernanza de Agentes de IA | Alfabetización técnica en Python

📧 plsalonio@gmail.com  
🔗 [LinkedIn](https://www.linkedin.com/in/pablosalonio)  
💻 [GitHub](https://github.com/P-SAL)

---

## 🙏 Agradecimientos

- Desarrollado para la revista académica **Revista Visión Conjunta**
- Diseñado para equipos editoriales que requieren revisión gramatical y cumplimiento APA 7 en español  
- Impulsado por [Ollama](https://ollama.ai/) para procesamiento LLM local enfocado en privacidad

---

**⭐ Si encuentras útil este proyecto, considera darle una estrella al repositorio**

