"""
Silvina Editorial v0.7 - Configuration
"""

# ============================================================
# OLLAMA CONFIGURATION
# ============================================================
OLLAMA_MODEL = 'llama3-gradient:8b-instruct-1048k-q4_K_M'
OLLAMA_URL = 'http://localhost:11434/api/generate'

# ============================================================
# ARTICLE CLASSIFICATION THRESHOLDS
# ============================================================
# Character limits
CIENTIFICO_MIN_CHARS = 30000
CIENTIFICO_MAX_CHARS = 50000
DIVULGACION_TARGET_CHARS = 30000
DIVULGACION_TOLERANCE = 5000

# Citation thresholds
MIN_CITATIONS_SCIENTIFIC = 5
MIN_SECTIONS_SCIENTIFIC = 3
MIN_BIBLIOGRAPHY_SCIENTIFIC = 1000  # characters

# ============================================================
# LLM ANALYSIS SETTINGS
# ============================================================
FULL_LLM_THRESHOLD = 5000  # words - when to use full document vs sampling

# LLM generation parameters
LLM_TEMPERATURE = 0.4
LLM_NUM_PREDICT = 150  # Token limit (allows ~120 words in Spanish)
LLM_TOP_P = 0.95
LLM_TOP_K = 40
LLM_REPEAT_PENALTY = 1.3
LLM_NUM_CTX = 8192

# Stop sequences
LLM_STOP_SEQUENCES = [
    '\n\n\n',
    '\n4.',
    'RECOMENDACIONES',
    'En conclusión',
    '---',
    '==='
]

# ============================================================
# IMRYD STRUCTURE REQUIREMENTS
# ============================================================
REQUIRED_SECTIONS = {
    "introducción": {"order": 1, "min_words": 300, "aliases": ["introduccion", "marco teórico", "marco teorico"]},
    "métodos": {"order": 2, "min_words": 200, "aliases": ["metodos", "metodología", "metodologia", "método"]},
    "resultados": {"order": 3, "min_words": 300, "aliases": ["resultados y análisis", "resultados y analisis"]},
    "discusión": {"order": 4, "min_words": 300, "aliases": ["discusion"]},
    "conclusiones": {"order": 5, "min_words": 150, "aliases": ["conclusión", "conclusion"]},
}