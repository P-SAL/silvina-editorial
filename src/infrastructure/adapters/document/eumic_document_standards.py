# Margin standards
REQUIRED_MARGIN_CM: float = 2.5
MARGIN_TOLERANCE_CM: float = 0.3

# Font standards
REQUIRED_FONT_SIZE_PT: int = 12
FONT_SIZE_TOLERANCE_PT: float = 1.0

# Text alignment standards
MAX_UNJUSTIFIED_PARAGRAPH_RATIO: float = 0.3

# Figure detection
FIGURE_CAPTION_PREFIXES: tuple[str, ...] = ("figura", "fig.", "figure")
FIGURE_NUMBERING_PATTERN: str = r"figura\s+(\d+)"

# Table detection
TABLE_CAPTION_PREFIXES: tuple[str, ...] = ("tabla", "table", "cuadro")
TABLE_NUMBERING_PATTERN: str = r"tabla\s+(\d+)"

# Abstract standards
ABSTRACT_SECTION_KEYWORDS: tuple[str, ...] = ("resumen", "abstract", "síntesis", "sumario")
ABSTRACT_PARAGRAPH_LOOKAHEAD: int = 5
MIN_WORDS_FOR_ABSTRACT_CHECK: int = 1000
ABSTRACT_MIN_WORD_COUNT: int = 100
ABSTRACT_MAX_WORD_COUNT: int = 300

# Keywords standards
KEYWORD_SECTION_MARKERS: tuple[str, ...] = (
    "palabras clave",
    "keywords",
    "key words",
    "descriptores",
)
MIN_KEYWORD_COUNT: int = 3
MAX_KEYWORD_COUNT: int = 5
