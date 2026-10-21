from enum import Enum


class FormulaXmlMarker(str, Enum):
    OMATH = "<m:oMath"
    WORD_EQUATION = "<w:equation"
