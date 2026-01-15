"""
Article Analyzer - Main orchestrator for document analysis
"""

from typing import Dict, List
from dataclasses import dataclass

from data_access.word_reader import WordDocumentReader
from data_access.content_extractor import ContentExtractor
from business_logic.structure_validator import StructureValidator
from business_logic.citation_matcher import CitationMatcher
from business_logic.article_classifier import ArticleClassifier
from business_logic.quality_analyzer import QualityAnalyzer

from domain.models import Citation, Reference, Section
from domain.enums import ArticleType


@dataclass
class AnalysisResult:
    """Complete analysis result."""
    citations: List[Citation]
    references: List[Reference]
    sections: List[Section]
    classification: Dict
    citation_report: str
    quality_report: str
    word_count: int
    char_count: int
    section_type: str


class ArticleAnalyzer:
    """Main orchestrator for document analysis."""
    
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.reader = WordDocumentReader(filepath)
        
        # Analysis results
        self.citations = []
        self.references = []
        self.sections = []
        self.classification = None
        self.section_type = "Referencias"
        self.word_count = 0
        self.char_count = 0
    
    def analyze(self) -> AnalysisResult:
        """
        Run complete Tier 1 + Tier 2 analysis.
        
        Returns:
            AnalysisResult with all findings
        """
        print("\n" + "="*70)
        print("SILVINA v0.7 - ANÁLISIS COMPLETO")
        print("="*70 + "\n")
        
        # ===== TIER 1: STRUCTURAL ANALYSIS (Deterministic) =====
        print("🔍 TIER 1: ANÁLISIS ESTRUCTURAL (Determinístico)")
        print("-" * 70)
        
        # Open document
        if not self.reader.open():
            raise Exception("No se pudo abrir el documento")
        
        # Get basic metrics
        self.word_count = self.reader.get_word_count()
        self.char_count = self.reader.get_character_count()
        
        print(f"📊 Palabras: {self.word_count:,}")
        print(f"📊 Caracteres: {self.char_count:,}\n")
        
        # Extract paragraphs (cached for performance)
        paragraphs = self.reader.get_paragraphs()
        print(f"✅ Extraídos {len(paragraphs)} párrafos\n")
        
        # Extract structured data
        extractor = ContentExtractor(paragraphs)
        
        self.citations = extractor.extract_citations()
        self.references, self.section_type = extractor.extract_references()
        self.sections = extractor.extract_sections()
        
        # Classify article type
        print("\n📋 Clasificando tipo de artículo...")
        classifier = ArticleClassifier()
        self.classification = classifier.classify({
            'citations_count': len(self.citations),
            'imryd_sections': len(self.sections),
            'bibliography_chars': sum(len(r.text) for r in self.references),
            'char_count': self.char_count
        })
        
        print(f"✅ Clasificación: {self.classification['type'].value}")
        print(f"   Confianza: {self.classification['confidence']}")
        print(f"   Puntuación: {self.classification['score']}/10")
        
        # Validate citations
        print("\n🔗 Validando integridad de citas...")
        matcher = CitationMatcher(self.citations, self.references)
        citation_report = matcher.generate_report(self.section_type)
        
        # ===== TIER 2: QUALITY ANALYSIS (LLM) =====
        print("\n" + "="*70)
        print("🧠 TIER 2: ANÁLISIS DE CALIDAD (LLM)")
        print("-" * 70)
        
        # Prepare content for LLM
        content = self._prepare_llm_content(paragraphs)
        
        # Build Tier 1 statistics
        tier1_stats = {
            'citations': len(self.citations),
            'references': len(self.references),
            'imryd_sections': len(self.sections),
            'formal_valid': len(matcher.find_orphaned_citations()) == 0
        }
        
        # Run LLM analysis
        quality_analyzer = QualityAnalyzer()
        quality_report = quality_analyzer.analyze(
            content=content,
            article_type=self.classification['type'],
            tier1_stats=tier1_stats
        )
        
        # Close Word connection
        self.reader.close()
        
        # Return complete results
        return AnalysisResult(
            citations=self.citations,
            references=self.references,
            sections=self.sections,
            classification=self.classification,
            citation_report=citation_report,
            quality_report=quality_report,
            word_count=self.word_count,
            char_count=self.char_count,
            section_type=self.section_type
        )
    
    def _prepare_llm_content(self, paragraphs: List[str]) -> str:
        """Prepare document content for LLM analysis."""
        # For now, send full document (up to 8000 chars in prompt)
        # In future, we can implement strategic sampling for large docs
        
        # Filter out empty paragraphs and join
        meaningful_paras = [p for p in paragraphs if len(p.strip()) > 20]
        full_text = '\n\n'.join(meaningful_paras)
        
        return full_text