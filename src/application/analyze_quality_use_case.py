from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.dtos.quality_result_dto import QualityResultDTO
from src.domain.quality.quality_analyzer import QualityAnalyzer


class AnalyzeQualityUseCase:
    def __init__(self, analyzer: QualityAnalyzer) -> None:
        self._analyzer = analyzer

    def execute(self, document_content: DocumentContentDTO, article_type) -> QualityResultDTO:
        return self._analyzer.analyze(document_content=document_content, article_type=article_type)
