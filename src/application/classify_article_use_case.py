from src.domain.classification.article_classifier import ArticleClassifier
from src.domain.dtos.classification_result_dto import ClassificationResultDTO
from src.domain.dtos.document_content_dto import DocumentContentDTO


class ClassifyArticleUseCase:
    def __init__(self, classifier: ArticleClassifier) -> None:
        self._classifier = classifier

    def execute(self, document_content: DocumentContentDTO) -> ClassificationResultDTO:
        return self._classifier.classify(document_content=document_content)
