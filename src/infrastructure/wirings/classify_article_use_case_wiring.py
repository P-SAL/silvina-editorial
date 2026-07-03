from os import getenv

from dotenv import load_dotenv

from src.application.classify_article_use_case import ClassifyArticleUseCase
from src.domain.classification.article_classification_response_parser import (
    ArticleClassificationResponseParser,
)
from src.domain.classification.article_classification_text_sampler import (
    ArticleClassificationTextSampler,
)
from src.domain.classification.article_classifier import ArticleClassifier
from src.domain.classification.article_size_classifier import ArticleSizeClassifier
from src.domain.classification.classification_rule_table import ClassificationRuleTable
from src.domain.classification.imryd_signal_detector import ImrydSignalDetector
from src.domain.classification.methodological_vocabulary_detector import (
    MethodologicalVocabularyDetector,
)
from src.domain.classification.reference_signal_detector import ReferenceSignalDetector
from src.domain.ports.llm_generator_port import LlmGeneratorPort
from src.infrastructure.adapters.llm_generator.ollama_generator_adapter import (
    OllamaGeneratorAdapter,
)
from src.infrastructure.resources.prompts.classification import PROMPTS_DIR
from src.infrastructure.resources.text_resource_loader import read_text_resource

load_dotenv()


class ClassifyArticleUseCaseWiring:
    def create_use_case(self) -> ClassifyArticleUseCase:
        return ClassifyArticleUseCase(classifier=self._get_article_classifier())

    def _get_article_classifier(self) -> ArticleClassifier:
        return ArticleClassifier(
            llm_generator=self._get_llm_generator(),
            signal_detector=ImrydSignalDetector(),
            article_size_classifier=ArticleSizeClassifier(),
            text_sampler=ArticleClassificationTextSampler(),
            response_parser=ArticleClassificationResponseParser(),
            signal_prompt_template=read_text_resource(
                directory=PROMPTS_DIR, filename="s4_s5_s6_signal_prompt.txt"
            ),
            temperature=float(getenv("ARTICLE_CLASSIFIER_TEMPERATURE", "0.1")),
            num_predict=int(getenv("ARTICLE_CLASSIFIER_NUM_PREDICT", "300")),
            methodological_vocabulary_detector=MethodologicalVocabularyDetector(),
            reference_signal_detector=ReferenceSignalDetector(),
            rule_table=ClassificationRuleTable(),
        )

    def _get_llm_generator(self) -> LlmGeneratorPort:
        model_name = getenv("OLLAMA_MODEL_NAME", "llama3-gradient:8b-instruct-1048k-q4_K_M")
        base_url = getenv("OLLAMA_BASE_URL", "http://localhost:11434")
        return OllamaGeneratorAdapter(model_name=model_name, base_url=base_url)
