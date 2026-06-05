"""
config.py
Configuration settings for Silvina Editorial Assistant v0.7
"""

import os
from pathlib import Path


class Config:
    """Configuration class for Silvina Editorial Assistant."""
    
    def __init__(self):
        """Initialize configuration with default values."""
        
        # Ollama Configuration
        self.ollama_base_url = os.getenv('OLLAMA_BASE_URL', 'http://localhost:11434')
        self.ollama_model = os.getenv(
            'OLLAMA_MODEL',
            'llama3-gradient:8b-instruct-1048k-q4_K_M'
        )
        
        # Analysis Configuration
        self.min_word_count = 1000  # Minimum words for full analysis
        self.max_word_count = 50000  # Maximum words to process
        
        # Quality Analysis Thresholds
        self.quality_thresholds = {
            'excellent': 9.0,
            'good': 7.0,
            'acceptable': 5.0,
            'needs_improvement': 3.0
        }
        
        # Structure Validation
        self.required_sections = {
            'research_article': [
                'resumen', 'abstract', 'introducción', 'metodología',
                'resultados', 'discusión', 'conclusiones', 'referencias'
            ],
            'review_article': [
                'resumen', 'abstract', 'introducción', 'desarrollo',
                'conclusiones', 'referencias'
            ],
            'reflection_article': [
                'resumen', 'abstract', 'introducción', 'desarrollo',
                'conclusiones', 'referencias'
            ]
        }
        
        # Citation Analysis
        self.min_citations = 5  # Minimum expected citations
        self.min_references = 5  # Minimum expected references
        
        # Output Configuration
        self.output_formats = ['txt', 'docx', 'json']
        self.report_language = 'es'  # Spanish
        
        # LLM Settings
        self.llm_temperature = 0.3  # Lower = more consistent
        self.llm_max_tokens = 2000
        self.llm_timeout = 120  # seconds
        
        # Paths
        self.project_root = Path(__file__).parent.parent
        self.output_dir = self.project_root / 'output'
        
        # Ensure output directory exists
        self.output_dir.mkdir(exist_ok=True)
    
    def get_required_sections(self, article_type: str) -> list:
        """
        Get required sections for a specific article type.
        
        Args:
            article_type: Type of article (research_article, review_article, etc.)
            
        Returns:
            List of required section names
        """
        return self.required_sections.get(article_type, [])
    
    def validate(self) -> bool:
        """
        Validate configuration settings.
        
        Returns:
            True if configuration is valid, False otherwise
        """
        try:
            # Check Ollama URL format
            if not self.ollama_base_url.startswith('http'):
                print(f"⚠ Warning: Invalid Ollama URL: {self.ollama_base_url}")
                return False
            
            # Check model name is not empty
            if not self.ollama_model:
                print("⚠ Warning: Ollama model name is empty")
                return False
            
            # Check thresholds are reasonable
            if not (0 < self.llm_temperature <= 1):
                print(f"⚠ Warning: Invalid temperature: {self.llm_temperature}")
                return False
            
            return True
            
        except Exception as e:
            print(f"⚠ Configuration validation error: {e}")
            return False
    
    def __repr__(self):
        """String representation of configuration."""
        return f"""
Silvina Configuration:
  Ollama URL: {self.ollama_base_url}
  Ollama Model: {self.ollama_model}
  Min Word Count: {self.min_word_count}
  Output Formats: {', '.join(self.output_formats)}
  Project Root: {self.project_root}
        """.strip()


# Create a default configuration instance
default_config = Config()


# Convenience function to load configuration
def load_config() -> Config:
    """
    Load configuration settings.
    
    Returns:
        Config object with loaded settings
    """
    config = Config()
    
    if not config.validate():
        print("⚠ Warning: Configuration validation failed. Using defaults anyway.")
    
    return config


# Environment variable documentation
"""
Environment Variables (optional):

OLLAMA_BASE_URL - URL for Ollama API (default: http://localhost:11434)
OLLAMA_MODEL - Name of the Ollama model to use (default: llama3-gradient:8b-instruct-1048k-q4_K_M)

Example:
  export OLLAMA_BASE_URL="http://localhost:11434"
  export OLLAMA_MODEL="llama3-gradient:8b-instruct-1048k-q4_K_M"
  
Or on Windows:
  set OLLAMA_BASE_URL=http://localhost:11434
  set OLLAMA_MODEL=llama3-gradient:8b-instruct-1048k-q4_K_M
"""
