import ollama

def check_grammar(text, model='llama3.2:1b'):
    """
    Sends text to Ollama for grammar and style review in Spanish.
    
    Args:
        text (str): The text to review
        model (str): Ollama model to use
    
    Returns:
        str: Grammar and style feedback in Spanish
    """
    
    # Create prompt for academic Spanish review
    prompt = f"""Eres un revisor editorial de textos académicos en español.

Revisa el siguiente texto y proporciona:
1. Errores gramaticales (si los hay)
2. Sugerencias de estilo académico
3. Calificación general (Excelente/Bueno/Necesita revisión)

Texto a revisar:
{text}

Responde en español de forma concisa y profesional."""

    try:
        response = ollama.chat(
            model=model,
            messages=[
                {
                    'role': 'user',
                    'content': prompt
                }
            ]
        )
        
        return response['message']['content']
    
    except Exception as e:
        return f"Error al conectar con Ollama: {str(e)}"


# Test the function
if __name__ == "__main__":
    # Sample academic text in Spanish
    test_text = """
    La inteligencia artificial es una tecnología que esta transformando
    diversos sectores. Los investigadores a encontrado aplicaciones
    en medicina, educación y gobernanza.
    """
    
    print("=== REVISIÓN EDITORIAL ===\n")
    print("Texto original:")
    print(test_text)
    print("\n" + "="*50 + "\n")
    
    print("Análisis del LLM:")
    result = check_grammar(test_text)
    print(result)