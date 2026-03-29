from typing import Tuple, List, Dict, Any

def check_gramatica(paragraphs: List[str]) -> Tuple[float, str, List[Dict[str, Any]]]:
    """
    Check grammar and spelling using LanguageTool.
    
    Returns:
        Tuple of (score, summary_feedback, detailed_errors)
    """
    import language_tool_python
        
    # Sample first 5000 chars for analysis
    text_sample = ' '.join(paragraphs[:20])[:5000]
    
    try:
        tool = language_tool_python.LanguageTool('es')
        matches = [m for m in tool.check(text_sample)
                   if m.rule_issue_type != 'misspelling']
        error_count = len(matches)
                
        # Calculate score based on error count
        if error_count == 0:
            score = 10.0
            feedback = "✅ Sin errores detectados"
        elif error_count <= 5:
            score = 8.5
            feedback = f"⚠️ {error_count} errores menores detectados"
        elif error_count <= 15:
            score = 7.0
            feedback = f"⚠️ {error_count} errores detectados"
        else:
            score = 5.0
            feedback = f"❌ {error_count} errores detectados - requiere revisión"
        
        # Build detailed error list
        detailed_errors = []
        for i, match in enumerate(matches[:10], 1):  # Limit to first 10 errors
            error_detail = {
                'number': i,
                'message': match.message,
                'context': match.context,
                'offset': match.offset,
                'length': match.error_length,
                'replacements': match.replacements[:3] if match.replacements else []
            }
            detailed_errors.append(error_detail)
        
        return score, feedback, detailed_errors

    except Exception as e:
        print(f"      ⚠️ Grammar checker error: {e}")
        return 7.0, "Verificación no disponible", []

