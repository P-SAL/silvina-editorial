"""
gradio_app.py
Gradio web interface for Silvina Editorial Assistant v0.8
Provides a user-friendly interface for non-technical editorial staff.
"""

import gradio as gr
import html
import re
import sys
import json
import traceback
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any
import base64
import webbrowser

from src.domain.dtos.analysis_result_dto import AnalysisResultDTO
from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.report_input_dto import ReportInputDTO
from src.domain.enums.publication_verdict import PublicationVerdict
from src.domain.enums.recommendation_priority import RecommendationPriority
from src.domain.exceptions.base_src_error import BaseSrcError
from src.infrastructure.wirings.analyze_document_use_case_wiring import (
    AnalyzeDocumentUseCaseWiring,
)
from src.infrastructure.wirings.export_report_wiring import ExportReportWiring

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))


analyze_document_use_case = AnalyzeDocumentUseCaseWiring().create_use_case()
export_report_use_case = ExportReportWiring().create_use_case()

# ============================================
# CONFIGURATION
# ============================================
EUMIC_COLORS = {
    "primary": "#2C3E50",
    "success": "#27AE60",
    "warning": "#F39C12",
    "danger": "#E74C3C",
}


# ============================================
# JSON SERIALIZATION HELPER
# ============================================
def _prepare_for_json(data: Any) -> Any:
    """Recursively convert enums, datetimes, and DTOs into JSON-serializable structures."""
    if isinstance(data, dict):
        return {k: _prepare_for_json(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [_prepare_for_json(item) for item in data]
    elif isinstance(data, Enum):
        return data.value
    elif isinstance(data, datetime):
        return data.isoformat()
    elif isinstance(data, BaseDTO):
        return _prepare_for_json(data.as_dict())
    else:
        return data


# ============================================
# CORE ANALYSIS FUNCTION
# ============================================
def process_document(uploaded_file):
    """
    Wraps Silvina's analysis pipeline for Gradio.
    Returns: (status_message, results_html, word_report_path, json_report_path)
    """
    if uploaded_file is None:
        return (
            "⚠️ Por favor, cargue un documento Word (.docx)",
            "",
            None,
            None,
            "",
            gr.Button(interactive=True),
        )

    try:
        # Run analysis on uploaded file
        print(f"\n🔍 Procesando: {uploaded_file.name}")
        report = analyze_document_use_case.execute(uploaded_file.name)

        # Generate report filenames
        base_name = Path(uploaded_file.name).stem
        output_dir = Path.home() / "Documents" / "Silvina" / "reports"
        output_dir.mkdir(parents=True, exist_ok=True)

        word_report_path = output_dir / f"{base_name}_analisis.docx"
        json_report_path = output_dir / f"{base_name}_analisis.json"

        # Export Word report via ExportReportUseCase
        export_report_use_case.execute(report_input=report, output_path=str(word_report_path))

        # Instantiate AnalysisResultDTO and save the JSON report
        analysis_result = AnalysisResultDTO(
            filename=report.filename,
            document_content=report.document_content,
            classification=report.classification,
            quality=report.quality,
            structure=report.structure,
            citations=report.citations,
        )
        json_data = _prepare_for_json(analysis_result)
        with open(json_report_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)

        # Create visual results display
        results_html = create_results_display(report)

        success_msg = "✅ Análisis completado exitosamente"

        return (
            success_msg,
            results_html,
            str(word_report_path),
            str(json_report_path),
            str(word_report_path),
            gr.Button(interactive=True),
        )

    except BaseSrcError as exc:
        # Extract clean domain message to display in UI without traceback
        error_msg = f"❌ Error de validación: {exc.dict().get('error', str(exc))}"
        print(f"\n[Domain Error] {error_msg}")
        return (error_msg, "", None, None, "", gr.Button(interactive=True))

    except Exception as e:
        # Log generic unexpected runtime errors
        error_msg = f"❌ Error al procesar el documento: {str(e)}"
        print(f"\n[System Error] {error_msg}")
        traceback.print_exc()
        return (error_msg, "", None, None, "", gr.Button(interactive=True))


# ============================================
# RESULTS DISPLAY
# ============================================
_MARKDOWN_BOLD_PATTERN = re.compile(r"\*\*(.+?)\*\*")


def _render_feedback_html(text: str) -> str:
    """Escape HTML special characters, then render **bold** Markdown as <strong> tags."""
    return _MARKDOWN_BOLD_PATTERN.sub(r"<strong>\1</strong>", html.escape(text))


_SUITABILITY_VERDICT_COLORS = {
    "SUSTENTADA": EUMIC_COLORS["success"],
    "ALINEADO": EUMIC_COLORS["success"],
    "PARCIAL": EUMIC_COLORS["warning"],
    "PARCIALMENTE ALINEADO": EUMIC_COLORS["warning"],
    "NO SUSTENTADA": EUMIC_COLORS["danger"],
    "NO ALINEADO": EUMIC_COLORS["danger"],
}


def _verdict_color(verdict: str) -> str:
    """Return the status color for a contribution/alignment verdict string."""
    return _SUITABILITY_VERDICT_COLORS.get(verdict, "#666")


def _render_editorial_suitability_html(editorial_suitability) -> str:
    """Render the Contribución/Pertinencia editorial suitability section, or empty if absent."""
    if editorial_suitability is None:
        return ""

    return f"""
    <div style="background: #f8f9fa; padding: 20px; border-radius: 6px; margin-bottom: 20px;">
        <h4 style="margin: 0 0 15px 0; color: {
        EUMIC_COLORS["primary"]
    };">🎯 Pertinencia Editorial</h4>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px;">
            <div>
                <div style="font-weight: 500; color: #333; margin-bottom: 4px;">Contribución</div>
                <span style="display: inline-block; padding: 4px 10px; border-radius: 4px; background: {
        _verdict_color(editorial_suitability.contribution_verdict)
    }; color: white; font-size: 13px; font-weight: bold;">
                    {html.escape(editorial_suitability.contribution_verdict)}
                </span>
                <div style="font-size: 13px; color: #666; margin-top: 6px;">{
        _render_feedback_html(editorial_suitability.contribution_observation)
    }</div>
            </div>
            <div>
                <div style="font-weight: 500; color: #333; margin-bottom: 4px;">Alineación con líneas de investigación</div>
                <span style="display: inline-block; padding: 4px 10px; border-radius: 4px; background: {
        _verdict_color(editorial_suitability.alignment_verdict)
    }; color: white; font-size: 13px; font-weight: bold;">
                    {html.escape(editorial_suitability.alignment_verdict)}
                </span>
                <div style="font-size: 13px; color: #666; margin-top: 6px;">{
        _render_feedback_html(editorial_suitability.alignment_justification)
    }</div>
                <div style="font-size: 12px; color: #999; margin-top: 4px;">Líneas: {
        html.escape(editorial_suitability.alignment_lines)
    }</div>
            </div>
        </div>
    </div>
    """


def create_results_display(report: ReportInputDTO) -> str:
    """
    Creates professional HTML display of analysis results.
    """

    # Extract key DTOs
    document_content = report.document_content
    classification = report.classification
    quality = report.quality
    grammar = report.grammar
    structure = report.structure
    citations = report.citations
    apa_validation = report.apa_validation
    verdict = report.verdict

    # Determine overall status from the final publication verdict
    if verdict.verdict == PublicationVerdict.APPROVED:
        status = "APTO PARA PUBLICACIÓN"
        status_color = EUMIC_COLORS["success"]
        status_icon = "✅"
    elif verdict.verdict == PublicationVerdict.WARNING:
        status = "REQUIERE REVISIÓN"
        status_color = EUMIC_COLORS["warning"]
        status_icon = "⚠️"
    else:  # PublicationVerdict.CRITICAL
        status = "NO APTO"
        status_color = EUMIC_COLORS["danger"]
        status_icon = "❌"

    # Count errors
    grammar_errors = len(grammar.errors)
    apa_errors = apa_validation.violation_count
    missing_sections = len(structure.missing_sections)
    unmatched_citations = citations.unmatched_count

    # Filter critical recommendations by priority
    critical_recommendations = [
        rec for rec in report.recommendations if rec.priority == RecommendationPriority.HIGH
    ]

    # Grammar and semantic scores
    grammar_score = grammar.score
    semantic_score = quality.overall_score

    html = f"""
    <div style="font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; background: white; border-radius: 8px;">

        <!-- Document Header -->
        <div style="background: #f8f9fa; padding: 15px; border-radius: 6px; margin-bottom: 20px;">
            <h3 style="margin: 0 0 10px 0; color: {EUMIC_COLORS["primary"]};">
                📄 {document_content.title or "Sin título"}
            </h3>
            <p style="margin: 0; color: #666;">
                <strong style="color: #666;">Autor:</strong> {
        document_content.authors or "No especificado"
    } |
                <strong style="color: #666;">Palabras:</strong> {document_content.word_count:,} |
                <strong style="color: #666;">Tipo:</strong> {
        classification.article_type.value.upper()
    }
            </p>
        </div>

        <!-- Overall Status -->
        <div style="background: {
        status_color
    }; color: white; padding: 25px; border-radius: 8px; margin-bottom: 25px; text-align: center;">
            <h2 style="margin: 0; font-size: 32px;">{status_icon} {status}</h2>
            <p style="margin: 10px 0 0 0; font-size: 16px; opacity: 0.95;">{
        _render_feedback_html(verdict.message)
    }</p>
        </div>

        <!-- Quality Scores -->
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-bottom: 20px;">
            <div style="background: #f8f9fa; padding: 20px; border-radius: 6px; text-align: center;">
                <div style="font-size: 18px; color: #666; margin-bottom: 8px;">Gramática y Ortografía</div>
                <div style="font-size: 42px; font-weight: bold; color: {EUMIC_COLORS["primary"]};">
                    {grammar_score:.1f}<span style="font-size: 24px; color: #999;">/10</span>
                </div>
                <div style="color: #666; margin-top: 5px; font-size: 14px;">
                    {_render_feedback_html(grammar.feedback)}
                </div>
            </div>
            <div style="background: #f8f9fa; padding: 20px; border-radius: 6px; text-align: center;">
                <div style="font-size: 18px; color: #666; margin-bottom: 8px;">Calidad Semántica</div>
                <div style="font-size: 42px; font-weight: bold; color: {EUMIC_COLORS["primary"]};">
                    {semantic_score:.1f}<span style="font-size: 24px; color: #999;">/10</span>
                </div>
                <div style="color: #666; margin-top: 5px; font-size: 14px;">
                    {quality.quality_level.value}
                </div>
            </div>
        </div>

        <!-- Error Summary -->
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 25px;">
            <div style="background: #fff3cd; padding: 15px; border-radius: 6px; text-align: center; border: 2px solid #ffc107;">
                <div style="font-size: 28px; font-weight: bold; color: #856404;">{
        grammar_errors
    }</div>
                <div style="color: #856404; margin-top: 5px; font-size: 13px;">Errores gramaticales</div>
            </div>
            <div style="background: #fff3cd; padding: 15px; border-radius: 6px; text-align: center; border: 2px solid #ffc107;">
                <div style="font-size: 28px; font-weight: bold; color: #856404;">{apa_errors}</div>
                <div style="color: #856404; margin-top: 5px; font-size: 13px;">Errores APA 7</div>
            </div>
            <div style="background: #fff3cd; padding: 15px; border-radius: 6px; text-align: center; border: 2px solid #ffc107;">
                <div style="font-size: 28px; font-weight: bold; color: #856404;">{
        unmatched_citations
    }</div>
                <div style="color: #856404; margin-top: 5px; font-size: 13px;">Citas sin referencia</div>
            </div>
            <div style="background: #fff3cd; padding: 15px; border-radius: 6px; text-align: center; border: 2px solid #ffc107;">
                <div style="font-size: 28px; font-weight: bold; color: #856404;">{
        missing_sections
    }</div>
                <div style="color: #856404; margin-top: 5px; font-size: 13px;">Secciones faltantes</div>
            </div>
        </div>

        <!-- Semantic Dimensions -->
        <div style="background: #f8f9fa; padding: 20px; border-radius: 6px; margin-bottom: 20px;">
            <h4 style="margin: 0 0 15px 0; color: {
        EUMIC_COLORS["primary"]
    };">Dimensiones Semánticas</h4>
            {
        "".join(
            [
                f'''
            <div style="margin-bottom: 10px;">
                <div style="display: flex; justify-content: space-between; margin-bottom: 4px;">
                    <span style="font-weight: 500; color: #333;">{dim.capitalize()}</span>
                    <span style="font-weight: bold; color: {EUMIC_COLORS["primary"]};">{data["score"]:.1f}/10</span>
                </div>
                <div style="background: #dee2e6; height: 8px; border-radius: 4px; overflow: hidden;">
                    <div style="background: {EUMIC_COLORS["primary"]}; height: 100%; width: {data["score"] * 10}%;"></div>
                </div>
                <div style="font-size: 13px; color: #666; margin-top: 4px;">{_render_feedback_html(data.get("feedback", ""))}</div>
            </div>
            '''
                for dim, data in quality.dimension_scores.items()
            ]
        )
    }
        </div>

        <!-- Editorial Suitability -->
        {_render_editorial_suitability_html(quality.editorial_suitability)}

        <!-- Critical Issues -->
        {
        f'''
        <div class="critical-issues-box" style="background: #f8d7da; border-left: 4px solid {EUMIC_COLORS["danger"]}; padding: 15px; margin-bottom: 20px; border-radius: 4px;">
            <h4 style="margin: 0 0 10px 0;">⚠️ Problemas Críticos Detectados</h4>
            <ul style="margin: 0; padding-left: 20px;">
                {"".join([f"<li>{_render_feedback_html(rec.message)}</li>" for rec in critical_recommendations])}
            </ul>
        </div>
        '''
        if critical_recommendations
        else ""
    }

        <!-- Instructions -->
        <div style="background: #e7f3ff; border-left: 4px solid {
        EUMIC_COLORS["primary"]
    }; padding: 15px; border-radius: 4px;">
            <p style="margin: 0; color: #004085; line-height: 1.6;">
                <strong style="color: #004085;">Próximos pasos:</strong><br>
                1. Descargue el <strong style="color: #004085;">informe detallado en Word</strong> para revisar observaciones específicas<br>
                2. Proporcione su <strong style="color: #004085;">evaluación experta</strong> a continuación<br>
                3. Sus comentarios ayudarán a <strong style="color: #004085;">mejorar futuros análisis</strong>
            </p>
        </div>
    </div>
    """

    return html


# ============================================
# FEEDBACK HANDLER
# ============================================
def save_expert_feedback(
    document_name,
    evaluation,
    classification_correct,
    quality_score_fair,
    grammar_real_errors,
    structure_correct,
    citations_correct,
    weakest_section,
    editor_recommendation,
    comments,
):
    """
    Saves structured expert feedback for system improvement.
    """
    if not evaluation:
        return "⚠️ Por favor, seleccione una evaluación general"

    feedback = {
        "timestamp": datetime.now().isoformat(),
        "document": document_name,
        "overall_precision": evaluation,
        "classification_correct": classification_correct or "Sin respuesta",
        "quality_score_fair": quality_score_fair or "Sin respuesta",
        "grammar_real_errors": grammar_real_errors or "Sin respuesta",
        "structure_correct": structure_correct or "Sin respuesta",
        "citations_correct": citations_correct or "Sin respuesta",
        "weakest_section": weakest_section or "Sin respuesta",
        "editor_recommendation": editor_recommendation or "Sin respuesta",
        "comments": comments or "Sin comentarios adicionales",
    }

    # Saves to report file
    reports_dir = Path.home() / "Documents" / "Silvina" / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)
    doc_stem = Path(document_name).stem.replace("_analisis", "") if document_name else "unknown"
    feedback_file = reports_dir / f"{doc_stem}_feedback.json"

    try:
        with open(feedback_file, "w", encoding="utf-8") as f:
            json.dump(feedback, f, ensure_ascii=False, indent=2)

        return "✅ Gracias. Su evaluación ha sido registrada y contribuirá a mejorar el sistema."
    except Exception as e:
        return f"❌ Error al guardar evaluación: {str(e)}"


# ============================================
# GRADIO INTERFACE
# ============================================
def create_interface():
    """
    Creates the Gradio web interface.
    """
    # Load logo as base64
    try:
        logo_path = project_root / "assets" / "SILVINA V08.png"
        with open(logo_path, "rb") as f:
            logo_data = base64.b64encode(f.read()).decode()
        logo_src = f"data:image/png;base64,{logo_data}"
    except Exception as e:
        print(f"⚠️ No se pudo cargar el logo: {e}")
        logo_src = ""

    with gr.Blocks(
        title="Silvina - Asistente Editorial EUMIC V0.8",
    ) as interface:
        # Store document name for feedback
        document_name_state = gr.State("")

        # Header with logo
        if logo_src:
            gr.HTML(f"""
            <div class="logo-container">
                <img src="{logo_src}" alt="EUMIC Logo" style="max-height: 100px; display: block; margin: 0 auto;">
                <h1>Silvina - Asistente Editorial</h1>
                <p>Sistema de revisión automatizado para manuscritos académicos EUMIC</p>
            </div>
            """)
        else:
            gr.HTML(f"""
            <div class="logo-container">
                <h1>Silvina - Asistente Editorial</h1>
                <p>Sistema de revisión automatizado para manuscritos académicos EUMIC</p>
            </div>
            """)

        gr.Markdown("---")

        # Step 1: Upload
        gr.Markdown("### 📄 Paso 1: Cargar Documento")
        gr.Markdown("Seleccione el manuscrito en formato Word (.docx) que desea analizar.")

        file_input = gr.File(label="Documento Word", file_types=[".docx"], type="filepath")

        analyze_btn = gr.Button("🔍 Analizar Documento", variant="primary", size="lg")

        status_msg = gr.Textbox(label="Estado del Análisis", interactive=False, show_label=True)

        gr.Markdown("---")

        # Step 2: Results
        gr.Markdown("### 📊 Paso 2: Resultados del Análisis")

        results_display = gr.HTML()

        with gr.Row():
            word_download = gr.File(label="📥 Informe Detallado (Word)", interactive=False)
            json_download = gr.File(label="📥 Datos Técnicos (JSON)", interactive=False)

        gr.Markdown("---")

        # Step 3: Expert Feedback
        gr.Markdown("### 💬 Paso 3: Evaluación Experta")
        gr.Markdown("""
        Estimado usuario, su evaluación es fundamental para mejorar el sistema.
        Por favor, indique su opinión sobre la precisión del análisis realizado.
        """)

        expert_evaluation = gr.Radio(
            choices=[
                "El análisis es correcto y preciso",
                "El análisis es mayormente correcto, con observaciones menores",
                "El análisis tiene errores significativos que debo señalar",
            ],
            label="¿Qué tan preciso fue el análisis general?",
            info="Su evaluación ayuda a mejorar futuros análisis",
        )

        classification_correct = gr.Radio(
            choices=["Sí", "No", "Parcialmente"],
            label="¿La clasificación del artículo (Científico/Divulgación) fue correcta?",
        )

        quality_score_fair = gr.Radio(
            choices=["Muy alta", "Alta", "Correcta", "Baja", "Muy baja"],
            label="¿La puntuación de calidad semántica fue justa?",
        )

        grammar_real_errors = gr.Radio(
            choices=["Todos son reales", "La mayoría son reales", "Muchos son falsos positivos"],
            label="¿Los errores gramaticales detectados fueron reales?",
        )

        structure_correct = gr.Radio(
            choices=["Sí", "No", "Parcialmente"], label="¿La validación de estructura fue correcta?"
        )

        citations_correct = gr.Radio(
            choices=["Sí", "No", "No aplica"], label="¿La detección de citas fue correcta?"
        )

        weakest_section = gr.Dropdown(
            choices=[
                "Clasificación",
                "Calidad semántica",
                "Gramática",
                "Estructura",
                "Citas y referencias",
            ],
            label="¿Qué sección fue menos útil?",
            info="Opcional",
        )

        editor_recommendation = gr.Radio(
            choices=["Sí", "No", "Con revisiones menores", "Con revisiones mayores"],
            label="¿Usted recomendaría publicar este artículo?",
        )

        expert_comments = gr.Textbox(
            label="Comentarios Adicionales (opcional)",
            placeholder="Describa cualquier error específico, aspecto que el sistema pasó por alto, o sugerencia de mejora...",
            lines=5,
        )

        feedback_btn = gr.Button("Enviar Evaluación", variant="secondary")
        feedback_status = gr.Textbox(label="", interactive=False, show_label=False)

        # Event handlers
        def analyze_and_store_name(file, progress=gr.Progress()):
            """Run analysis with step-by-step progress display"""
            import threading
            import time

            if file is None:
                return (
                    "⚠️ Seleccione un documento primero",
                    "",
                    None,
                    None,
                    "",
                    gr.Button(interactive=True),
                )

            result_container = {"result": None, "error": None, "done": False}

            def run_analysis():
                try:
                    result_container["result"] = process_document(file)
                except Exception as e:
                    result_container["error"] = str(e)
                finally:
                    result_container["done"] = True

            thread = threading.Thread(target=run_analysis)
            thread.start()

            steps = [
                (0.05, "🔧 Iniciando Silvina...", 2),
                (0.15, "[1/7] 📖 Leyendo documento...", 3),
                (0.25, "[2/7] 🔍 Extrayendo contenido...", 3),
                (0.40, "[3/7] 📚 Analizando citas y APA...", 8),
                (0.55, "[4/7] 🏷️  Clasificando artículo...", 10),
                (0.75, "[5/7] ⭐ Analizando calidad...", 45),
                (0.88, "[6/7] 📋 Validando estructura...", 3),
                (0.93, "[7/7] 🔗 Relacionando citas...", 3),
                (0.97, "💾 Generando reportes...", 3),
            ]

            for progress_val, desc, wait_seconds in steps:
                if result_container["done"]:
                    break
                progress(progress_val, desc=desc)
                elapsed = 0
                while elapsed < wait_seconds and not result_container["done"]:
                    time.sleep(0.3)
                    elapsed += 0.3

            thread.join()

            if result_container["error"]:
                return (
                    f"❌ Error: {result_container['error']}",
                    "",
                    None,
                    None,
                    "",
                    gr.Button(interactive=True),
                )

            progress(1.0, desc="✅ Análisis completado")
            status, html, word, json_path, doc_name, btn = result_container["result"]
            doc_name = word
            return status, html, word, json_path, doc_name, gr.Button(interactive=True)

        analyze_btn.click(
            fn=analyze_and_store_name,
            inputs=[file_input],
            outputs=[
                status_msg,
                results_display,
                word_download,
                json_download,
                document_name_state,
                analyze_btn,
            ],
        )

        feedback_btn.click(
            fn=save_expert_feedback,
            inputs=[
                document_name_state,
                expert_evaluation,
                classification_correct,
                quality_score_fair,
                grammar_real_errors,
                structure_correct,
                citations_correct,
                weakest_section,
                editor_recommendation,
                expert_comments,
            ],
            outputs=[feedback_status],
        )

        gr.Markdown("---")

        # Shutdown Section
        gr.Markdown("### 🔴 Cerrar Silvina")
        gr.Markdown("Cuando termine de trabajar, presione el botón para cerrar la aplicación.")

        shutdown_btn = gr.Button("🔴 Cerrar Silvina", variant="stop", size="lg")

        shutdown_msg = gr.HTML("")

        def shutdown_server():
            """Closes browser tab and stops the server."""
            import threading
            import os

            def delayed_shutdown():
                import time

                time.sleep(3)
                os._exit(0)

            threading.Thread(target=delayed_shutdown, daemon=True).start()

            return """
            <div style="background: #f8d7da; border: 2px solid #dc3545; padding: 25px;
                        border-radius: 8px; text-align: center;">
                <h3 style="color: #721c24; margin: 0 0 10px 0;">
                    🔴 Silvina se está cerrando...
                </h3>
                <p style="color: #721c24; margin: 0 0 10px 0;">
                    El servidor se cerrará en 3 segundos.
                </p>
                <p style="color: #721c24; margin: 0; font-weight: bold;">
                    ✋ Por favor cierre esta ventana haciendo clic en la <strong style="color: #721c24;">X</strong> de l pestaña de Chrome o (Ctrl+ W)
                </p>
            </div>
            """

        shutdown_btn.click(fn=shutdown_server, inputs=[], outputs=[shutdown_msg])

        # Footer
        gr.Markdown("""
        <div style="text-align: center; color: #666; font-size: 12px; padding: 20px;">
            <p><strong style="color: #666;">Silvina Editorial Assistant v0.8</strong> | Desarrollado para EUMIC</p>
        </div>
        """)

    return interface


# ============================================
# LAUNCH
# ============================================
if __name__ == "__main__":
    print("\n" + "=" * 80)
    print("   SILVINA EDITORIAL ASSISTANT v0.8 - Interfaz Web")
    print("=" * 80 + "\n")

    app = create_interface()

    print("🚀 Iniciando servidor Gradio...")
    print("📍 La aplicación estará disponible en: http://127.0.0.1:7861")
    print("⏹️  Presione Ctrl+C para detener el servidor\n")

    # Function to open Chrome after server starts
    def open_in_chrome():
        import time

        time.sleep(4)  # Wait for server to start
        chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"
        try:
            webbrowser.register("chrome", None, webbrowser.BackgroundBrowser(chrome_path))
            webbrowser.get("chrome").open("http://127.0.0.1:7861")
            print("✅ Abriendo en Google Chrome...")
        except:
            print("⚠️ Abra Chrome manualmente en: http://127.0.0.1:7861")

    # Start browser opening in background thread
    import threading

    threading.Thread(target=open_in_chrome, daemon=True).start()

    # Launch server
    silvina_reports_dir = str(Path.home() / "Documents" / "Silvina" / "reports")
    app.launch(
        server_name="127.0.0.1",
        server_port=7861,
        share=False,
        allowed_paths=[silvina_reports_dir],
        show_error=True,
        inbrowser=False,
        theme=gr.themes.Soft(primary_hue="slate"),
        css="""
        .logo-container {text-align: center; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);}
        .logo-container img {max-width: 180px; max-height: 100px;}
        .logo-container h1 {color: white; margin: 15px 0 5px 0;}
        .logo-container p {color: rgba(255,255,255,0.9); margin: 0;}
        .critical-issues-box, .critical-issues-box h4, .critical-issues-box li {color: #721c24 !important;}
        footer {display: none !important;}
        """,
    )
