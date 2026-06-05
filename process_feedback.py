"""
process_feedback.py
Processes Silvina v0.8 expert feedback files and generates:
1. Human-readable summary report (Markdown)
2. Development prompt for v0.9 vibe coding

Usage:
    python process_feedback.py --folder feedback_received/
    python process_feedback.py --folder feedback_received/ --output summary.md
"""

import json
import argparse
from pathlib import Path
from datetime import datetime
from collections import Counter


def load_feedback_files(folder: Path) -> list:
    files = list(folder.glob("*_feedback.json"))
    if not files:
        print(f"No feedback files found in {folder}")
        return []
    print(f"Found {len(files)} feedback files")
    records = []
    for f in files:
        try:
            with open(f, encoding='utf-8') as fh:
                data = json.load(fh)
                data['_file'] = f.name
                records.append(data)
                print(f"  Loaded: {f.name}")
        except Exception as e:
            print(f"  Could not read {f.name}: {e}")
    return records


def analyze_feedback(records: list) -> dict:
    n = len(records)
    return {
        'n': n,
        'overall': Counter(r.get('overall_precision', 'Sin respuesta') for r in records),
        'classification': Counter(r.get('classification_correct', 'Sin respuesta') for r in records),
        'quality_score': Counter(r.get('quality_score_fair', 'Sin respuesta') for r in records),
        'grammar': Counter(r.get('grammar_real_errors', 'Sin respuesta') for r in records),
        'structure': Counter(r.get('structure_correct', 'Sin respuesta') for r in records),
        'citations': Counter(r.get('citations_correct', 'Sin respuesta') for r in records),
        'weakest': Counter(r.get('weakest_section', 'Sin respuesta') for r in records),
        'recommendation': Counter(r.get('editor_recommendation', 'Sin respuesta') for r in records),
        'comments': [
            r.get('comments', '')
            for r in records
            if r.get('comments') and r.get('comments') != 'Sin comentarios adicionales'
        ],
        'records': records
    }


def identify_issues(data: dict) -> list:
    n = data['n']
    issues = []

    fp = data['grammar'].get('Muchos son falsos positivos', 0)
    if fp > 0:
        issues.append({
            'count': fp, 'pct': round(fp/n*100),
            'component': 'Gramatica',
            'issue': 'Falsos positivos en deteccion de errores gramaticales',
            'action': 'Ampliar whitelist de nombres propios, siglas y terminos tecnicos militares'
        })

    wrong = data['classification'].get('No', 0) + data['classification'].get('Parcialmente', 0)
    if wrong > 0:
        issues.append({
            'count': wrong, 'pct': round(wrong/n*100),
            'component': 'Clasificacion',
            'issue': 'Clasificacion Cientifico/Divulgacion incorrecta o parcial',
            'action': 'Revisar criterios de clasificacion y umbrales de confianza'
        })

    wrong_struct = data['structure'].get('No', 0) + data['structure'].get('Parcialmente', 0)
    if wrong_struct > 0:
        issues.append({
            'count': wrong_struct, 'pct': round(wrong_struct/n*100),
            'component': 'Estructura',
            'issue': 'Validacion de estructura incorrecta o parcial',
            'action': 'Mejorar deteccion de secciones con formato no estandar'
        })

    too_high = data['quality_score'].get('Muy alta', 0) + data['quality_score'].get('Alta', 0)
    too_low = data['quality_score'].get('Muy baja', 0) + data['quality_score'].get('Baja', 0)
    if too_high > 0 or too_low > 0:
        direction = 'infladas' if too_high > too_low else 'defladas'
        issues.append({
            'count': too_high + too_low, 'pct': round((too_high + too_low)/n*100),
            'component': 'Calidad semantica',
            'issue': f'Puntuaciones de calidad semantica {direction}',
            'action': 'Recalibrar prompts LLM'
        })

    wrong_cit = data['citations'].get('No', 0)
    if wrong_cit > 0:
        issues.append({
            'count': wrong_cit, 'pct': round(wrong_cit/n*100),
            'component': 'Citas y referencias',
            'issue': 'Deteccion de citas incorrecta',
            'action': 'Revisar parser de citas - posible problema con formato Fuente: o notas al pie'
        })

    issues.sort(key=lambda x: x['count'], reverse=True)
    return issues


def generate_summary(data: dict, issues: list) -> str:
    n = data['n']
    lines = []
    lines.append("# Silvina v0.8 - Resumen de Evaluaciones Expertas")
    lines.append(f"**Fecha:** {datetime.now().strftime('%d/%m/%Y')}")
    lines.append(f"**Total de evaluaciones procesadas:** {n}\n")
    lines.append("---\n")

    lines.append("## Evaluacion General\n")
    for k, v in data['overall'].most_common():
        lines.append(f"- {k}: **{v}/{n}** ({round(v/n*100)}%)")

    lines.append("\n## Recomendacion de Publicacion (Editores)\n")
    for k, v in data['recommendation'].most_common():
        if k != 'Sin respuesta':
            lines.append(f"- {k}: **{v}/{n}** ({round(v/n*100)}%)")

    lines.append("\n## Problemas Identificados (por frecuencia)\n")
    if not issues:
        lines.append("No se identificaron problemas sistematicos.")
    else:
        for i, issue in enumerate(issues, 1):
            priority = "CRITICO" if issue['pct'] >= 50 else "MODERADO" if issue['pct'] >= 25 else "MENOR"
            lines.append(f"### {i}. {issue['component']} - {priority}")
            lines.append(f"- **Problema:** {issue['issue']}")
            lines.append(f"- **Frecuencia:** {issue['count']}/{n} evaluaciones ({issue['pct']}%)")
            lines.append(f"- **Accion recomendada:** {issue['action']}\n")

    lines.append("## Seccion Menos Util (segun editores)\n")
    for k, v in data['weakest'].most_common():
        if k != 'Sin respuesta':
            lines.append(f"- {k}: {v} menciones")

    if data['comments']:
        lines.append("\n## Comentarios Textuales de Editores\n")
        for i, c in enumerate(data['comments'], 1):
            lines.append(f"{i}. \"{c}\"")

    lines.append("\n---")
    lines.append("*Generado automaticamente por process_feedback.py*")
    return '\n'.join(lines)


def generate_dev_prompt(data: dict, issues: list) -> str:
    lines = []
    lines.append("# Silvina v0.9 - Development Prompt")
    lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d')}")
    lines.append(f"Based on: {data['n']} expert feedback submissions\n")
    lines.append("## Priority Fixes\n")

    for i, issue in enumerate(issues, 1):
        priority = "HIGH" if issue['pct'] >= 50 else "MEDIUM" if issue['pct'] >= 25 else "LOW"
        lines.append(f"### Fix {i} [{priority}] - {issue['component']} ({issue['count']} cases, {issue['pct']}%)")
        lines.append(f"**Issue:** {issue['issue']}")
        lines.append(f"**Action:** {issue['action']}\n")

    if data['comments']:
        lines.append("## Editor Comments to Address\n")
        for i, c in enumerate(data['comments'], 1):
            lines.append(f"{i}. {c}")

    lines.append("\n## Weakest Section Ranking\n")
    for section, count in data['weakest'].most_common():
        if section != 'Sin respuesta':
            lines.append(f"- {section}: {count} mentions")

    return '\n'.join(lines)


def main():
    parser = argparse.ArgumentParser(description='Process Silvina expert feedback files')
    parser.add_argument('--folder', type=str, default='feedback_received',
                        help='Folder containing _feedback.json files')
    parser.add_argument('--output', type=str, default=None,
                        help='Output filename for summary')
    args = parser.parse_args()

    folder = Path(args.folder)
    if not folder.exists():
        print(f"Folder not found: {folder}")
        return

    print(f"\nProcessing feedback from: {folder}")
    records = load_feedback_files(folder)
    if not records:
        return

    data = analyze_feedback(records)
    issues = identify_issues(data)

    summary = generate_summary(data, issues)
    summary_file = Path(args.output) if args.output else Path(f"feedback_summary_{datetime.now().strftime('%Y%m%d')}.md")
    summary_file.write_text(summary, encoding='utf-8')
    print(f"Summary saved: {summary_file}")

    dev_prompt = generate_dev_prompt(data, issues)
    prompt_file = Path(f"v09_dev_prompt_{datetime.now().strftime('%Y%m%d')}.md")
    prompt_file.write_text(dev_prompt, encoding='utf-8')
    print(f"Dev prompt saved: {prompt_file}")

    print(f"\n{'='*50}")
    print(f"ISSUES FOUND: {len(issues)}")
    for issue in issues:
        print(f"  [{issue['pct']}%] {issue['component']}: {issue['issue']}")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()
