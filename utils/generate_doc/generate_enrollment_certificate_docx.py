from docx import Document
import json
from utils.generate_doc.flatten_json import flatten_json

# === 스타일 유지하며 run 내부 치환 ===
def replace_in_runs(paragraphs, replacements):
    for para in paragraphs:
        buffer = ""
        new_runs = []
        style_info = []

        # 1. 기존 run 정보 수집
        for run in para.runs:
            buffer += run.text
            style_info.append({
                "style": run.style,
                "size": run.font.size,
                "bold": run.bold,
                "italic": run.italic,
                "underline": run.underline
            })
            new_runs.append(run)

        # 2. {{key}} -> value 치환
        for key, val in replacements.items():
            buffer = buffer.replace(f"{{{{{key}}}}}", str(val))  # {{}} escape

        # 3. 기존 run 제거
        for run in new_runs:
            para._element.remove(run._element)

        # 4. 새 run 삽입 (첫 run 스타일 유지)
        if style_info:
            run = para.add_run(buffer)
            run.style = style_info[0]["style"]
            run.font.size = style_info[0]["size"]
            run.bold = style_info[0]["bold"]
            run.italic = style_info[0]["italic"]
            run.underline = style_info[0]["underline"]
        else:
            para.add_run(buffer)

def generate_enrollment_certificate_docx(json_path: str, lang: str) -> Document:
    # JSON 로드
    with open(json_path, 'r', encoding='utf-8') as f:
        raw_data = json.load(f)
        replacements = flatten_json(raw_data)

    # 문서 열기
    doc = Document("sample/template_enrollment_certificate.docx")

    # 본문
    replace_in_runs(doc.paragraphs)

    # 머리말
    for section in doc.sections:
        replace_in_runs(section.header.paragraphs)

    return doc