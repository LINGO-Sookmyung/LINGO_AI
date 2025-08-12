from docx import Document
import json
import re

def has_drawing(run):
    """run 안에 도형(w:drawing, w:pict)이 있는지 확인"""
    return bool(run._element.xpath('.//w:drawing') or run._element.xpath('.//w:pict'))

def replace_in_runs(paragraphs, replacements):
    for para in paragraphs:
        if not para.runs:
            continue

        # 도형이 있는 run과 없는 run을 분리
        text_runs = []
        drawing_runs = []
        for run in para.runs:
            if has_drawing(run):
                drawing_runs.append(run)
            else:
                text_runs.append(run)

        # 텍스트 run들을 합쳐서 치환
        buffer = "".join(r.text for r in text_runs)
        for key, val in replacements.items():
            buffer = re.sub(
                r"\{\{\s*" + re.escape(key) + r"\s*\}\}",  # {{ key }} 공백 허용
                str(val),
                buffer
            )

        # 기존 텍스트 run 제거
        for run in text_runs:
            para._element.remove(run._element)

        # 새 텍스트 run 삽입 (첫 run 스타일 유지)
        if text_runs:
            first_style = text_runs[0]
            new_run = para.add_run(buffer)
            new_run.style = first_style.style
            new_run.font.size = first_style.font.size
            new_run.bold = first_style.bold
            new_run.italic = first_style.italic
            new_run.underline = first_style.underline

        # 도형 run은 원래 순서 유지하면서 다시 붙이기
        for run in drawing_runs:
            para._element.append(run._element)

def generate_enrollment_certificate_docx(json_path: str, lang: str) -> Document:
    # JSON 로드
    with open(json_path, 'r', encoding='utf-8') as f:
        replacements = json.load(f)

    if lang=="일본어":
        doc = Document("templates/japanese_template_enrollment_certificate.docx")
    elif lang=="중국어":
        doc = Document("templates/chinese_template_enrollment_certificate.docx")
    elif lang=="베트남어":
        doc = Document("templates/vietnamese_template_enrollment_certificate.docx")
    else:
        doc = Document("templates/english_template_enrollment_certificate.docx")
    # 본문
    replace_in_runs(doc.paragraphs, replacements)

    # 머리말
    for section in doc.sections:
        replace_in_runs(section.header.paragraphs, replacements)

    return doc