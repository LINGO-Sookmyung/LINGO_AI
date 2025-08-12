from docx import Document
from collections import defaultdict
import json
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from utils.generate_doc.flatten_json import flatten_json

def generate_building_registry_docx(json_path: str, ocr_path: str, lang: str) -> Document:
    # OCR 결과 불러오기
    with open(ocr_path, encoding="utf-8") as f:
        raw_ocr_data = json.load(f)
        ocr_data = raw_ocr_data[0]  # 첫 번째 항목 선택
        images = ocr_data["ocr_result"]["images"]  # 실제 테이블 정보

    # JSON 로드
    with open(json_path, 'r', encoding='utf-8') as f:
        raw_data = json.load(f)
        replacements = flatten_json(raw_data)

    doc = Document()  # 문서 객체 생성

    if lang=="일본어":
        serial_number_label = "固有番号"
        date_of_issue_label = "交付日"
        competent_registry_Office_label = "管轄登記所"
        blank_notice_label = "以下余白"
    elif lang=="중국어":
        serial_number_label = "固有番号"
        date_of_issue_label = "发证日期"
        competent_registry_Office_label = "管辖登记机关"
        blank_notice_label = "以下为空白"
    elif lang=="베트남어":
        serial_number_label = "Số định danh"
        date_of_issue_label = "Ngày cấp"
        competent_registry_Office_label = "Cơ quan đăng ký có thẩm quyền"
        blank_notice_label = "Phần còn lại của trang này để trống"
    else:
        serial_number_label = "Serial Number"
        date_of_issue_label = "Date Of Issue"
        competent_registry_Office_label = "Competent Registry Office"
        blank_notice_label = "Nothing follows"

    print(replacements)  
    # === 문서 상단 텍스트 (표 위에) 추가 ===
    headings = [
        (f"{replacements['documentType']}", 16, True, WD_PARAGRAPH_ALIGNMENT.CENTER),
        (f"- {replacements['typeOfRegistration']} -", 16, True, WD_PARAGRAPH_ALIGNMENT.CENTER),
        (f"{serial_number_label} {replacements['serialNumber']}", 11, False, WD_PARAGRAPH_ALIGNMENT.RIGHT),
        (f"[{replacements['typeOfRegistration']}] {replacements['address']}", 11, False, WD_PARAGRAPH_ALIGNMENT.LEFT)
    ]

    for text, size, bold, align in headings:
        p = doc.add_paragraph()
        p.alignment = align
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.bold = bold

    # 전체 이미지 순회
    for image in images:
        for table_data in image.get("tables", []):
            cells = table_data["cells"]

            # === 셀 분류 ===
            header_row0 = [c for c in cells if c["rowIndex"] == 0]
            header_row1 = [c for c in cells if c["rowIndex"] == 1]
            body_cells  = [c for c in cells if c["rowIndex"] >= 2]

            # === 본문 병합 그룹 ===
            merge_groups = defaultdict(list)
            for cell in body_cells:
                key = (cell["columnIndex"], cell["columnSpan"])
                merge_groups[key].append(cell)

            merged_regions = []
            for (col_idx, col_span), group_cells in merge_groups.items():
                min_row = min(c["rowIndex"] for c in group_cells)
                max_row = max(c["rowIndex"] + c["rowSpan"] - 1 for c in group_cells)

                text_lines = []
                for c in sorted(group_cells, key=lambda x: x["rowIndex"]):
                    for line in c.get("cellTextLines", []):
                        text = ''.join(w["inferText"] for w in line.get("cellWords", []))
                        text_lines.append(text)

                merged_regions.append((min_row, max_row, col_idx, col_idx + col_span - 1, text_lines))

            # === 테이블 크기 ===
            max_row = max([c["rowIndex"] + c.get("rowSpan", 1) for c in cells])
            max_col = max([c["columnIndex"] + c.get("columnSpan", 1) for c in cells])

            # === 테이블 추가 ===
            table = doc.add_table(rows=max_row, cols=max_col, style="Table Grid")

            # === row 0: 가로 병합 ===
            for cell in header_row0:
                r, c = cell["rowIndex"], cell["columnIndex"]
                span = cell.get("columnSpan", 1)
                tgt_cell = table.cell(r, c)
                if span > 1:
                    tgt_cell = tgt_cell.merge(table.cell(r, c + span - 1))

                lines = []
                for line in cell.get("cellTextLines", []):
                    text = ''.join(w["inferText"] for w in line.get("cellWords", []))
                    lines.append(text)
                tgt_cell.text = '\n'.join(lines)

            # 스타일 적용: 12pt + bold
            for para in tgt_cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(12)
                    run.bold = True


            # === row 1: 텍스트만 ===
            for cell in header_row1:
                r, c = cell["rowIndex"], cell["columnIndex"]
                lines = []
                for line in cell.get("cellTextLines", []):
                    text = ''.join(w["inferText"] for w in line.get("cellWords", []))
                    lines.append(text)
                tgt_cell = table.cell(r, c)
                tgt_cell.text = '\n'.join(lines)

                # 스타일 적용: 10pt, 일반
                for para in tgt_cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)


            # === 본문 병합 + 텍스트 ===
            for start_row, end_row, start_col, end_col, lines in merged_regions:
                tgt_cell = table.cell(start_row, start_col)
                if end_row > start_row or end_col > start_col:
                    tgt_cell = tgt_cell.merge(table.cell(end_row, end_col))
                tgt_cell.text = '\n'.join(lines)

                # 폰트 스타일 적용: 10pt
                for para in tgt_cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)

            doc.add_paragraph()  # 테이블 간 간격

    # === 문서 하단 텍스트 (표 아래에) 추가 ===
    p = doc.add_paragraph(f"-- {blank_notice_label} --")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # 관할등기소 오른쪽 정렬
    p = doc.add_paragraph(f"{competent_registry_Office_label} {replacements['competentRegistryOffice']}")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    remarks_text = '\n'.join(replacements['remarks'])  # 한 문장으로 결합
    p = doc.add_paragraph()
    run = p.add_run(remarks_text)
    run.font.size = Pt(9)

    # 하단
    doc.add_paragraph(f"{date_of_issue_label} : {replacements['dateOfIssue']}")

    return doc