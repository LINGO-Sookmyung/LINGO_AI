from docx import Document
from collections import defaultdict
import json
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt

def generate_building_registry_docx(json_path: str, ocr_path: str) -> Document:
    # OCR 결과 불러오기
    with open(ocr_path, encoding="utf-8") as f:
        data = json.load(f)

    # JSON 로드
    with open(json_path, 'r', encoding='utf-8') as f:
        raw_data = json.load(f)
        replacements = flatten_json(raw_data)

    doc = Document()  # 문서 객체 생성

    # === 문서 상단 텍스트 (표 위에) 추가 ===
    headings = [
        ("Partial Certificate of Registry (Current Ownership Status)", 16, True, WD_PARAGRAPH_ALIGNMENT.CENTER),
        (f"- Building -", 16, True, WD_PARAGRAPH_ALIGNMENT.CENTER),
        (f"Serial Number {replacements.serialNumber}", 11, False, WD_PARAGRAPH_ALIGNMENT.RIGHT),
        (f"[Building] {replacements.address}", 11, False, WD_PARAGRAPH_ALIGNMENT.LEFT)
    ]

    for text, size, bold, align in headings:
        p = doc.add_paragraph()
        p.alignment = align
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.bold = bold

    # 전체 이미지 순회
    for image in data["images"]:
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
    p = doc.add_paragraph("-- Nothing follows --")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # 관할등기소 오른쪽 정렬
    p = doc.add_paragraph(f"Competent Registry Office {서울서부지방법원 등기국}")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    # 참고사항 (작은 폰트 적용)
    notes = [
        "[ 참고사항 ]\n가. 등기기록에서 유효한 지분을 가진 소유자 혹은 공유자 현황을 표시합니다.\n나. 최종지분은 등기명의인이 가진 최종지분이며, 2개 이상의 순위번호의 지분을 가진 경우 그 지분을 합산하였습니다.\n다. 순위번호는 등기명의인을 기준으로 부여된 등기 순위번호입니다.\n라. 신청사항과 관련이 없는 소유권(갑구) 소유권 이외의 권리(을구사항)은 표시되지 않았습니다.\n마. 지분이 분봉되어 전세된 자료는 전체의 지분을 종합하여 정리한 것입니다.",
        "* 실선으로 그어진 부분은 말소사항을 표시함.    * 기재사항 없는 갑구, 을구는 '기재사항 없음'으로 표시함."
    ]
    for line in notes:
        p = doc.add_paragraph()
        run = p.add_run(line)
        run.font.size = Pt(9)

    # 하단
    doc.add_paragraph(f"Date of Issue : {"2025년07월17일 20시29분47초"}")


    # === 최종 저장 ===
    output_path = "outputs/building_registry.docx"
    doc.save(output_path)