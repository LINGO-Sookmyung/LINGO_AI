from docx import Document
from collections import defaultdict
import json
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from utils.generate_doc.flatten_json import flatten_json


def _collect_block_text(rows_2d, r0, r1, c0, c1):
    # GPT rows에서 [r0..r1], [c0..c1] 범위 텍스트 합치기
    lines = []
    R = len(rows_2d)
    for rr in range(max(0, r0), min(R, r1+1)):
        row = rows_2d[rr] if isinstance(rows_2d[rr], list) else []
        C = len(row)
        for cc in range(max(0, c0), min(C, c1+1)):
            val = row[cc]
            if isinstance(val, dict):
                val = val.get("text", "")
            lines.append("" if val is None else str(val))
    text = "\n".join(lines)
    return text if text.strip() else ""


def generate_building_registry_docx(json_path: str, ocr_path: str, lang: str) -> Document:
    # OCR 결과 불러오기
    with open(ocr_path, encoding="utf-8") as f:
        raw_ocr_data = json.load(f)
        ocr_data = raw_ocr_data[0]  # 파일 구조가 리스트라면 OK
        images = ocr_data["ocr_result"]["images"]  # 실제 테이블 정보

    # GPT 구조화 JSON 로드 (실데이터)
    with open(json_path, 'r', encoding='utf-8') as f:
        replacements = json.load(f)

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

    rep_tables = replacements.get("tables", [])  # gpt_structured_result.json 의 tables
    rep_idx = 0  # ★ OCR 테이블 ↔ GPT 테이블 순서 매칭용 인덱스

    for image in images:
        for tdata in image.get("tables", []):
            # 현재 OCR 테이블에 대응하는 GPT 테이블 선택(순서 매칭)
            rep = rep_tables[rep_idx] if rep_idx < len(rep_tables) else {}
            data_cols = rep.get("columns", [])
            data_rows = rep.get("rows", [])

            cells = tdata["cells"]

            # === 셀 분류 ===
            header_row0 = [c for c in cells if c["rowIndex"] == 0]  # 표제(큰 헤더)
            header_row1 = [c for c in cells if c["rowIndex"] == 1]  # 컬럼 헤더
            body_cells  = [c for c in cells if c["rowIndex"] >= 2]  # 본문

            # === 본문 병합 그룹 === (OCR 레이아웃 유지)
            merge_groups = defaultdict(list)
            for cell in body_cells:
                key = (cell["columnIndex"], cell.get("columnSpan", 1))
                merge_groups[key].append(cell)

            merged_regions = []
            for (col_idx, col_span), group_cells in merge_groups.items():
                min_row = min(c["rowIndex"] for c in group_cells)
                max_row = max(c["rowIndex"] + c.get("rowSpan", 1) - 1 for c in group_cells)

                # replacements.rows의 인덱스는 OCR rowIndex - 2 가 대응
                start_r = max(0, min_row - 2)
                end_r   = max(0, max_row - 2)
                start_c = col_idx
                end_c   = col_idx + col_span - 1

                text_block = _collect_block_text(data_rows, start_r, end_r, start_c, end_c)
                merged_regions.append((min_row, max_row, col_idx, col_idx + col_span - 1, text_block))

            # === 테이블 크기 계산 ===
            max_row = max([c["rowIndex"] + c.get("rowSpan", 1) for c in cells])
            max_col = max([c["columnIndex"] + c.get("columnSpan", 1) for c in cells])

            # === docx 테이블 생성 ===
            doc_table = doc.add_table(rows=max_row, cols=max_col, style="Table Grid")

            # === row 0: 표제(헤더) = GPT header로 채움 (레이아웃 병합은 OCR대로) ===
            if header_row0:
                leftmost = min(header_row0, key=lambda x: x["columnIndex"])
                r, c = leftmost["rowIndex"], leftmost["columnIndex"]
                total_span = sum(c0.get("columnSpan", 1) for c0 in header_row0)
                tgt_cell = doc_table.cell(r, c)
                if total_span > 1:
                    tgt_cell = tgt_cell.merge(doc_table.cell(r, c + total_span - 1))
                tgt_cell.text = rep.get("header", "")  # ★ 헤더는 GPT 값만 사용
                for para in tgt_cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(12)
                        run.bold = True

            # === row 1: 컬럼 헤더 = GPT columns 순서대로 ===
            for cinfo in header_row1:
                r, c = cinfo["rowIndex"], cinfo["columnIndex"]
                tgt_cell = doc_table.cell(r, c)
                tgt_cell.text = str(data_cols[c]) if 0 <= c < len(data_cols) else ""
                for para in tgt_cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)

            # === 본문: OCR 병합영역에 GPT rows 매핑 ===
            for start_row, end_row, start_col, end_col, text in merged_regions:
                tgt_cell = doc_table.cell(start_row, start_col)
                if end_row > start_row or end_col > start_col:
                    tgt_cell = tgt_cell.merge(doc_table.cell(end_row, end_col))
                tgt_cell.text = text
                for para in tgt_cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)

            doc.add_paragraph()  # 테이블 간 간격
            rep_idx += 1  # ★ 다음 OCR 테이블은 다음 GPT 테이블과 매칭

    # === 문서 하단 텍스트 (표 아래에) 추가 ===
    p = doc.add_paragraph(f"-- {blank_notice_label} --")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # 관할등기소 오른쪽 정렬
    p = doc.add_paragraph(f"{competent_registry_Office_label} {replacements['competentRegistryOffice']}")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    remarks_text = '\n'.join(replacements['remarks'])
    p = doc.add_paragraph()
    run = p.add_run(remarks_text)
    run.font.size = Pt(9)

    doc.add_paragraph(f"{date_of_issue_label} : {replacements['dateOfIssue']}")

    return doc
