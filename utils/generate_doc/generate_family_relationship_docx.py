from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import json
from utils.generate_doc.flatten_json import flatten_json

# 열 너비 설정 함수
def set_column_widths(table, widths):
    for row in table.rows:
        for idx, width in enumerate(widths):
            tc = row.cells[idx]._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:type'), 'dxa')
            tcW.set(qn('w:w'), str(width))  # width in twentieths of a point (1 inch = 1440)
            tcPr.append(tcW)

# 등록기준지 열 너비 조정
def set_reg_table_widths(table):
    widths = [2500, 7000]  # 왼쪽 좁게, 오른쪽 넓게
    for row in table.rows:
        for idx, width in enumerate(widths):
            tc = row.cells[idx]._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:type'), 'dxa')
            tcW.set(qn('w:w'), str(width))
            tcPr.append(tcW)

# 가족사항 라벨 열 너비 축소
def set_label_table_width(table, width=2000):
    cell = table.cell(0, 0)._tc
    tcPr = cell.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:type'), 'dxa')
    tcW.set(qn('w:w'), str(width))
    tcPr.append(tcW)


def generate_family_relationship_docx(json_path: str, lang: str) -> Document:
    doc = Document()

    # 제목
    title = doc.add_paragraph(data["documentType"])
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.size = Pt(16)
    title.runs[0].bold = True

    # 등록기준지
    note = doc.add_paragraph("[Consulate General of the Republic of Korea in New York]")
    note.paragraph_format.space_before = Pt(0)
    note.paragraph_format.space_after = Pt(0)
    reg_table = doc.add_table(rows=1, cols=2)
    reg_table.style = 'Table Grid'
    reg_table.cell(0, 0).text = "Original Domicile"
    reg_table.cell(0, 1).text = data["registrationBase"]
    for cell in reg_table.row_cells(0):
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_reg_table_widths(reg_table)  # 🔧 열 너비 조정

    # 본인 정보
    space = doc.add_paragraph()
    space.paragraph_format.space_before = Pt(0)
    space.paragraph_format.space_after = Pt(0)
    table = doc.add_table(rows=2, cols=6)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = "Category"
    hdr[1].text = "Name"
    hdr[2].text = "Date of Birth"
    hdr[3].text = "Resident Registration No."
    hdr[4].text = "Sex"
    hdr[5].text = "Origin of Surname"
    row = table.rows[1].cells
    row[0].text = "본인"
    row[1].text = data["person"]["name"]
    row[2].text = data["person"]["birthDate"]
    row[3].text = data["person"]["residentRegistrationNumber"]
    row[4].text = data["person"]["gender"]
    row[5].text = data["person"]["origin"]
    # 가운데 정렬
    for row in table.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # 열 너비 조정 (단위: 1/20 pt → 1인치 = 1440)
    set_column_widths(table, [1100, 2500, 2500, 3000, 900, 1000])

    # 가족사항 라벨
    space = doc.add_paragraph()
    space.paragraph_format.space_before = Pt(0)
    space.paragraph_format.space_after = Pt(0)
    label_table = doc.add_table(rows=1, cols=1)
    label_table.style = 'Table Grid'
    label_cell = label_table.cell(0, 0)
    label_cell.text = "Family Details"
    label_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_label_table_width(label_table)  # 🔧 너비 축소

    # 가족사항 상세 표
    space = doc.add_paragraph()
    space.paragraph_format.space_before = Pt(0)
    space.paragraph_format.space_after = Pt(0)
    # 1. 가족 목록 분리
    parents = [m for m in data["family"] if m["relation"] in ["부", "모"]]
    spouse = [m for m in data["family"] if m["relation"] == "배우자"]
    children = [m for m in data["family"] if "자녀" in m["relation"] or m["relation"] == "자"]

    # 2. 부모 테이블 (헤더 포함)
    fam_table = doc.add_table(rows=len(parents) + 1, cols=6)
    fam_table.style = 'Table Grid'
    fam_hdr = fam_table.rows[0].cells
    fam_hdr[0].text = "Category"
    fam_hdr[1].text = "Name"
    fam_hdr[2].text = "Date of Birth"
    fam_hdr[3].text = "Resident Registration No."
    fam_hdr[4].text = "Sex"
    fam_hdr[5].text = "Origin of Surname"
    for i, member in enumerate(parents):
        r = fam_table.rows[i + 1].cells
        r[0].text = member["relation"]
        r[1].text = member["name"]
        r[2].text = member["birthDate"]
        r[3].text = member["residentRegistrationNumber"]
        r[4].text = member["gender"]
        r[5].text = member["origin"]
    for row in fam_table.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_column_widths(fam_table, [1100, 2500, 2500, 3000, 900, 1000])

    # 3. 배우자 테이블 (헤더 없음, 간격 추가)
    if spouse:
        space = doc.add_paragraph()
        space.paragraph_format.space_before = Pt(0)
        space.paragraph_format.space_after = Pt(0)
        spouse_table = doc.add_table(rows=len(spouse), cols=6)
        spouse_table.style = 'Table Grid'
        for i, member in enumerate(spouse):
            r = spouse_table.rows[i].cells
            r[0].text = member["relation"]
            r[1].text = member["name"]
            r[2].text = member["birthDate"]
            r[3].text = member["residentRegistrationNumber"]
            r[4].text = member["gender"]
            r[5].text = member["origin"]
        for row in spouse_table.rows:
            for cell in row.cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_column_widths(spouse_table, [1100, 2500, 2500, 3000, 900, 1000])

    # 4. 자녀 테이블 (헤더 없음, 간격 추가)
    if children:
        space = doc.add_paragraph()
        space.paragraph_format.space_before = Pt(0)
        space.paragraph_format.space_after = Pt(0)
        child_table = doc.add_table(rows=len(children), cols=6)
        child_table.style = 'Table Grid'
        for i, member in enumerate(children):
            r = child_table.rows[i].cells
            r[0].text = member["relation"]
            r[1].text = member["name"]
            r[2].text = member["birthDate"]
            r[3].text = member["residentRegistrationNumber"]
            r[4].text = member["gender"]
            r[5].text = member["origin"]
        for row in child_table.rows:
            for cell in row.cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_column_widths(child_table, [1100, 2500, 2500, 3000, 900, 1000])

    # 문구 및 발급일
    doc.add_paragraph()
    note1 = doc.add_paragraph(data["notes"][0])
    note1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    issuedDate = doc.add_paragraph()
    run = issuedDate.add_run(data["issuedDate"])
    run.font.size = Pt(12)
    issuedDate.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 인증 기관 정보
    org = doc.add_paragraph()
    run = org.add_run(f'{data["certifiedBy"]["organization"]} {data["certifiedBy"]["responsible"]}')
    run.bold = True
    run.font.size = Pt(13)
    org.alignment = WD_ALIGN_PARAGRAPH.CENTER

    note2 = doc.add_paragraph(data["notes"][1])
    note2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    note2.paragraph_format.line_spacing = 1

    # 발급정보
    doc.add_paragraph()
    issuedTime = doc.add_paragraph(f'발급시간 : {data["issuedTime"]}')
    applicant = doc.add_paragraph(f'신청인 : {data["applicant"]}')
    issuedTime.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    issuedTime.paragraph_format.space_after = Pt(0)
    applicant.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    applicant.paragraph_format.space_before = Pt(0)

    # 발행번호
    doc.add_paragraph()
    certificateNumber = doc.add_paragraph()
    run = certificateNumber.add_run(f'발행번호 : {data["certificateNumber"]}')
    run.font.size = Pt(10)
    certificateNumber.alignment = WD_ALIGN_PARAGRAPH.LEFT
    certificateNumber.paragraph_format.space_after = Pt(0)

    # 주석
    note3 = doc.add_paragraph()
    run = note3.add_run(data["notes"][2])
    run.font.size = Pt(10)
    note3.alignment = WD_ALIGN_PARAGRAPH.LEFT
    note3.paragraph_format.line_spacing = 1
    note2.paragraph_format.space_before = Pt(0)

    return doc