from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import json

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
    # JSON 로드
    with open(json_path, 'r', encoding='utf-8') as f:
        replacements = json.load(f)

    doc = Document()

    if lang == "일본어":
        original_domicile = "本籍地"  # 등록기준지
        family_detail = "家族事項"    # 가족사항
        time_of_issue_label = "発行日"  # 발급일자
        applicant_label = "申請者"      # 신청인
        certificate_number_label = "証明書番号"  # 증명서 번호
    elif lang == "중국어":
        original_domicile = "户籍所在地"  # 등록기준지
        family_detail = "家庭情况"       # 가족사항
        time_of_issue_label = "签发日期"   # 발급일자
        applicant_label = "申请人"         # 신청인
        certificate_number_label = "证书编号"  # 증명서 번호
    elif lang == "베트남어":
        original_domicile = "Nơi đăng ký hộ tịch"  # 등록기준지
        family_detail = "Thông tin gia đình"      # 가족사항
        time_of_issue_label = "Ngày cấp"          # 발급일자
        applicant_label = "Người nộp đơn"         # 신청인
        certificate_number_label = "Số giấy chứng nhận"  # 증명서 번호
    else:  # 기본: 영어
        original_domicile = "Registered Domicile"
        family_detail = "Family Details"
        time_of_issue_label = "Time of Issue"
        applicant_label = "Applicant"
        certificate_number_label = "Certificate Number"

    # 제목
    title = doc.add_paragraph(replacements["documentType"])
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.size = Pt(16)
    title.runs[0].bold = True

    # 등록기준지
    space = doc.add_paragraph()
    space.paragraph_format.space_before = Pt(0)
    space.paragraph_format.space_after = Pt(0)
    reg_table = doc.add_table(rows=1, cols=2)
    reg_table.style = 'Table Grid'
    reg_table.cell(0, 0).text = original_domicile
    reg_table.cell(0, 1).text = replacements["placeOfFamilyRegistration"]
    for cell in reg_table.row_cells(0):
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_reg_table_widths(reg_table)  # 🔧 열 너비 조정

    # 본인 정보
    space = doc.add_paragraph()
    space.paragraph_format.space_before = Pt(0)
    space.paragraph_format.space_after = Pt(0)
    table = doc.add_table(rows=2, cols=6)
    table.style = 'Table Grid'
    # columns 값 가져오기
    columns = replacements["columns"]
    hdr = table.rows[0].cells
    for idx, col_name in enumerate(columns):
        if idx < len(hdr):  # 안전 체크
            hdr[idx].text = col_name
    row = table.rows[1].cells
    row[0].text = replacements["registrant"]["category"]
    row[1].text = replacements["registrant"]["fullName"]
    row[2].text = replacements["registrant"]["dateOfBirth"]
    row[3].text = replacements["registrant"]["residentRegistrationNumber"]
    row[4].text = replacements["registrant"]["sex"]
    row[5].text = replacements["registrant"]["originOfSurname"]
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
    label_cell.text = family_detail
    label_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_label_table_width(label_table)  # 🔧 너비 축소

    # 가족사항 상세 표
    space = doc.add_paragraph()
    space.paragraph_format.space_before = Pt(0)
    space.paragraph_format.space_after = Pt(0)
    # 1. 가족 목록 분리
    parents = [m for m in replacements["familyMembers"] if m["category"] in ["Father", "Mother","父","母","父亲","母亲","Cha","Mẹ"]]
    spouse = [m for m in replacements["familyMembers"] if m["category"] in ["Spouse","配偶者","配偶", "Người phối ngẫu"]]
    children = [m for m in replacements["familyMembers"] if m["category"] in ["Children","子女","子","Con"]]

    # 2. 부모 테이블 (헤더 포함)
    fam_table = doc.add_table(rows=len(parents) + 1, cols=6)
    fam_table.style = 'Table Grid'
    # columns 값 가져오기
    columns = replacements["columns"]
    fam_hdr = fam_table.rows[0].cells
    for idx, col_name in enumerate(columns):
        if idx < len(fam_hdr):  # 안전 체크
            fam_hdr[idx].text = col_name
    for i, member in enumerate(parents):
        r = fam_table.rows[i + 1].cells
        r[0].text = member["category"]
        r[1].text = member["fullName"]
        r[2].text = member["dateOfBirth"]
        r[3].text = member["residentRegistrationNumber"]
        r[4].text = member["sex"]
        r[5].text = member["originOfSurname"]
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
            r[0].text = member["category"]
            r[1].text = member["fullName"]
            r[2].text = member["dateOfBirth"]
            r[3].text = member["residentRegistrationNumber"]
            r[4].text = member["sex"]
            r[5].text = member["originOfSurname"]
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
            r[0].text = member["category"]
            r[1].text = member["fullName"]
            r[2].text = member["dateOfBirth"]
            r[3].text = member["residentRegistrationNumber"]
            r[4].text = member["sex"]
            r[5].text = member["originOfSurname"]
        for row in child_table.rows:
            for cell in row.cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_column_widths(child_table, [1100, 2500, 2500, 3000, 900, 1000])

    # 문구 및 발급일
    doc.add_paragraph()
    note1 = doc.add_paragraph(replacements["remarks"][0])
    note1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    issuedDate = doc.add_paragraph()
    run = issuedDate.add_run(replacements["dateOfIssue"])
    run.font.size = Pt(12)
    issuedDate.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 인증 기관 정보
    org = doc.add_paragraph()
    run = org.add_run(f'{replacements["issuingAuthority"]["organization"]} {replacements["issuingAuthority"]["authorizedOfficer"]}')
    run.bold = True
    run.font.size = Pt(13)
    org.alignment = WD_ALIGN_PARAGRAPH.CENTER

    note2 = doc.add_paragraph(replacements["remarks"][1])
    note2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    note2.paragraph_format.line_spacing = 1

    # 발급정보
    doc.add_paragraph()
    issuedTime = doc.add_paragraph(f'{time_of_issue_label} : {replacements["timeOfIssue"]}')
    applicant = doc.add_paragraph(f'{applicant_label} : {replacements["applicant"]}')
    issuedTime.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    issuedTime.paragraph_format.space_after = Pt(0)
    applicant.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    applicant.paragraph_format.space_before = Pt(0)

    # 발행번호
    doc.add_paragraph()
    certificateNumber = doc.add_paragraph()
    run = certificateNumber.add_run(f'{certificate_number_label} : {replacements["certificateNumber"]}')
    run.font.size = Pt(10)
    certificateNumber.alignment = WD_ALIGN_PARAGRAPH.LEFT
    certificateNumber.paragraph_format.space_after = Pt(0)

    # 주석
    note3 = doc.add_paragraph()
    run = note3.add_run(replacements["remarks"][2])
    run.font.size = Pt(10)
    note3.alignment = WD_ALIGN_PARAGRAPH.LEFT
    note3.paragraph_format.line_spacing = 1
    note2.paragraph_format.space_before = Pt(0)

    return doc