from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from typing import List, Dict, Any
import json, re

TOP_KEYS = [
    "documentType", "typeOfRegistration", "serialNumber",
    "address", "competentRegistryOffice", "dateOfIssue"
]

def _normalize_structured(obj: Any) -> Dict[str, Any]:
    """dict 또는 dict 리스트를 단일 dict로 정규화."""
    if isinstance(obj, dict):
        for k in ("data", "payload", "result"):
            if isinstance(obj.get(k), dict):
                return _normalize_structured(obj[k])
        for k in ("items", "results", "list", "pages"):
            if isinstance(obj.get(k), list) and obj[k] and isinstance(obj[k][0], dict):
                return _merge_pages(obj[k])
        return dict(obj)
    if isinstance(obj, list) and obj and isinstance(obj[0], dict):
        return _merge_pages(obj)
    return {"tables": [], "remarks": []}

def _merge_pages(pages: List[Dict[str, Any]]) -> Dict[str, Any]:
    out = {k: "" for k in TOP_KEYS}
    out["tables"] = []
    out["remarks"] = []
    for p in pages:
        if not isinstance(p, dict): continue
        for k in TOP_KEYS:
            if not out.get(k) and isinstance(p.get(k), str) and p[k].strip():
                out[k] = p[k].strip()
        if isinstance(p.get("tables"), list):
            out["tables"].extend(p["tables"])
        if isinstance(p.get("remarks"), list):
            for r in p["remarks"]:
                if isinstance(r, str) and r.strip() and r not in out["remarks"]:
                    out["remarks"].append(r)
    return out

def _coerce_legacy_sections(rep: dict) -> dict:
    """partOfTitle / owner 형태를 tables 스키마로 자동 변환."""
    if rep.get("tables"):
        return rep
    tables = []

    pt = rep.get("partOfTitle") or {}
    if isinstance(pt, dict) and (pt.get("columns") or pt.get("rows")):
        header = pt.get("header") or ""
        columns = pt.get("columns") or []
        rows_in = pt.get("rows") or []
        norm_rows = []
        if rows_in and isinstance(rows_in[0], dict):
            order = ["descriptionNo","acceptance","location","buildingDetails","causeOfRegistrationAndOtherInformation"]
            for r in rows_in:
                norm_rows.append([(r.get(k, "") or "") for k in order])
        else:
            norm_rows = rows_in
        tables.append({"header": header, "columns": columns, "rows": norm_rows})

    ow = rep.get("owner") or {}
    if isinstance(ow, dict) and (ow.get("columns") or ow.get("rows")):
        header = ow.get("header") or ""
        columns = ow.get("columns") or []
        rows_in = ow.get("rows") or []
        norm_rows = []
        if rows_in and isinstance(rows_in[0], dict):
            order = ["registeredOwner","registrationNumber","finalShare","ownerAddress","priorityNumber"]
            for r in rows_in:
                norm_rows.append([(r.get(k, "") or "") for k in order])
        else:
            norm_rows = rows_in
        tables.append({"header": header, "columns": columns, "rows": norm_rows})

    if tables:
        rep["tables"] = tables
    return rep

def _normalize_header(h: str) -> str:
    """'표 제 부' → '표제부', 괄호/기호 주변 공백 정리."""
    if not isinstance(h, str):
        return ""
    s = re.sub(r"\s+", " ", h.strip())
    if re.fullmatch(r"[가-힣ㄱ-ㅎㅏ-ㅣ](?:\s[가-힣ㄱ-ㅎㅏ-ㅣ]){1,30}", s):
        s = s.replace(" ", "")
    s = s.replace("【 ", "【").replace(" 】", "】")
    s = s.replace("】 (", "】(").replace("( ", "(").replace(" )", ")")
    return s

def _gs(d: Dict[str, Any], k: str) -> str:
    v = d.get(k, "")
    return "" if v is None else str(v)

def _merge_cont_rows_on_rep(rep: Dict[str, Any]) -> Dict[str, Any]:
    """워드 출력 전 안전망 병합(연속행을 앞행 뒤에 붙이기)."""
    tables = rep.get("tables") or []
    fixed = []
    for t in tables:
        cols = max(len(t.get("columns") or []), max((len(r) for r in (t.get("rows") or [])), default=0))
        rows = [list(r) + [""]*(cols-len(r)) for r in (t.get("rows") or [])]
        out = []
        for r in rows:
            left_empty = all(not (c or "").strip() for c in r[:max(1, cols//2)])
            right_val  = any((c or "").strip() for c in r[max(1, cols//2):])
            if left_empty and right_val and out:
                prev = out[-1]
                for j in range(cols):
                    if (r[j] or "").strip():
                        prev[j] = (prev[j] + ("\n" if prev[j] else "") + r[j]).strip()
            else:
                out.append(r)
        t2 = dict(t); t2["rows"] = out; fixed.append(t2)
    rep["tables"] = fixed
    return rep

def generate_building_registry_docx(json_path: str, ocr_path: str, lang: str) -> Document:
    """
    정책:
    - 표는 오직 rep['tables']만 사용(= 구조화 결과 기반)
    - partOfTitle/owner → tables 변환
    - 헤더를 정규화해서 '표 제 부' 같은 끊김 제거
    - 연속행 병합(워드 출력 전 보정)
    - 하단 고정 블록(빈칸 알림/관할/참고/일시) 삽입
    """
    with open(json_path, encoding="utf-8") as f:
        raw_struct = json.load(f)

    rep = _normalize_structured(raw_struct)
    rep = _coerce_legacy_sections(rep)
    rep = _merge_cont_rows_on_rep(rep)  # 안전망 병합

    # 다국어 라벨
    if lang == "일본어":
        L_SN, L_DOI, L_OFFICE, L_BLANK = "固有番号", "交付日", "管轄登記所", "以下余白"
    elif lang == "중국어":
        L_SN, L_DOI, L_OFFICE, L_BLANK = "固有编号", "发证日期", "管辖登记机关", "以下为空白"
    elif lang == "베트남어":
        L_SN, L_DOI, L_OFFICE, L_BLANK = "Số định danh", "Ngày cấp", "Cơ quan đăng ký có thẩm quyền", "Phần còn lại của trang này để trống"
    else:
        L_SN, L_DOI, L_OFFICE, L_BLANK = "Serial Number", "Date Of Issue", "Competent Registry Office", "Nothing follows"

    doc = Document()

    # 상단 제목
    headings = [
        (_gs(rep, "documentType"), 16, True, WD_PARAGRAPH_ALIGNMENT.CENTER),
        (f"- {_gs(rep, 'typeOfRegistration')} -", 16, True, WD_PARAGRAPH_ALIGNMENT.CENTER),
        (f"{L_SN} {_gs(rep, 'serialNumber')}", 11, False, WD_PARAGRAPH_ALIGNMENT.RIGHT),
        (f"[{_gs(rep, 'typeOfRegistration')}] {_gs(rep, 'address')}", 11, False, WD_PARAGRAPH_ALIGNMENT.LEFT),
    ]
    for text, size, bold, align in headings:
        if not text:
            continue
        p = doc.add_paragraph()
        p.alignment = align
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.bold = bold

    # 표 렌더링
    rep_tables = rep.get("tables", []) or []
    for rep_tbl in rep_tables:
        header  = _normalize_header(rep_tbl.get("header", ""))
        columns = rep_tbl.get("columns", []) or []
        rows    = rep_tbl.get("rows", []) or []

        cols = max(len(columns), max((len(r) for r in rows), default=0))
        cols = max(cols, 1)

        table = doc.add_table(rows=0, cols=cols, style="Table Grid")

        # 큰 헤더(병합 1행)
        if header:
            hdr = table.add_row().cells
            hdr[0].text = header
            for j in range(1, cols):
                hdr[0].merge(hdr[j])
            for para in hdr[0].paragraphs:
                for run in para.runs:
                    run.font.size = Pt(12)
                    run.bold = True

        # 컬럼 헤더
        if columns:
            tr = table.add_row().cells
            for j in range(cols):
                tr[j].text = str(columns[j]) if j < len(columns) else ""
            for c in tr:
                for run in c.paragraphs[0].runs:
                    run.font.size = Pt(10)
                    run.bold = True

        # 데이터 행
        for r in rows:
            tr = table.add_row().cells
            for j in range(cols):
                tr[j].text = "" if j >= len(r) or r[j] is None else str(r[j])
            for c in tr:
                for run in c.paragraphs[0].runs:
                    run.font.size = Pt(10)

        doc.add_paragraph()  # 표 간 간격

    # === 문서 하단 텍스트 (표 아래에) 추가 ===
    p = doc.add_paragraph(f"-- {L_BLANK} --")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # 관할등기소 오른쪽 정렬 (요구사항: 고정 문자열 사용)
    p = doc.add_paragraph(f"{L_OFFICE} 서울서부지방법원 등기국")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    # 참고사항 (작은 폰트 적용, 고정 텍스트)
    notes = [
        "[ 참고사항 ]\n가. 등기기록에서 유효한 지분을 가진 소유자 혹은 공유자 현황을 표시합니다.\n나. 최종지분은 등기명의인이 가진 최종지분이며, 2개 이상의 순위번호의 지분을 가진 경우 그 지분을 합산하였습니다.\n다. 순위번호는 등기명의인을 기준으로 부여된 등기 순위번호입니다.\n라. 신청사항과 관련이 없는 소유권(갑구) 소유권 이외의 권리(을구사항)은 표시되지 않았습니다.\n마. 지분이 분봉되어 전세된 자료는 전체의 지분을 종합하여 정리한 것입니다.",
        "* 실선으로 그어진 부분은 말소사항을 표시함.    * 기재사항 없는 갑구, 을구는 '기재사항 없음'으로 표시함."
    ]
    for line in notes:
        p = doc.add_paragraph()
        run = p.add_run(line)
        run.font.size = Pt(9)

    # 하단 날짜 (요구사항: 고정 문자열 사용)
    doc.add_paragraph(f"{L_DOI} : 2025년07월17일 20시29분47초")

    return doc