import json
import os
from typing import List, Dict, Any
from dotenv import load_dotenv
import openai

load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY") or os.getenv("GPT-API-KEY")


def _cell_text(cell: Dict[str, Any]) -> str:
    """
    셀 내 텍스트를 '사람이 보는 그대로' 복원한다.
    우선순위:
      1) cellTextLines[*].text (있으면 그대로 사용; 빈 문자열도 줄로 취급)
      2) cellTextLines[*].cellWords[*].inferText (단어들을 공백으로 join)
      3) (위가 모두 비면) 루트의 cellWords[*].inferText
      4) 그래도 없으면 빈 문자열 한 줄
    줄바꿈은 줄 단위로 유지한다(빈 줄도 유지).
    """
    lines: list[str] = []

    for line in (cell.get("cellTextLines") or []):
        raw_line_text = line.get("text")
        if raw_line_text is not None:
            lines.append(str(raw_line_text))
            continue

        words = []
        for w in (line.get("cellWords") or []):
            t = w.get("inferText")
            if t is not None and str(t).strip() != "":
                words.append(str(t).strip())

        if words:
            lines.append(" ".join(words))
        else:
            lines.append("")

    if not lines:
        root_words = []
        for w in (cell.get("cellWords") or []):
            t = w.get("inferText")
            if t is not None and str(t).strip() != "":
                root_words.append(str(t).strip())
        if root_words:
            lines.append(" ".join(root_words))
        else:
            lines.append("")

    return "\n".join(lines)

def _summarize_one_image(image: Dict[str, Any]) -> Dict[str, Any]:
    out = {"name": image.get("name"),
           "pageIndex": image.get("convertedImageInfo", {}).get("pageIndex"),
           "tables": [], "freeText": []}

    for t in image.get("tables", []) or []:
        cells = t.get("cells", []) or []
        tab = {"cells": []}
        for c in cells:
            raw_words = []
            for w in c.get("cellWords", []) or []:
                tw = (w.get("inferText") or "").strip()
                if tw:
                    raw_words.append(tw)
            for line in c.get("cellTextLines", []) or []:
                for w in line.get("cellWords", []) or []:
                    tw = (w.get("inferText") or "").strip()
                    if tw:
                        raw_words.append(tw)

            tab["cells"].append({
                "rowIndex": c.get("rowIndex"),
                "columnIndex": c.get("columnIndex"),
                "rowSpan": c.get("rowSpan", 1),
                "columnSpan": c.get("columnSpan", 1),
                "text": _cell_text(c),
                "rawWords": " ".join(raw_words) if raw_words else ""
            })
        out["tables"].append(tab)

    # freeText 수집
    for maybe in ("fields", "lines"):
        if maybe in image and isinstance(image[maybe], list):
            for node in image[maybe]:
                t = (node.get("inferText") or "").strip()
                if t:
                    out["freeText"].append(t)
    return out

def _summarize_ocr_result(ocr_item: Dict[str, Any]) -> Dict[str, Any]:
    images = (ocr_item.get("ocr_result") or {}).get("images") or []
    pages = [_summarize_one_image(img) for img in images]
    return {
        "original_image": ocr_item.get("original_image"),
        "binary_image": ocr_item.get("binary_image"),
        "pages": pages
    }

def _merge_continuations_in_struct(parsed: dict) -> dict:
    """
    GPT가 만든 structured JSON에서 표제부(partOfTitle)처럼 연속 행을
    앞 행으로 병합해 1행으로 만드는 안전한 후처리.
    - 같은 descriptionNo 이거나
    - descriptionNo/acceptance/location 이 모두 빈 행이면
      -> 직전 행의 연속으로 보고 특정 키들을 줄바꿈으로 이어붙임.
    """
    def merge_rows(rows: list[dict]) -> list[dict]:
        if not isinstance(rows, list):
            return rows
        merged: list[dict] = []
        for row in rows:
            # 연속(붙여쓰기) 조건
            same_id = bool(merged) and (row.get("descriptionNo") or "").strip() != "" and \
                      row.get("descriptionNo") == merged[-1].get("descriptionNo")
            leading_blank = bool(merged) and \
                (row.get("descriptionNo","") == "" and row.get("acceptance","") == "" and row.get("location","") == "")

            if same_id or leading_blank:
                prev = merged[-1]
                for k in ["buildingDetails", "causeOfRegistrationAndOtherInformation"]:
                    old = (prev.get(k) or "").rstrip()
                    new = (row.get(k) or "").strip()
                    if new:
                        prev[k] = old + ("\n" if old else "") + new
            
                for k in ["descriptionNo", "acceptance", "location"]:
                    if not (prev.get(k) or "").strip() and (row.get(k) or "").strip():
                        prev[k] = row[k]
            else:
                merged.append(dict(row))
        return merged

    # 표제부
    po = parsed.get("partOfTitle")
    if isinstance(po, dict) and isinstance(po.get("rows"), list):
        po["rows"] = merge_rows(po["rows"])
    return parsed

# 프롬프트
def get_prompts_by_doc_type(doc_type: str) -> tuple[str, str]:
    if doc_type == "부동산등기부등본":
        return (
            """
            당신은 등기부 등본 이미지를 JSON 구조로 정리해주는 전문가입니다.
            주어진 이미지들은 등기사항일부증명서의 스캔본입니다. 표 안에 있는 내용을 그대로 분석해서, 최대한 문서 구조를 유지한 JSON 형태로 변환해주세요.
            아래 사항을 반드시 지켜주세요:
            - 셀 안의 텍스트를 사람이 보이는 대로 그대로 사용해주세요.
            - 중복되는 내용이 있더라도 정리하지 말고 그대로 적어주세요.
            - 표제부, 명의인 등은 항목 단위로 나누고, 내부 항목은 딕셔너리처럼 구성해주세요.
            - 항목명이 없는 셀이나 병합된 셀도 보이는 대로 묶어 적어주세요.
            - 참고사항 및 비고도 적어주세요.
            - 날짜, 주소, 이름, 지분 등을 해석하지 말고 그대로 써 주세요.
            - key 값은 예시에 주어진 값 그대로 사용해주세요.
            - value값은 번역하지 말고 그대로 사용해주세요.
            - 여러 rows가 있다면 같은 key의 value값을 병합해주세요.
            """,
            """
            예시:
            {
            "documentType": "등기사항일부증명서(현재 소유현황)",
            "typeOfRegistration": "건물",
            "serialNumber": "...",
            "address": "...",
            "partOfTitle": {
                "header": "【표제부】(건물의 표시)",
                "columns": ["표시번호", "접수", "소재지번, 건물명칭 및 번호", "건물내역", "등기원인 및 기타사항"],
                "rows": [
                {
                    "descriptionNo": "1",
                    "acceptance": "2011년 4월 23일",
                    "location": "...",
                    "buildingDetails": "...",
                    "causeOfRegistrationAndOtherInformation": "..."
                }
                ]
            },
            "owner": {
                "header": "【명의인】",
                "columns": ["등기명의인", "(주민)등록번호", "최종지분", "주소", "순위번호"],
                "rows": [
                {
                    "registeredOwner": "...",
                    "registrationNumber": "...",
                    "finalShare": "...",
                    "ownerAddress": "...",
                    "priorityNumber": "..."
                }
                ]
            },
            "competentRegistryOffice": "...",
            "dateOfIssue": "...",
            "remarks": [
                "[ 참고사항 ]",
                "가. 등기기록에서 유효한 지분을 가진 소유자 혹은 공유자 현황을 표시합니다.",
                "나. 최종지분은 등기명의인이 가진 최종지분이며, 2개 이상의 순위번호에 지분을 가진 경우 그 지분을 합산하였습니다.",
                "다. 순위번호는 등기명의인이 지분을 가진 등기 순위번호입니다.",
                "라. 신청사항과 관련이 없는 소유권(갑구)과 소유권 이외의 권리(을구)사항은 표시되지 않았습니다.",
                "마. 지분이 통분되어 공시된 경우는 전체의 지분을 통분하여 공시한 것입니다.",
                "* 실선으로 그어진 부분은 말소사항을 표시합니다.",
                "* 기록사항 없는 갑구, 을구는 ‘기록사항 없음’으로 표시합니다."
            ]
            }
            """
        )
   

def call_gpt_for_structured_from_ocr(ocr_list: List[Dict[str, Any]], doc_type: str) -> str:

    summarized = [_summarize_ocr_result(item) for item in ocr_list]

    system_prompt, user_example_text = get_prompts_by_doc_type(doc_type)

    user_payload = {
        "instruction": (
            user_example_text
            + "\n\n[중요 추가 규칙]\n"
              "- documentType은 OCR 상단 제목을 그대로 사용하고 예시의 '등기부등본' 등으로 대체 금지\n"
              "- 표의 모든 셀 텍스트( cell.text 가 비면 cell.rawWords )를 절대 누락하지 말고 JSON에 반영할 것.\n"
              "- 표 병합/행·열 의미 유지, 누락 없이 rows[][]에 모두 채움\n"
              "- 표 밖 줄글/참고문구는 remarks[]에 순서대로 모두 포함\n"
              "- 번역/요약/정규화 금지, 원문 그대로\n"
        ),
        "ocr_summary": summarized
    }

    resp = openai.chat.completions.create(
        model="gpt-4o",
        temperature=0,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": [
                {"type": "text", "text": json.dumps(user_payload, ensure_ascii=False)}
            ]},
        ],
    )
    text = (resp.choices[0].message.content or "").strip()


    if text.startswith("```json"):
        text = text[len("```json"):].lstrip()
    if text.endswith("```"):
        text = text[:-3].rstrip()

    # JSON 파싱
    try:
        parsed = json.loads(text)
    except Exception:
        parsed = {"_raw": text, "_note": "JSON 파싱 실패. 원문 그대로 반환."}
        return json.dumps(parsed, ensure_ascii=False, indent=2)


    parsed = _merge_continuations_in_struct(parsed)

    return json.dumps(parsed, ensure_ascii=False, indent=2)