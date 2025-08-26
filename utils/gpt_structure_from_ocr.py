import json
import os
from typing import List, Dict, Any
from dotenv import load_dotenv
import openai

load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY") or os.getenv("GPT-API-KEY")


def _cell_text(cell: Dict[str, Any]) -> str:
    lines: list[str] = []

    for line in cell.get("cellTextLines", []) or []:
        tline = (line.get("text") or "").strip()
        if tline:
            lines.append(tline)
            continue
        words = []
        for w in line.get("cellWords", []) or []:
            t = (w.get("inferText") or "").strip()
            if t:
                words.append(t)
        if words:
            lines.append(" ".join(words))

    # 셀 루트에 바로 cellWords만 있는 경우
    if not lines:
        root_words = []
        for w in cell.get("cellWords", []) or []:
            t = (w.get("inferText") or "").strip()
            if t:
                root_words.append(t)
        if root_words:
            lines.append(" ".join(root_words))

    # 폴백
    if not lines:
        for k in ("text", "inferText"):
            t = (cell.get(k) or "").strip()
            if t:
                lines.append(t)
                break

    return "\n".join(lines).strip()

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
                "columns": ["표시번호", "접수", "소재지번, 건물명칭 및 번호", "건물내역","등기원인 및 기타사항"]
                "descriptionNo": "1",
                "acceptance": "2011년 4월 23일",
                "location": "...",
                "buildingDetails": "...",
                "causeOfRegistrationAndOtherInformation": "..."
            },
            "owner": {
                "header": "【명의인】",
                "columns": ["등기명의인","(주민)등록번호", "최종지분","주소","순위번호"]
                "registeredOwner": "...",
                "registrationNumber": "...",
                "finalShare": "..."
                "ownerAddress": "..."
                "priorityNumber": "..."
            },
            "competentRegistryOffice": "...",
            "dateOfIssue": "...",
            "remarks": [
                "[ 참 고 사 항 ]\n
                가. 등기기록에서 유효한 지분을 가진 소유자 혹은 공유자 현황을 표시합니다.
                나. 최종지분은 등기명의인이 가진 최종지분이며, 2개 이상의 순위번호에 지분을 가진 경우 그 지분을 합산하였습니다.
                다. 순위번호는 등기명의인이 지분을 가진 등기 순위번호입니다.
                라. 신청사항과 관련이 없는 소유권(갑구)과 소유권 이외의 권리(을구)사항은 표시되지 않았습니다.
                마. 지분이 통분되어 공시된 경우는 전체의 지분을 통분하여 공시한 것입니다.
                * 실선으로 그어진 부분은 말소사항을 표시함. * 기록사항 없는 갑구, 을구는 ‘기록사항 없음’으로 표시함."
                ] 
            }
            """
        )
    elif doc_type == "가족관계증명서":
        return (
            """
            당신은 가족관계증명서 이미지를 JSON 구조로 정리해주는 전문가입니다.

            주어진 이미지는 가족관계증명서(일반)이며, 표에 표시된 항목들을 사람이 읽는 그대로 정리해 주세요.

            다음 사항을 반드시 지켜주세요:
            - 문서에 적힌 항목은 순서대로 모두 반영해주세요.
            - 이름, 주민등록번호, 생년월일, 본, 관계 등은 **해석하지 말고** 그대로 추출해주세요.
            - 원본에 있는 한자는 **번역하지 말고** 그대로 "originOfSurname": "金海" 이런 식으로 가져오세요.
            - 본관(originOfSurname) **제발 한문 그대로** 가져오세요.
            - 가족 구성원은 리스트로 구성하고, 각각 `관계`, `성명`, `출생연월일`, `주민등록번호`, `성별`, `본` 항목을 그대로 써 주세요.
            - 문서 상단과 하단의 발급 정보 및 인증 정보도 함께 JSON에 포함해 주세요.
            - 날짜, 번호, 기관명, 책임자 이름은 텍스트 그대로 옮겨 적어 주세요.
            - key 값은 예시에 주어진 값 그대로 사용해주세요.
            - value값은 번역하지 말고 그대로 사용해주세요.
            """,
            """
            예시:
            {
                "documentType": "가족관계증명서(일반)",
                "placeOfFamilyRegistration": "서울특별시 중구 세종대로 100",
                "dateOfIssue": "2025-07-18",
                "timeOfIssue": "14:54",
                "applicant": "김가영",
                "certificateNumber": "9192-2003-5983-1870",
                "columns": ["구분", "성명", "출생연월일", "주민등록번호", "성별", "본"]
                "registrant": {
                    "category": "본인",
                    "fullName": "김가영(金佳榮)",
                    "dateOfBirth": "2000-08-12",
                    "residentRegistrationNumber": "000812-4******",
                    "sex": "여",
                    "originOfSurname": "金海"
                },
                "familyMembers": [
                    {
                        "category": "부",
                        "fullName": "김철수(金哲洙)",
                        "dateOfBirth": "1970-05-10",
                        "residentRegistrationNumber": "700510-1******",
                        "sex": "남",
                        "originOfSurname": "金海"
                    },
                    {
                        "category": "모",
                        "fullName": "이영희(李英姬)",
                        "dateOfBirth": "1972-09-28",
                        "residentRegistrationNumber": "720928-2******",
                        "sex": "여",
                        "originOfSurname": "全州"
                    },
                    {
                        "category": "배우자",
                        "fullName": "박동수(朴東洙)",
                        "dateOfBirth": "1999-03-23",
                        "residentRegistrationNumber": "990323-3******",
                        "sex": "남",
                        "originOfSurname": "密陽"
                    },
                    {
                        "category": "자녀",
                        "fullName": "박지우(朴智雨)",
                        "dateOfBirth": "2022-11-01",
                        "residentRegistrationNumber": "221101-4******",
                        "sex": "여",
                        "originOfSurname": "密陽"
                    },
                    {
                        "category": "자녀",
                        "fullName": "박하준(朴河準)",
                        "dateOfBirth": "2024-02-14",
                        "residentRegistrationNumber": "240214-1******",
                        "sex": "남",
                        "originOfSurname": "密陽"
                    }
                    ],
                    "issuingAuthority": {
                    "organization": "법원행정처 전산정보중앙관리소",
                    "authorizedOfficer": "전산운영책임관 박준우"
                    },
                    "remarks": [
                    "위 가족관계증명서(일반)는 가족관계등록부의 기록사항과 틀림없음을 증명합니다.",
                    "위 증명서는 「가족관계의 등록 등에 관한 법률」 제15조제2항에 따른 등록사항을 전출한 일반증명서입니다.",
                    "전자 가족관계등록시스템(https://efamily.scourt.go.kr)의 증명서 진위확인 메뉴에서 발급일로부터 3개월까지 위변조 여부를 확인할 수 있습니다."
                    ]
                }
            """
        )
    elif doc_type == "재학증명서":
        return (
             """
            당신은 재학증명서 이미지를 JSON 구조로 정리해주는 전문가입니다.

            주어진 이미지는 재학증명서이며, 표에 표시된 항목들을 사람이 읽는 그대로 정리해 주세요.

            다음 사항을 반드시 지켜주세요:
            - 문서에 적힌 항목은 순서대로 모두 반영해주세요.
            - 일치하는 항목이 없는경우 "" 으로 놔두세요.
            - 문서 상단과 하단의 발급 정보 및 인증 정보도 함께 JSON에 포함해 주세요.
            - 날짜, 번호, 기관명, 발급인(총장, 이사, 이름 등)은 텍스트 그대로 옮겨 적어 주세요.
            - key 값은 예시에 주어진 값 그대로 사용해주세요.
            - value값은 번역하지 말고 그대로 사용해주세요.
            """,
            """
            예시:
            {
                "authenticationNo": "",
                "receiver": "",
                "use": "",
                "fullName": "",
                "dateOfBirth": "",
                "major": "",
                "grade": "",
                "dateOfIssue": "",
                "universityName": "",
                "authorizedOfficer": "",
                "content": "(예시)위의 사실을 증명함"
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