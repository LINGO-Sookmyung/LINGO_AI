from dotenv import load_dotenv
import os

load_dotenv()  # .env 파일 로드

invoke_url = os.getenv("INVOKE_URL")
ocr_secret = os.getenv("X-OCR-SECRET")

def call_ocr(image_path: str, save_json_path: str) -> dict:
    import uuid, time, json, requests, os

    # OCR 요청에 필요한 메타데이터 구성
    # 'images': 분석할 이미지 파일 정보 (format, name)
    # 'requestId': 요청을 식별할 고유 ID
    # 'version': API 버전
    # 'enableTableDetection': 표 구조 인식 활성화 여부
    # 'timestamp': 현재 시간 (밀리초 단위)
    request_json = {
        'images': [{'format': 'png', 'name': 'demo'}],
        'requestId': str(uuid.uuid4()),
        'version': 'V2',
        'enableTableDetection': 'true',
        'timestamp': int(round(time.time() * 1000))
    }

    # 요청 본문(payload) 생성
    # message 필드에 JSON 문자열을 UTF-8로 인코딩해서 포함
    payload = {'message': json.dumps(request_json).encode('UTF-8')}

    # 이미지 파일을 multipart 형식으로 전송하기 위해 준비
    # ('file', 파일객체) 튜플을 리스트 형태로 전달
    files = [('file', open(image_path, 'rb'))]

    # 요청 헤더에 비밀 키 포함 (CLOVA OCR 인증용)
    headers = {'X-OCR-SECRET': ocr_secret}

    # OCR API로 POST 요청 전송
    response = requests.post(invoke_url, headers=headers, data=payload, files=files)

    # 요청 실패 시 예외 발생 (4xx, 5xx)
    response.raise_for_status()

    # 응답 본문(JSON)을 파이썬 dict로 파싱
    result = response.json()

    # 응답 결과를 JSON 파일로 저장
    with open(save_json_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    # 결과 dict 반환 (파일 저장 + 메모리 반환)
    return result
