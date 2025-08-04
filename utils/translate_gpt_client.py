import openai
import base64
from dotenv import load_dotenv
from utils.clean_gpt_response import clean_gpt_response
import os

load_dotenv()  # .env 파일 로드

gpt_api_key = os.getenv("GPT-API-KEY")


def call_gpt_for_translate_json(json_path: str, lang: str) -> str:
    with open(json_path, "r", encoding="utf-8") as f:
        json_content = f.read()
    
    system_prompt = f"""
    당신은 공증문서 번역가입니다.
    이 JSON 파일을 {lang}로 번역해주세요.
    JSON 형태를 유지해주세요.
    JSON 외의 불필요한 텍스트는 제거하세요.
    key값은 유지하고 value값만 번역해주세요.
    """


    messages = [
        {"role": "system", "content": system_prompt},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": json_content}
            ]
        }
    ]

    response = openai.chat.completions.create(
        model="gpt-4o",
        messages=messages
    )

    raw_result = response.choices[0].message.content
    clean_result = clean_gpt_response(raw_result)

    return clean_result