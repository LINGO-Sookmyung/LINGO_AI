def clean_gpt_response(raw_text: str) -> str:
    # 양쪽 공백 제거
    text = raw_text.strip()

    # 첫 줄이 ```json 또는 ```이면 제거
    if text.startswith("```json"):
        text = text[len("```json"):].lstrip()
    elif text.startswith("```"):
        text = text[len("```"):].lstrip()

    # 마지막 줄이 ```이면 제거
    if text.endswith("```"):
        text = text[:-len("```")].rstrip()

    return text
