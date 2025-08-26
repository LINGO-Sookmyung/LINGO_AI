import os
import re
import json
import time
from pathlib import Path
from zipfile import ZipFile
from typing import Any, List, Tuple
import openai
from dotenv import load_dotenv

load_dotenv()

openai.api_key = os.getenv("OPENAI_API_KEY") or os.getenv("GPT-API-KEY")


OPENAI_MODEL = os.getenv("TRANSLATE_MODEL", "gpt-4o-mini")
MAX_CHARS = int(os.getenv("TRANSLATE_BATCH_MAX_CHARS", "4000"))

SYSTEM_TMPL = (
    "당신은 공증문서 번역가입니다.\n"
    "주어진 JSON 조각의 value 문자열만 {lang}로 번역하세요.\n"
    "- key는 절대 바꾸지 마세요\n"
    "- 가족관계증명서에 있는 한자는 절대 번역하지 마세요\n"
    "- 원본에 있는 한자는 **한글로 번역이나 치환하지 말고** 그대로 가져오세요.\n"
    "- 본관(originOfSurname) **제발 한문 그대로** 가져오세요.\n"
    "- JSON 스키마 유지\n"
    "- 숫자/날짜/식별자는 번역하지 말 것\n"
    
    "- 결과는 오직 JSON만 반환: {{\"values\": [ ... ]}}\n"
)

# 숫자/날짜/식별자 스킵 패턴
_NUMERIC_LIKE = re.compile(r"^\s*[\d\-\./:,\s]+$")        
_ID_LIKE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-_/]+$")


def _is_translatable_string(s: str) -> bool:
    if not s or not s.strip():
        return False
    if _NUMERIC_LIKE.match(s):
        return False
    if _ID_LIKE.match(s) and len(s) >= 6:
        return False
    return True


def _collect_strings(n: Any, path: Tuple = ()) -> List[Tuple[Tuple, str]]:
    out: List[Tuple[Tuple, str]] = []
    if isinstance(n, dict):
        for k, v in n.items():
            out.extend(_collect_strings(v, path + (k,)))
    elif isinstance(n, list):
        for i, v in enumerate(n):
            out.extend(_collect_strings(v, path + (i,)))
    elif isinstance(n, str):
        if _is_translatable_string(n):
            out.append((path, n))
    return out


def _inject_strings(root: Any, pairs: List[Tuple[Tuple, str]]) -> None:
    for path, val in pairs:
        cur = root
        for p in path[:-1]:
            cur = cur[p]
        cur[path[-1]] = val


def _make_batches(items: List[Tuple[Tuple, str]], max_chars: int = None) -> List[List[Tuple[Tuple, str]]]:
    if max_chars is None:
        max_chars = MAX_CHARS
    batches, cur, cur_len = [], [], 0
    for it in items:
        add = len(it[1]) + 10
        if cur and cur_len + add > max_chars:
            batches.append(cur)
            cur, cur_len = [], 0
        cur.append(it)
        cur_len += add
    if cur:
        batches.append(cur)
    return batches


def _call_openai_with_retry(messages, max_retries=5, initial_wait=2):
    wait = initial_wait
    last_err = None
    for i in range(max_retries):
        try:
            return openai.chat.completions.create(
                model=OPENAI_MODEL,
                messages=messages,
                temperature=0,
            )
        except openai.RateLimitError as e:
            last_err = e
            msg = str(e).lower()
            if "insufficient_quota" in msg:
                raise RuntimeError("OpenAI 쿼터 부족(insufficient_quota)") from e
            if i == max_retries - 1:
                raise
            time.sleep(wait)
            wait = min(wait * 2, 20)
        except openai.APIError as e:
            last_err = e
            if i == max_retries - 1:
                raise
            time.sleep(wait)
            wait = min(wait * 2, 20)
    if last_err:
        raise last_err


def _translate_batch(values: List[str], lang: str) -> List[str]:
    system_prompt = SYSTEM_TMPL.format(lang=lang)
    user_payload = json.dumps({"values": values}, ensure_ascii=False)

    resp = _call_openai_with_retry([
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": [{"type": "text", "text": user_payload}]}
    ])
    text = (resp.choices[0].message.content or "").strip()

    # JSON 파싱
    try:
        data = json.loads(text)
        if isinstance(data, dict) and isinstance(data.get("values"), list):
            return [str(x) for x in data["values"]]
        raise ValueError("unexpected shape")
    except Exception:
        # 폴백
        out = []
        for v in values:
            one = json.dumps({"values": [v]}, ensure_ascii=False)
            r = _call_openai_with_retry([
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": [{"type": "text", "text": one}]}
            ])
            t = (r.choices[0].message.content or "").strip()
            try:
                d = json.loads(t)
                out.append(str(d["values"][0]))
            except Exception:
                out.append(v) 
        return out

#JSON 문자열을 로드 → value들만 번역 → JSON 문자열로 반환
def _translate_json_text(json_text: str, lang: str) -> str:
    root = json.loads(json_text)
    pairs = _collect_strings(root)
    if not pairs:
        return json.dumps(root, ensure_ascii=False, indent=2)

    batches = _make_batches(pairs, max_chars=MAX_CHARS)
    translated_pairs: List[Tuple[Tuple, str]] = []

    for batch in batches:
        vals = [v for _, v in batch]
        tr_vals = _translate_batch(vals, lang)
        translated_pairs.extend([(path, tv) for (path, _), tv in zip(batch, tr_vals)])

    _inject_strings(root, translated_pairs)
    return json.dumps(root, ensure_ascii=False, indent=2)


def call_gpt_for_translate_json(json_path: str, lang: str) -> str:
    print(f"[DEBUG] 요청받은 JSON 경로: {json_path}")
    p = Path(json_path) 
    json_text = p.read_text(encoding="utf-8-sig")

    return _translate_json_text(json_text, lang)