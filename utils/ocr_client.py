from dotenv import load_dotenv
import os
import json
import time
import uuid
import mimetypes
import requests
from pathlib import Path
from typing import Iterable, List, Union

load_dotenv()  # .env 로드
 
INVOKE_URL = os.getenv("INVOKE_URL")         
OCR_SECRET = os.getenv("X_OCR_SECRET")

def _guess_format_and_mime(p: Path):
    ext = p.suffix.lower().lstrip(".") or "jpg"
    mime, _ = mimetypes.guess_type(str(p))
    if not mime:
        # 기본값
        mime = "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
        
    fmt = "jpg" if ext == "jpeg" else ext
    return fmt, mime

def call_ocr(
    image_paths: Union[str, Path, Iterable[Union[str, Path]]],
    save_json_path: Union[str, Path],
    timeout_sec: int = 60
) -> dict:

    if not INVOKE_URL or not OCR_SECRET:
        raise RuntimeError("INVOKE_URL 또는 X-OCR-SECRET 환경변수가 비어 있습니다.")

    # image_paths를 리스트로 변경
    if isinstance(image_paths, (str, Path)):
        paths: List[Path] = [Path(image_paths)]
    else:
        paths = [Path(p) for p in image_paths]

    if not paths:
        raise ValueError("image_paths가 비었습니다.")

    # message 구성
    images_meta = []
    for i, p in enumerate(paths):
        fmt, _ = _guess_format_and_mime(p)
        images_meta.append({"format": fmt, "name": f"page-{i+1}"})

    message = {
        "version": "V2",
        "requestId": str(uuid.uuid4()),
        "timestamp": int(time.time() * 1000),
        "images": images_meta,
        "enableTableDetection": True,
    }

    # files 배열 구성
    files = []
    file_objs = []
    try:
        for p in paths:
            fmt, mime = _guess_format_and_mime(p)
            f = open(p, "rb")
            file_objs.append(f)
            files.append(
                ("file", (p.name, f, mime))
            )

        headers = {"X-OCR-SECRET": OCR_SECRET}
        data = {"message": json.dumps(message, ensure_ascii=False)}

        resp = requests.post(
            INVOKE_URL,
            headers=headers,
            data=data,           
            files=files,         
            timeout=timeout_sec
        )

        if resp.status_code >= 400:
            raise requests.HTTPError(
                f"CLOVA OCR {resp.status_code}: {resp.text}",
                response=resp
            )

        result = resp.json()

        save_path = Path(save_json_path)
        save_path.parent.mkdir(parents=True, exist_ok=True)
        save_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

        return result

    finally:
        for f in file_objs:
            try:
                f.close()
            except Exception:
                pass