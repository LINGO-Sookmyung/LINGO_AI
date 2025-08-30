import os
from urllib.parse import urlparse, unquote

import boto3
import requests

__all__ = ["ensure_local", "is_http_url", "is_s3_url"]

s3 = boto3.client("s3")

def is_http_url(p: str) -> bool:
    return p.startswith("http://") or p.startswith("https://")

def is_s3_url(p: str) -> bool:
    return p.startswith("s3://")

def download_http(url: str, out_dir: str) -> str:
    os.makedirs(out_dir, exist_ok=True)
    parsed = urlparse(url)
    base = unquote(os.path.basename(parsed.path)) or "image"
    name, ext = os.path.splitext(base)

    r = requests.get(url, stream=True, timeout=30)
    r.raise_for_status()

    if not ext:
        ct = r.headers.get("Content-Type", "")
        if "jpeg" in ct:
            ext = ".jpg"
        elif "png" in ct:
            ext = ".png"
        elif "webp" in ct:
            ext = ".webp"
        else:
            ext = ".bin"

    local = os.path.join(out_dir, f"{name}{ext}")
    with open(local, "wb") as f:
        for chunk in r.iter_content(8192):
            if chunk:
                f.write(chunk)
    return local

#s3://bucket/key 를 로컬 파일로 다운로드
def download_s3_url(s3_url: str, out_dir: str) -> str:
   
    os.makedirs(out_dir, exist_ok=True)
    parsed = urlparse(s3_url)  
    bucket = parsed.netloc
    key = parsed.path.lstrip("/")
    base = os.path.basename(key) or "image"
    local = os.path.join(out_dir, base)
    s3.download_file(bucket, key, local)
    return local


def ensure_local(path_or_url: str, out_dir: str) -> str:
    if is_http_url(path_or_url):
        return download_http(path_or_url, out_dir)
    if is_s3_url(path_or_url):
        return download_s3_url(path_or_url, out_dir)
    return path_or_url