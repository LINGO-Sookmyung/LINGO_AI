from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List
import os
import json
import zipfile
import shutil
import uuid
from docx import Document
from tempfile import NamedTemporaryFile
from utils.image_processing import binarize_image
from utils.ocr_client import call_ocr
from utils.gpt_client import call_gpt_for_structured_json
from utils.translate_gpt_client import call_gpt_for_translate_json
from utils.generate_doc.generate_building_registry_docx import generate_building_registry_docx
from utils.generate_doc.generate_building_registry_docx_simple import generate_building_registry_docx_simple
from utils.generate_doc.generate_enrollment_certificate_docx import generate_enrollment_certificate_docx
from utils.generate_doc.generate_family_relationship_docx import generate_family_relationship_docx

app = FastAPI()

class MultiImagePathRequest(BaseModel):
    image_paths: List[str]  # 여러 이미지 경로 리스트
    doc_type: str

class JsonPathRequest(BaseModel):
    json_path: str
    lang: str

class CreateDocRequest(BaseModel):
    doc_type:str
    json_path:str
    ocr_path:str
    lang:str

def delete_directory(path: str):
    shutil.rmtree(path)

@app.post("/binarize-and-ocr-multi")
def binarize_and_ocr_multi(request: MultiImagePathRequest, background_tasks: BackgroundTasks):
    session_id = str(uuid.uuid4())
    output_dir = os.path.join("outputs", session_id)
    os.makedirs(output_dir, exist_ok=True)

    results = []
    image_paths_for_gpt = []

    try:
        for path in request.image_paths:
            base_name = os.path.basename(path).split('.')[0]
            binary_path = binarize_image(path, output_dir)
            result_path = os.path.join(output_dir, f"{base_name}__{session_id}_result.json")

            entry = {
                "original_image": path,
                "binary_image": binary_path,
            }

            # doc_type이 부동산등기부등본인 경우에만 OCR 수행
            if request.doc_type == "부동산등기부등본":
                result_path = os.path.join(output_dir, f"{base_name}__{session_id}_result.json")
                result = call_ocr(binary_path, result_path)
                entry["ocr_json_file"] = result_path
                entry["ocr_result"] = result

            results.append(entry)
            image_paths_for_gpt.append(binary_path)

        # GPT 결과 호출
        gpt_json_result = call_gpt_for_structured_json(image_paths_for_gpt, request.doc_type)
        gpt_result_path = os.path.join(output_dir, f"{base_name}__{session_id}_gpt_structured_result.json")
        with open(gpt_result_path, 'w', encoding='utf-8') as f:
            f.write(gpt_json_result)

        # 결과 저장 및 반환 방식 분기
        if request.doc_type == "부동산등기부등본":
            # OCR 결과도 합쳐서 저장
            ocr_merged_path = os.path.join(output_dir, f"{base_name}__{session_id}_ocr_results.json")
            with open(ocr_merged_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False)

            # Zip으로 묶음
            zip_path = os.path.join(output_dir, f"{base_name}__{session_id}_ocr_and_gpt_results.zip")
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                zipf.write(ocr_merged_path, arcname=f"{base_name}__{session_id}_ocr_results.json")
                zipf.write(gpt_result_path, arcname=f"{base_name}__{session_id}_gpt_structured_result.json")
            
            # 요청 끝나고 디렉토리 삭제 예약
            background_tasks.add_task(delete_directory, output_dir)

            return FileResponse(
                path=zip_path,
                media_type="application/zip",
                filename=f"{base_name}__{session_id}_ocr_and_gpt_results.zip"
            )
        else:
            # 요청 끝나고 디렉토리 삭제 예약
            background_tasks.add_task(delete_directory, output_dir)

            # OCR 없이 GPT 결과만 리턴
            return FileResponse(
                path=gpt_result_path,
                media_type="application/json",
                filename=f"{base_name}__{session_id}_gpt_structured_result.json"
            )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/translate")
def translate (request: JsonPathRequest):
    base_name = os.path.basename(request.json_path).split('.')[0].split('__')[0]
    session_id = str(uuid.uuid4())
    output_dir = os.path.join("outputs", session_id)
    os.makedirs(output_dir, exist_ok=True)
    
    gpt_json_result = call_gpt_for_translate_json(request.json_path, request.lang)
    gpt_result_path = os.path.join(output_dir, f"{base_name}__{session_id}_gpt_translate_result.json")
    with open(gpt_result_path, 'w', encoding='utf-8') as f:
        f.write(gpt_json_result)

    return FileResponse(
                path=gpt_result_path,
                media_type="application/json",
                filename=f"{base_name}__{session_id}_gpt_translate_result.json"
            )

@app.post("/generate-doc")
def generate_doc (request: CreateDocRequest):
    if request.doc_type == "부동산등기부등본":
        doc = generate_building_registry_docx(request.json_path, request.ocr_path, request.lang)
    elif request.doc_type == "가족관계증명서":
        doc = generate_family_relationship_docx(request.json_path, request.lang)
    elif request.doc_type == "재학증명서":
        doc = generate_enrollment_certificate_docx(request.json_path, request.lang)
    else:
        raise ValueError("지원하지 않는 문서 유형입니다.")

    session_id = str(uuid.uuid4())
    temp_path = os.path.join("translated_outputs", session_id)
    os.makedirs(temp_path, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(request.json_path))[0].split('__')[0]  # 파일명 추출

    # 임시 파일로 저장
    with NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        temp_path = tmp.name
        doc.save(temp_path)

    return FileResponse(
        path=temp_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"{base_name}__{session_id}_translated.docx"
    )
