from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List
import traceback
import os
import json
import zipfile
import shutil
import uuid
from docx import Document
from fastapi.staticfiles import StaticFiles
from tempfile import NamedTemporaryFile
from utils.image_processing import binarize_image
from utils.ocr_client import call_ocr
from utils.gpt_client import call_gpt_for_structured_json
from utils.s3_http_downloader import ensure_local
from utils.translate_gpt_client import call_gpt_for_translate_json
#from utils.generate_doc.generate_building_registry_docx import generate_building_registry_docx
from utils.generate_doc.generate_building_registry_docx_copy import generate_building_registry_docx
from utils.generate_doc.generate_enrollment_certificate_docx import generate_enrollment_certificate_docx
from utils.generate_doc.generate_family_relationship_docx import generate_family_relationship_docx
from pydantic import BaseModel
from utils.gpt_structure_from_ocr import call_gpt_for_structured_from_ocr
app = FastAPI()
os.makedirs("outputs", exist_ok=True)
app.mount("/outputs", StaticFiles(directory="outputs"), name="outputs")
from fastapi.responses import FileResponse

@app.get("/outputs/{uuid}/{filename}")
def get_output_file(uuid: str, filename: str):
    file_path = f"outputs/{uuid}/{filename}"
    return FileResponse(file_path, media_type="application/json")
class GptStructureRequest(BaseModel):
    ocr_path: str
    doc_type: str

'''@app.post("/gpt-structure")
def gpt_structure(request: GptStructureRequest):
    try:
        # OCR 결과 파일 경로 (이미 존재하는 merged_results.json)
        ocr_path = request.ocr_path  
        with open(ocr_path, "r", encoding="utf-8") as f:
            ocr_data = json.load(f)

        # GPT 호출 (OCR JSON을 템플릿 dict로 변환)
        gpt_json_result = call_gpt_for_structured_json(
            [item["original_image"] for item in ocr_data], 
            request.doc_type
        )

        # 결과 저장
        base_name = os.path.splitext(os.path.basename(ocr_path))[0]
        session_id = str(uuid.uuid4())
        output_dir = os.path.join("outputs", session_id)
        os.makedirs(output_dir, exist_ok=True)

        gpt_result_path = os.path.join(output_dir, f"{base_name}_gpt_structured.json")
        with open(gpt_result_path, "w", encoding="utf-8") as f:
            f.write(gpt_json_result)

        return {"path": gpt_result_path}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
        '''

@app.post("/gpt-structure")
def gpt_structure(request: GptStructureRequest):
    try:
        # 1) OCR 리스트 로드 (여러 페이지)
        with open(request.ocr_path, "r", encoding="utf-8") as f:
            ocr_data = json.load(f)
        if not isinstance(ocr_data, list):
            raise ValueError("ocr_path 파일 형식이 리스트가 아닙니다. merged_results.json을 전달하세요.")

        # 2) OCR 요약을 GPT에 전달하여 구조화(dict) 생성 (누락 방지)
        gpt_json_result = call_gpt_for_structured_from_ocr(ocr_data, request.doc_type)

        # 3) 저장
        base = os.path.splitext(os.path.basename(request.ocr_path))[0]
        sid = str(uuid.uuid4())
        outdir = os.path.join("outputs", sid); os.makedirs(outdir, exist_ok=True)
        out_path = os.path.join(outdir, f"{base}_gpt_structured.json")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(gpt_json_result)

        return {"path": out_path}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    


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
def binarize_and_ocr_multi(request: MultiImagePathRequest):
    print("image_paths:", request.image_paths)
    session_id = str(uuid.uuid4())
    output_dir = os.path.join("outputs", session_id)
    os.makedirs(output_dir, exist_ok=True)

    if not request.image_paths:
        raise HTTPException(status_code=400, detail="image_paths is empty")

    results = []
    image_paths_for_gpt = []  
    try:
        for p in request.image_paths:
            local_input_path = ensure_local(p, output_dir)
            base_name = os.path.splitext(os.path.basename(local_input_path))[0]

            # 1) 이진화
            binary_path = binarize_image(local_input_path, output_dir)

            result_item = {
                "original_image": p,
                "binary_image": binary_path,
            }

         
            if request.doc_type == "부동산등기부등본":
                ocr_json_path = os.path.join(output_dir, f"{base_name}_ocr.json")
                ocr_result = call_ocr(binary_path, ocr_json_path)
                result_item.update({
                    "ocr_json_file": ocr_json_path,
                    "ocr_result": ocr_result,
                })
            else:

                image_paths_for_gpt.append(binary_path)

            results.append(result_item)

        merged_path = os.path.join(output_dir, "merged_results.json")
        with open(merged_path, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

       
        if request.doc_type == "부동산등기부등본":
            # 등기부등본은 기존대로 merged 결과 경로 반환
            return {"path": merged_path.replace("\\", "/")}
        else:
            # 그 외 문서는 GPT 구조화 실행 후 그 파일 경로를 반환
            if not image_paths_for_gpt:
           
                raise HTTPException(status_code=500, detail="No images to run GPT structuring.")
            gpt_result_path = os.path.join(output_dir, f"{session_id}_gpt_structured_result.json")
            gpt_json_result = call_gpt_for_structured_json(image_paths_for_gpt, request.doc_type)
            with open(gpt_result_path, "w", encoding="utf-8") as f:
                f.write(gpt_json_result)
            return {"path": gpt_result_path.replace("\\", "/")}

    except Exception:
        tb = traceback.format_exc()
        print("ERROR in /binarize-and-ocr-multi:", tb)
        raise HTTPException(status_code=500, detail=tb)

@app.post("/translate")
def translate(request: JsonPathRequest):
    try:
        base_name = os.path.basename(request.json_path).split('.')[0]
        session_id = str(uuid.uuid4())
        output_dir = os.path.join("outputs", session_id)
        os.makedirs(output_dir, exist_ok=True)

        gpt_json_result = call_gpt_for_translate_json(request.json_path, request.lang)
        gpt_result_path = os.path.join(output_dir, f"{base_name}_gpt_translate_result.json")
        with open(gpt_result_path, 'w', encoding='utf-8') as f:
            f.write(gpt_json_result)

      
        return {"path": gpt_result_path}

    except Exception as e:
        tb = traceback.format_exc()
        print("ERROR in /translate:", tb)
        raise HTTPException(status_code=500, detail=f"{e}\n\nTRACEBACK:\n{tb}")

@app.post("/generate-doc")
def generate_doc(request: CreateDocRequest):
    if request.doc_type == "부동산등기부등본":
        doc = generate_building_registry_docx(request.json_path, request.ocr_path, request.lang)
    elif request.doc_type == "가족관계증명서":
        doc = generate_family_relationship_docx(request.json_path, request.lang)
    elif request.doc_type == "재학증명서":
        doc = generate_enrollment_certificate_docx(request.json_path, request.lang)
    else:
        raise HTTPException(status_code=400, detail="지원하지 않는 문서 유형입니다.")

    # docx 바이너리는 그대로 FileResponse 유지
    session_id = str(uuid.uuid4())
    os.makedirs("translated_outputs", exist_ok=True)
    with NamedTemporaryFile(delete=False, suffix=".docx", dir="translated_outputs") as tmp:
        temp_path = tmp.name
        doc.save(temp_path)

    base_name = os.path.splitext(os.path.basename(request.json_path))[0]
    return FileResponse(
        path=temp_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"{base_name}_translated.docx"
    )