from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List
import traceback
import os
import json
import shutil
import uuid
from fastapi.staticfiles import StaticFiles
from tempfile import NamedTemporaryFile
from utils.image_processing import binarize_image
from utils.ocr_client import call_ocr
from utils.gpt_client import call_gpt_for_structured_json
from utils.s3_http_downloader import ensure_local
from utils.translate_gpt_client import call_gpt_for_translate_json
from utils.generate_doc.generate_building_registry_docx import generate_building_registry_docx
from utils.generate_doc.generate_enrollment_certificate_docx import generate_enrollment_certificate_docx
from utils.generate_doc.generate_family_relationship_docx import generate_family_relationship_docx
from pydantic import BaseModel
from utils.gpt_structure_from_ocr import call_gpt_for_structured_from_ocr
from fastapi.responses import FileResponse
from utils.is_within_directory import is_within_directory
from typing import Optional, Dict, Any
from fastapi import Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
import logging, sys, os, json, traceback


logging.basicConfig(
    level=logging.DEBUG,  
    format="%(asctime)s %(levelname)s %(name)s %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger("lingoai")

app = FastAPI()

os.makedirs("outputs", exist_ok=True)
app.mount("/outputs", StaticFiles(directory="outputs"), name="outputs")
 


class CreateDocRequest(BaseModel):
    doc_type: str
    lang: str
    json_path: Optional[str] = None 
    ocr_path: Optional[str] = None
    editedContentJson: Optional[Dict[str, Any]] = None #수정json


class MultiImagePathRequest(BaseModel):
    image_paths: List[str]  # 여러 이미지 경로 리스트
    doc_type: str

class JsonPathRequest(BaseModel):
    json_path: str
    lang: str

\
# 디렉터리 삭제
def delete_directory(path: str):
    shutil.rmtree(path)

@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    print("[422 BODY]", await request.body())
    print("[422 ERRORS]", exc.errors())
    return JSONResponse(status_code=422, content={"detail": exc.errors()})
#웹에서 파일 내용 확인용
@app.get("/outputs/{uuid}/{filename}")
def get_output_file(uuid: str, filename: str):
    file_path = f"outputs/{uuid}/{filename}"
    return FileResponse(file_path, media_type="application/json")
    

 #이진화 + OCR + GPT 구조화 (부동산등기부등본/가족관계증명서/재학증명서)
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

            # 이진화
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
            try:
                with open(merged_path, "r", encoding="utf-8") as f:
                    ocr_data = json.load(f)
                if not isinstance(ocr_data, list):
                    raise ValueError("merged_results.json 형식이 리스트가 아닙니다.")

                gpt_json_result = call_gpt_for_structured_from_ocr(ocr_data, request.doc_type)

                gpt_structured_path = os.path.join(output_dir, f"{session_id}_gpt_structured.json")
                with open(gpt_structured_path, "w", encoding="utf-8") as f:
                    f.write(gpt_json_result)
            

                return {"path": gpt_structured_path.replace("\\", "/")}
            
            except Exception as e:
                raise HTTPException(status_code=500, detail=f"등기부 GPT 구조화 실패: {e}")
        else:
            if not image_paths_for_gpt:
                raise HTTPException(status_code=500, detail="gpt 구조화를 위한 이미지가 없습니다.")
            
            gpt_result_path = os.path.join(output_dir, f"{session_id}_gpt_structured_result.json")
            gpt_json_result = call_gpt_for_structured_json(image_paths_for_gpt, request.doc_type)
            with open(gpt_result_path, "w", encoding="utf-8") as f:
                f.write(gpt_json_result)

            return {"path": gpt_result_path.replace("\\", "/")}

    except Exception:
        tb = traceback.format_exc()
        print("ERROR in /binarize-and-ocr-multi:", tb)
        raise HTTPException(status_code=500, detail=tb)
    

# 번역
@app.post("/translate")
def translate(request: JsonPathRequest, background_tasks: BackgroundTasks):
    try:
        base_name = os.path.basename(request.json_path).split('.')[0]
        session_id = str(uuid.uuid4())
        output_dir = os.path.join("outputs", session_id)
        os.makedirs(output_dir, exist_ok=True)

        gpt_json_result = call_gpt_for_translate_json(request.json_path, request.lang)

        # 파일로도 저장
        gpt_result_path = os.path.join(output_dir, f"{base_name}_gpt_translate_result.json")
        with open(gpt_result_path, 'w', encoding='utf-8') as f:
            f.write(gpt_json_result)

        #객체로 
        try:
            obj = json.loads(gpt_json_result)
        except Exception:
            obj = None  

        return {"path": gpt_result_path, "result": obj}
    
    except Exception:
        tb = traceback.format_exc()
        print("ERROR in /translate:", tb)
        raise HTTPException(status_code=500, detail="translate failed")


@app.post("/generate-doc")
def generate_doc(request: CreateDocRequest, background_tasks: BackgroundTasks):
    log.debug(f"/generate-doc payload keys={list(request.model_dump().keys())}")
    print("[DEBUG] doc_type=", request.doc_type)
    print("[DEBUG] json_path=", request.json_path)
    print("[DEBUG] ocr_path=", request.ocr_path)
    print("[DEBUG] exists(json_path)=", os.path.exists(request.json_path or ""))

    # editedContentJson 받으면 임시 파일로 저장해서 json_path로 사용
    temp_json_path = None
    if request.editedContentJson is not None:
        session_id = str(uuid.uuid4())
        output_dir = os.path.join("outputs", session_id)
        os.makedirs(output_dir, exist_ok=True)
        temp_json_path = os.path.join(output_dir, f"{session_id}_edited.json")
        with open(temp_json_path, "w", encoding="utf-8") as f:
            json.dump(request.editedContentJson, f, ensure_ascii=False, indent=2)
        request.json_path = temp_json_path  # 이후 로직은 기존과 동일하게 처리
        print("[DEBUG] wrote editedContentJson to:", temp_json_path)

    # json_path 없으면 에러
    if not request.json_path:
        raise HTTPException(status_code=400, detail="json_path 또는 editedContentJson 중 하나는 필수입니다.")

    if not os.path.exists(request.json_path):
        raise HTTPException(status_code=400, detail=f"json_path가 없습니다: {request.json_path}")

    # 문서 생성
    if request.doc_type == "부동산등기부등본":
        doc = generate_building_registry_docx(request.json_path, request.ocr_path or "", request.lang)
    elif request.doc_type == "가족관계증명서":
        doc = generate_family_relationship_docx(request.json_path, request.lang)
    elif request.doc_type == "재학증명서":
        doc = generate_enrollment_certificate_docx(request.json_path, request.lang)
    else:
        raise HTTPException(status_code=400, detail="지원하지 않는 문서 유형입니다.")

    # 결과 파일 저장 및 반환
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