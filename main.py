from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List
import os
import json
import zipfile
import shutil
import uuid
from utils.image_processing import binarize_image
from utils.ocr_client import call_ocr
from utils.gpt_client import call_gpt_for_structured_json

app = FastAPI()

class MultiImagePathRequest(BaseModel):
    image_paths: List[str]  # 여러 이미지 경로 리스트
    doc_type: str

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
            result_path = os.path.join(output_dir, f"{base_name}_result.json")

            entry = {
                "original_image": path,
                "binary_image": binary_path,
            }

            # doc_type이 부동산등기부등본인 경우에만 OCR 수행
            if request.doc_type == "부동산등기부등본":
                result_path = os.path.join(output_dir, f"{base_name}_result.json")
                result = call_ocr(binary_path, result_path)
                entry["ocr_json_file"] = result_path
                entry["ocr_result"] = result

            results.append(entry)
            image_paths_for_gpt.append(binary_path)

        # GPT 결과 호출
        gpt_json_result = call_gpt_for_structured_json(image_paths_for_gpt, request.doc_type)
        gpt_result_path = os.path.join(output_dir, "gpt_structured_result.json")
        with open(gpt_result_path, 'w', encoding='utf-8') as f:
            f.write(gpt_json_result)

        # 결과 저장 및 반환 방식 분기
        if request.doc_type == "부동산등기부등본":
            # OCR 결과도 합쳐서 저장
            ocr_merged_path = os.path.join(output_dir, "ocr_results.json")
            with open(ocr_merged_path, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False)

            # Zip으로 묶음
            zip_path = os.path.join(output_dir, "ocr_and_gpt_results.zip")
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                zipf.write(ocr_merged_path, arcname="ocr_results.json")
                zipf.write(gpt_result_path, arcname="gpt_structured_result.json")
            
            # 요청 끝나고 디렉토리 삭제 예약
            background_tasks.add_task(delete_directory, output_dir)

            return FileResponse(
                path=zip_path,
                media_type="application/zip",
                filename="ocr_and_gpt_results.zip"
            )
        else:
            # 요청 끝나고 디렉토리 삭제 예약
            background_tasks.add_task(delete_directory, output_dir)

            # OCR 없이 GPT 결과만 리턴
            return FileResponse(
                path=gpt_result_path,
                media_type="application/json",
                filename="gpt_structured_result.json"
            )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))