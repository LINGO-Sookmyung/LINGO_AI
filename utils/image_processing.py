def binarize_image(image_path: str, save_dir: str) -> str:
    import cv2, os

    # 이미지 파일이 존재하는지 확인
    if not os.path.exists(image_path):
        raise FileNotFoundError("이미지 경로가 존재하지 않습니다.")

    # 이미지를 OpenCV로 읽어옴 (기본은 컬러)
    img = cv2.imread(image_path)
    if img is None:
        raise ValueError("이미지를 불러올 수 없습니다.")  # 경로는 있지만 형식이 잘못됐을 수도 있음

    # 이미지 → 흑백(Grayscale)으로 변환
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 흑백 이미지 → 이진화 (픽셀 값을 0 또는 255로 분리)
    # threshold 값 200 이상이면 255(흰색), 그 이하면 0(검정)
    _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY)

    # 결과 저장 폴더가 없으면 생성
    os.makedirs(save_dir, exist_ok=True)

    # 원본 이미지 이름에서 확장자 제거 → 새 파일명 구성
    base_name = os.path.basename(image_path).split('.')[0]
    binary_path = os.path.join(save_dir, f"{base_name}_binary.png")

    # 이진화된 이미지를 파일로 저장
    success = cv2.imwrite(binary_path, binary)

    if not success:
        raise RuntimeError(f"이미지 저장 실패: {binary_path}")

    # 저장된 이진화 이미지 경로 반환
    return binary_path
