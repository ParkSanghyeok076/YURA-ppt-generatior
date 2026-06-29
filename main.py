"""
사내강사용 PPT 자동 생성 도구
Word 양식(.docx)을 입력받아 회사 표준 PPT 초안(.pptx)을 생성합니다.
"""

import argparse
import sys
from pathlib import Path

from parser import parse_docx, validate_slides
from generator import generate_pptx

# 기본 경로
BASE_DIR = Path(__file__).parent
DEFAULT_TEMPLATE = BASE_DIR / "template" / "base_template.pptx"


def main():
    args = parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else input_path.with_suffix(".pptx")
    template_path = Path(args.template) if args.template else DEFAULT_TEMPLATE

    # 1. 입력 파일 확인
    if not input_path.exists():
        print(f"오류: 입력 파일을 찾을 수 없습니다 — {input_path}")
        sys.exit(1)

    if not template_path.exists():
        print(f"오류: 템플릿 파일을 찾을 수 없습니다 — {template_path}")
        print("  template/ 폴더에 base_template.pptx를 넣어주세요.")
        sys.exit(1)

    # 2. Word 파싱
    print(f"[1/3] Word 양식 파싱 중... ({input_path.name})")
    slides_data = parse_docx(input_path)

    # 3. 유효성 검사
    warnings = validate_slides(slides_data)
    for w in warnings:
        print(f"  {w}")

    if not slides_data:
        print("오류: 파싱된 슬라이드가 없습니다. 양식을 확인하세요.")
        sys.exit(1)

    # 슬라이드 요약 출력
    type_counts = {}
    for s in slides_data:
        t = s.get("type", "기타")
        type_counts[t] = type_counts.get(t, 0) + 1
    summary = ", ".join(f"{t} {c}장" for t, c in type_counts.items())
    print(f"  → {len(slides_data)}장 파싱 완료 ({summary})")

    # 4. PPT 생성
    print(f"[2/3] PPT 생성 중... (템플릿: {template_path.name})")
    generate_pptx(slides_data, template_path, output_path)

    # 5. 완료
    print(f"[3/3] 완료! → {output_path}")
    print(f"  PowerPoint에서 열어 내용을 확인하고 수정하세요.")


def parse_args():
    p = argparse.ArgumentParser(
        description="사내강사용 PPT 자동 생성 도구",
        epilog="예시: python main.py 강의안.docx -o 강의안_초안.pptx",
    )
    p.add_argument(
        "input",
        help="Word 양식 파일 경로 (.docx)",
    )
    p.add_argument(
        "-o", "--output",
        help="출력 PPT 파일 경로 (기본: 입력파일명.pptx)",
    )
    p.add_argument(
        "-t", "--template",
        help="PPT 템플릿 파일 경로 (기본: template/base_template.pptx)",
    )
    return p.parse_args()


if __name__ == "__main__":
    main()
