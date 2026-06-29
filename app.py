"""
Flask 웹 서버
Word 경로 입력 → PPT 생성 → 같은 폴더에 저장
"""

import os
import sys
import webbrowser
import threading
from pathlib import Path

# Windows 콘솔 한글 출력: 환경변수로 설정 (stdout 래핑은 Flask WSGI와 충돌)
if sys.platform == "win32":
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")

from flask import Flask, render_template, request, jsonify

from docx_parser import parse_docx, validate_slides
from generator import generate_pptx

BASE_DIR = Path(__file__).parent
DEFAULT_TEMPLATE = BASE_DIR / "template" / "base_template.pptx"

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    filepath = data.get("filepath", "").strip()

    # 경로에서 따옴표 제거 (붙여넣기 시 포함될 수 있음)
    filepath = filepath.strip('"').strip("'")

    # 입력 검증
    input_path = Path(filepath)

    if not input_path.exists():
        return jsonify({"success": False, "error": f"파일을 찾을 수 없습니다: {filepath}"})

    if input_path.suffix.lower() != ".docx":
        return jsonify({"success": False, "error": ".docx 파일만 지원됩니다."})

    if not DEFAULT_TEMPLATE.exists():
        return jsonify({"success": False, "error": "템플릿 파일(base_template.pptx)이 없습니다."})

    # 출력 경로: 입력 파일과 같은 폴더에 저장
    output_path = input_path.with_suffix(".pptx")

    try:
        # Word 파싱
        slides_data = parse_docx(input_path)

        warnings = validate_slides(slides_data)
        if not slides_data:
            msg = "파싱된 슬라이드가 없습니다. Word 양식을 확인하세요."
            if warnings:
                msg += "\n" + "\n".join(warnings)
            return jsonify({"success": False, "error": msg})

        # PPT 생성
        generate_pptx(slides_data, DEFAULT_TEMPLATE, output_path)

        # 요약
        type_counts = {}
        for s in slides_data:
            t = s.get("type", "기타")
            type_counts[t] = type_counts.get(t, 0) + 1
        # 목차 자동 생성분 포함
        has_sections = any(s.get("type") == "섹션헤더" for s in slides_data)
        total = len(slides_data) + (1 if has_sections else 0)
        summary = ", ".join(f"{t} {c}장" for t, c in type_counts.items())
        if has_sections:
            summary += ", 목차 1장(자동)"
        summary = f"총 {total}장 ({summary})"

        warn_text = ""
        if warnings:
            warn_text = " | " + " / ".join(warnings)

        return jsonify({
            "success": True,
            "output_path": str(output_path),
            "summary": summary + warn_text,
        })

    except Exception as e:
        return jsonify({"success": False, "error": f"생성 중 오류 발생: {str(e)}"})


def open_browser():
    """서버 시작 후 브라우저 자동 열기"""
    webbrowser.open("http://localhost:5500")


if __name__ == "__main__":
    print("=" * 48)
    print("  PPT Auto Generator - Web UI")
    print("  http://localhost:5500")
    print("  (this window must stay open)")
    print("=" * 48)

    # 1.5초 후 브라우저 자동 열기
    threading.Timer(1.5, open_browser).start()

    app.run(host="127.0.0.1", port=5500, debug=False)
