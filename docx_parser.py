"""
Word 양식(.docx) 파서
[키워드]: 값 형식의 문서를 슬라이드 단위 딕셔너리 리스트로 변환
"""

from __future__ import annotations

from pathlib import Path
from docx import Document


# 지원하는 키워드 목록
KEYWORDS = {"슬라이드 타입", "제목", "부제목", "소제목", "내용"}

# 슬라이드 타입 정규화 매핑
TYPE_ALIASES = {
    "커버": "커버",
    "표지": "커버",
    "cover": "커버",
    "섹션헤더": "섹션헤더",
    "섹션": "섹션헤더",
    "section": "섹션헤더",
    "본문": "본문",
    "내용": "본문",
    "content": "본문",
    "목차": "목차",
    "toc": "목차",
}


def parse_docx(filepath: str | Path) -> list[dict]:
    """
    Word 양식을 파싱하여 슬라이드 딕셔너리 리스트를 반환한다.

    Returns:
        [
            {"type": "커버", "제목": "...", "부제목": "..."},
            {"type": "섹션헤더", "제목": "..."},
            {"type": "본문", "소제목": "...", "내용": ["...", "..."]},
            ...
        ]
    """
    doc = Document(str(filepath))
    slides = []
    current_slide = None
    collecting_content = False  # [내용]: 아래 불릿 수집 중 여부

    for para in doc.paragraphs:
        line = para.text.strip()

        if not line:
            continue

        # [키워드]: 값 형식 검사
        parsed = _parse_keyword_line(line)

        if parsed:
            keyword, value = parsed

            if keyword == "슬라이드 타입":
                # 새 슬라이드 시작
                if current_slide is not None:
                    slides.append(current_slide)
                slide_type = TYPE_ALIASES.get(value.lower(), value)
                current_slide = {"type": slide_type}
                collecting_content = False

            elif keyword == "내용":
                collecting_content = True
                if current_slide is not None:
                    current_slide["내용"] = []
                    # [내용]: 뒤에 바로 텍스트가 있는 경우
                    if value:
                        current_slide["내용"].append(value)

            elif current_slide is not None:
                current_slide[keyword] = value
                collecting_content = False

        elif collecting_content and current_slide is not None:
            # 불릿 포인트 수집 (- 로 시작하거나 일반 텍스트)
            bullet = line.lstrip("-").lstrip("•").lstrip("·").strip()
            if bullet:
                current_slide.setdefault("내용", []).append(bullet)

    # 마지막 슬라이드 추가
    if current_slide is not None:
        slides.append(current_slide)

    return slides


def _parse_keyword_line(line: str) -> tuple[str, str] | None:
    """
    '[키워드]: 값' 또는 '[키워드] : 값' 형식의 라인을 파싱한다.
    대괄호 없이 '키워드: 값' 형식도 지원한다.
    """
    # [키워드]: 값
    if line.startswith("["):
        bracket_end = line.find("]")
        if bracket_end > 0:
            keyword = line[1:bracket_end].strip()
            rest = line[bracket_end + 1:].strip()
            if rest.startswith(":"):
                value = rest[1:].strip()
            else:
                value = rest
            if keyword in KEYWORDS:
                return keyword, value

    # 키워드: 값 (대괄호 없는 형식)
    colon_idx = line.find(":")
    if 0 < colon_idx < 10:
        keyword = line[:colon_idx].strip()
        if keyword in KEYWORDS:
            value = line[colon_idx + 1:].strip()
            return keyword, value

    return None


def validate_slides(slides: list[dict]) -> list[str]:
    """슬라이드 데이터의 유효성을 검사하고 경고 메시지를 반환한다."""
    warnings = []

    if not slides:
        warnings.append("경고: 파싱된 슬라이드가 없습니다. Word 양식 형식을 확인하세요.")
        return warnings

    has_cover = any(s["type"] == "커버" for s in slides)
    if not has_cover:
        warnings.append("경고: 커버 슬라이드가 없습니다. [슬라이드 타입]: 커버 를 추가하세요.")

    for i, slide in enumerate(slides, 1):
        slide_type = slide.get("type", "알수없음")

        if slide_type == "커버" and "제목" not in slide:
            warnings.append(f"경고: 슬라이드 {i} (커버) — [제목]이 없습니다.")

        if slide_type == "섹션헤더" and "제목" not in slide:
            warnings.append(f"경고: 슬라이드 {i} (섹션헤더) — [제목]이 없습니다.")

        if slide_type == "본문":
            if "소제목" not in slide:
                warnings.append(f"경고: 슬라이드 {i} (본문) — [소제목]이 없습니다.")
            if "내용" not in slide or not slide["내용"]:
                warnings.append(f"경고: 슬라이드 {i} (본문) — [내용]이 없습니다.")

    return warnings
