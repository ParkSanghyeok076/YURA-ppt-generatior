"""
슬라이드 좌표, 폰트, 스타일 상수 정의
샘플 PPT(사내강사 강의안 sample.pptx)에서 추출한 실제 값 기반
"""

from pptx.util import Emu, Pt


# === 슬라이드 크기 (와이드스크린) ===
SLIDE_WIDTH = Emu(9906000)   # 10.83 inches
SLIDE_HEIGHT = Emu(6858000)  # 7.50 inches


# === 폰트 설정 ===
class Font:
    # 제목/소제목용
    TITLE_NAME = "HY헤드라인M"
    TITLE_SIZE = Pt(24)       # 304800 EMU

    # 소제목용
    SUBTITLE_NAME = "HY헤드라인M"
    SUBTITLE_SIZE = Pt(20)    # 254000 EMU

    # 본문용
    BODY_NAME = "맑은 고딕"
    BODY_SIZE = Pt(18)        # 228600 EMU
    BODY_SIZE_SMALL = Pt(16)  # 203200 EMU
    BODY_SIZE_TINY = Pt(14)   # 177800 EMU

    # 섹션 표시 바
    SECTION_BAR_NAME = "맑은 고딕"
    SECTION_BAR_SIZE = Pt(16)  # 203200 EMU

    # 목차
    TOC_TITLE_SIZE = Pt(36)    # 457200 EMU
    TOC_ITEM_SIZE = Pt(28)     # 355600 EMU


# === 색상 ===
class Color:
    TITLE_RGB = (0x29, 0x29, 0x29)       # 커버/섹션헤더 제목
    BODY_RGB = (0x00, 0x00, 0x00)        # 본문 텍스트
    SECTION_BAR_RGB = (0x00, 0x00, 0x00) # 섹션 표시 바


# === 레이아웃 인덱스 (마스터 0 기준) ===
class Layout:
    COVER = 0           # "1_제목 슬라이드" — 커버 + 섹션헤더용
    BLANK = 5           # "빈 화면" — 본문 + 목차용


# === 슬라이드 타입별 Shape 좌표 (EMU) ===

# 커버 슬라이드 (Slide 1 기준)
class CoverPos:
    TITLE_LEFT = Emu(1066800)
    TITLE_TOP = Emu(2624138)
    TITLE_WIDTH = Emu(7772400)
    TITLE_HEIGHT = Emu(579437)

    SUBTITLE_LEFT = Emu(1066800)
    SUBTITLE_TOP = Emu(3400000)
    SUBTITLE_WIDTH = Emu(7772400)
    SUBTITLE_HEIGHT = Emu(400000)


# 섹션헤더 슬라이드 (Slide 3 기준)
class SectionPos:
    TITLE_LEFT = Emu(1066800)
    TITLE_TOP = Emu(3139282)
    TITLE_WIDTH = Emu(7772400)
    TITLE_HEIGHT = Emu(579437)


# 본문 슬라이드 (Slide 4, 9 등 기준)
class ContentPos:
    # 섹션 표시 바 (Rectangle 3) — 좌측 상단
    SECTION_BAR_LEFT = Emu(50800)
    SECTION_BAR_TOP = Emu(0)
    SECTION_BAR_WIDTH = Emu(5478264)
    SECTION_BAR_HEIGHT = Emu(336550)

    # 소제목 (Text Box 68) — 섹션 바 아래
    SUBTITLE_LEFT = Emu(300038)
    SUBTITLE_TOP = Emu(388938)
    SUBTITLE_WIDTH = Emu(5732462)
    SUBTITLE_HEIGHT = Emu(336550)

    # 콘텐츠 영역 — 소제목 아래 나머지 공간
    BODY_LEFT = Emu(300038)
    BODY_TOP = Emu(900000)
    BODY_WIDTH = Emu(9300000)
    BODY_HEIGHT = Emu(5500000)


# 목차 슬라이드 (Slide 2 기준)
class TocPos:
    # "< 목 차 >" 타이틀
    TITLE_LEFT = Emu(350838)
    TITLE_TOP = Emu(765175)
    TITLE_WIDTH = Emu(9204325)
    TITLE_HEIGHT = Emu(649288)

    # 섹션 제목 리스트
    LIST_LEFT = Emu(344488)
    LIST_TOP = Emu(1628775)
    LIST_WIDTH = Emu(9205912)
    LIST_HEIGHT = Emu(4392613)


# === 로마자 번호 매핑 ===
ROMAN_NUMERALS = [
    "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
    "XI", "XII", "XIII", "XIV", "XV",
]
