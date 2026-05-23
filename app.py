"""
학생부 기반 모의 면접 질문지 생성기
- 동국대 학생부위주전형 가이드북 출제원리(WHAT-WHY-HOW-SO WHAT) 반영
- 영역 연계형 심화 질문 + 꼬리질문 2단계 구조
- 화면 표시 / PDF / DOCX 동시 다운로드
"""

import streamlit as st
import json
import io
import re
from datetime import datetime
from pathlib import Path

from google import genai
from google.genai import types

import pypdf
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, KeepTogether
)

# ═══════════════════════════════════════════════════════════
# 페이지 설정
# ═══════════════════════════════════════════════════════════
st.set_page_config(
    page_title="학생부 기반 모의 면접 질문지 생성기",
    page_icon="🎤",
    layout="wide",
)

# ═══════════════════════════════════════════════════════════
# 시스템 프롬프트 불러오기
# ═══════════════════════════════════════════════════════════
PROMPT_FILE = Path("prompts/system_rules.md")
if PROMPT_FILE.exists():
    SYSTEM_PROMPT = PROMPT_FILE.read_text(encoding="utf-8")
else:
    SYSTEM_PROMPT = "당신은 면접 출제 전문가입니다. 학생부 내용을 바탕으로 질문을 생성하세요."

# ═══════════════════════════════════════════════════════════
# 한글 폰트 등록 (PDF용)
# ═══════════════════════════════════════════════════════════
FONT_PATH = Path("fonts/NanumGothic.ttf") 
FONT_NAME = "NanumGothic"
FONT_LOADED = False
if FONT_PATH.exists():
    try:
        pdfmetrics.registerFont(TTFont(FONT_NAME, str(FONT_PATH)))
        FONT_LOADED = True
    except Exception:
        FONT_LOADED = False

# ═══════════════════════════════════════════════════════════
# 유틸: PDF에서 텍스트 추출
# ═══════════════════════════════════════════════════════════
def extract_text_from_pdf(file) -> str:
    """업로드된 PDF에서 텍스트 추출"""
    try:
        reader = pypdf.PdfReader(file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as e:
        st.error(f"PDF 읽기 실패: {e}")
        return ""

# ═══════════════════════════════════════════════════════════
# 유틸: Gemini 호출
# ═══════════════════════════════════════════════════════════
def call_gemini(api_key: str, model: str, user_input: str) -> dict:
    """Gemini API를 호출해 JSON 결과를 반환"""
    client = genai.Client(api_key=api_key)

    response = client.models.generate_content(
        model=model,
        contents=[
            {"role": "user", "parts": [{"text": user_input}]},
        ],
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT,
            temperature=0.4,
            response_mime_type="application/json",
        ),
    )

    raw = response.text.strip()
    raw = re.sub(r"^
http://googleusercontent.com/immersive_entry_chip/0

이제 저장하신 후 스트림릿 페이지를 새로고침 해보시면 에러가 완전히 사라진 화면을 보실 수 있을 것입니다!
