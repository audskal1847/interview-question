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

    # 파일이 없을 때 폴백

    SYSTEM_PROMPT = "당신은 면접 출제 전문가입니다. 학생부 내용을 바탕으로 질문을 생성하세요."



# ═══════════════════════════════════════════════════════════

# 한글 폰트 등록 (PDF용)

# ═══════════════════════════════════════════════════════════

FONT_PATH = Path("C:/Windows/Fonts/malgun.ttf") 

FONT_NAME = "MalgunGothic" # 폰트 이름도 변경

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

    # 혹시 모를 마크다운 코드블록 제거

    raw = re.sub(r"^```(?:json)?\s*", "", raw)

    raw = re.sub(r"\s*```$", "", raw)

    return json.loads(raw)



# ═══════════════════════════════════════════════════════════

# 유틸: DOCX 생성

# ═══════════════════════════════════════════════════════════

def build_docx(sheet: dict, meta: dict) -> bytes:

    """질문지를 DOCX 파일 바이트로 반환"""

    doc = Document()



    # 기본 스타일 (한글 글꼴)

    style = doc.styles["Normal"]

    style.font.name = "맑은 고딕"

    style.font.size = Pt(11)



    # 제목

    title = doc.add_heading("모의 면접 질문지", level=0)

    title.alignment = WD_ALIGN_PARAGRAPH.CENTER



    # 메타 정보

    p = doc.add_paragraph()

    p.add_run(

        f"희망 전공: {meta['major']}    |    "

        f"생성일: {meta['date']}    |    "

        f"메인 질문 {len(sheet['questions'])}개"

    ).italic = True

    doc.add_paragraph("─" * 50)



    # 질문 출력

    for q in sheet["questions"]:

        # 메인 질문

        h = doc.add_paragraph()

        run = h.add_run(f"Q{q['no']}. {q['main']}")

        run.bold = True

        run.font.size = Pt(12)



        # 메타 라인

        areas = " / ".join(q.get("areas", []))

        intent = " · ".join(q.get("intent", []))

        meta_line = doc.add_paragraph()

        meta_run = meta_line.add_run(

            f"   활동영역: {areas}  |  학년: {q.get('grade','-')}  |  의도: {intent}"

        )

        meta_run.font.size = Pt(9)

        meta_run.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)



        # 꼬리질문

        for i, fu in enumerate(q.get("followups", []), 1):

            fu_p = doc.add_paragraph(f"   └ 꼬리질문 {i}. {fu}")

            fu_p.paragraph_format.left_indent = Cm(0.5)



        # 근거

        if q.get("evidence"):

            ev_p = doc.add_paragraph()

            ev_run = ev_p.add_run(f"   📎 학생부 근거: {q['evidence']}")

            ev_run.italic = True

            ev_run.font.size = Pt(9)

            ev_run.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)



        doc.add_paragraph()  # 간격



    # 바이트로 변환

    buf = io.BytesIO()

    doc.save(buf)

    buf.seek(0)

    return buf.getvalue()



# ═══════════════════════════════════════════════════════════

# 유틸: PDF 생성

# ═══════════════════════════════════════════════════════════

def build_pdf(sheet: dict, meta: dict) -> bytes:

    """질문지를 PDF 파일 바이트로 반환"""

    if not FONT_LOADED:

        # 폰트가 없으면 안내 메시지 PDF 반환

        return _build_fallback_pdf()



    buf = io.BytesIO()

    doc = SimpleDocTemplate(

        buf, pagesize=A4,

        leftMargin=2*cm, rightMargin=2*cm,

        topMargin=2*cm, bottomMargin=2*cm,

    )



    styles = getSampleStyleSheet()

    H1 = ParagraphStyle("H1", parent=styles["Heading1"],

                        fontName=FONT_NAME, fontSize=18,

                        alignment=1, spaceAfter=14)

    META = ParagraphStyle("META", parent=styles["Normal"],

                          fontName=FONT_NAME, fontSize=9,

                          textColor=HexColor("#6b7280"),

                          alignment=1, spaceAfter=12)

    QMAIN = ParagraphStyle("QMAIN", parent=styles["Normal"],

                           fontName=FONT_NAME, fontSize=11,

                           leading=16, spaceBefore=8, spaceAfter=4)

    QTAG = ParagraphStyle("QTAG", parent=styles["Normal"],

                          fontName=FONT_NAME, fontSize=8.5,

                          textColor=HexColor("#6b7280"),

                          leading=12, spaceAfter=4)

    QFU = ParagraphStyle("QFU", parent=styles["Normal"],

                         fontName=FONT_NAME, fontSize=10,

                         leading=14, leftIndent=15, spaceAfter=2)

    QEV = ParagraphStyle("QEV", parent=styles["Normal"],

                         fontName=FONT_NAME, fontSize=8.5,

                         textColor=HexColor("#2563eb"),

                         leading=12, leftIndent=15, spaceAfter=12)



    story = []

    story.append(Paragraph("모의 면접 질문지", H1))

    story.append(Paragraph(

        f"희망 전공: {meta['major']}  |  생성일: {meta['date']}  |  "

        f"총 {len(sheet['questions'])}문항", META))



    for q in sheet["questions"]:

        block = []

        block.append(Paragraph(

            f"<b>Q{q['no']}.</b> {q['main']}", QMAIN))

        areas = " / ".join(q.get("areas", []))

        intent = " · ".join(q.get("intent", []))

        block.append(Paragraph(

            f"활동영역: {areas}  |  학년: {q.get('grade','-')}  |  의도: {intent}",

            QTAG))

        for i, fu in enumerate(q.get("followups", []), 1):

            block.append(Paragraph(f"└ 꼬리질문 {i}. {fu}", QFU))

        if q.get("evidence"):

            block.append(Paragraph(

                f"※ 학생부 근거: {q['evidence']}", QEV))

        # 한 질문은 가능하면 한 페이지에 묶어 출력

        story.append(KeepTogether(block))



    doc.build(story)

    buf.seek(0)

    return buf.getvalue()



def _build_fallback_pdf() -> bytes:

    """폰트가 없을 때 안내용 PDF"""

    buf = io.BytesIO()

    doc = SimpleDocTemplate(buf, pagesize=A4)

    styles = getSampleStyleSheet()

    doc.build([Paragraph(

        "Korean font (NanumGothic.ttf) not found in /fonts. "

        "PDF cannot be generated. Please use the DOCX download instead.",

        styles["Normal"])])

    buf.seek(0)

    return buf.getvalue()



# ═══════════════════════════════════════════════════════════

# 사이드바: 설정

# ═══════════════════════════════════════════════════════════

with st.sidebar:

    st.markdown("### 🔑 기본 설정")

    api_key = st.text_input("Google AI API 키", type="password",

                            help="https://aistudio.google.com/apikey 에서 무료 발급")

    st.markdown("[👉 무료 API 키 발급받기](https://aistudio.google.com/apikey)")



    st.divider()

    st.markdown("### ⚙️ 모델 설정")

    model_name = st.selectbox(

        "Gemini 모델",

        ["gemini-2.5-flash", "gemini-2.5-pro"],

        index=0,

        help="Pro가 품질은 더 좋지만 속도가 약간 느립니다.",

    )



    st.divider()

    st.markdown("### 📊 질문지 옵션")

    n_main = st.slider("메인 질문 수", 6, 12, 8)

    n_tail = st.slider("질문당 꼬리질문 수", 0, 3, 2)



    st.divider()

    with st.expander("💡 도움말"):

        st.markdown("""

        - **학생부 입력**: PDF 업로드 또는 직접 붙여넣기

        - **NEIS 원본 PDF**는 보안 처리되어 읽히지 않을 수 있습니다.

          그럴 땐 텍스트를 직접 복사해 붙여넣으세요.

        - **개인정보는 반드시 마스킹** 후 입력하세요.

        """)



# ═══════════════════════════════════════════════════════════

# 메인 영역

# ═══════════════════════════════════════════════════════════

st.title("🎤 학생부 기반 모의 면접 질문지 생성기")

st.caption(

    "동국대 학생부위주전형 가이드북 출제원리(WHAT-WHY-HOW-SO WHAT, 영역 연계, "

    "활동 내부 심화)를 반영합니다."

)



col_left, col_right = st.columns([1.2, 1])



# ─── 좌측: 학생부 입력

with col_left:

    st.subheader("1. 학생부 데이터 입력")



    uploaded = st.file_uploader(

        "📁 학생부 PDF 업로드 (선택)",

        type=["pdf"],

        help="NEIS 보안 PDF는 읽히지 않을 수 있습니다. 그럴 땐 아래에 직접 붙여넣으세요.",

    )



    extracted = ""

    if uploaded is not None:

        with st.spinner("PDF에서 텍스트 추출 중..."):

            extracted = extract_text_from_pdf(uploaded)

        if extracted.strip():

            st.success(f"PDF에서 {len(extracted):,}자를 추출했습니다.")

        else:

            st.warning("PDF에서 텍스트를 추출하지 못했습니다. 아래에 직접 붙여넣어 주세요.")



    record_text = st.text_area(

        "📝 학생부 내용 (PDF가 안 읽히면 여기에 직접 붙여넣기)",

        value=extracted,

        height=400,

        placeholder=(

            "[1학년]\n"

            "[자율활동] ...\n"

            "[동아리활동] ...\n"

            "[진로활동] ...\n"

            "[세부능력 및 특기사항] ...\n"

            "...\n\n"

            "[2학년] ...\n[3학년] ...\n"

            "[독서활동] ...\n[행동특성 및 종합의견] ..."

        ),

    )



# ─── 우측: 강조 포인트

# ─── 우측: 분석 옵션

with col_right:

    st.subheader("2. 분석 옵션")

    major = st.text_input(

        "🎓 희망 전공 / 학과",

        placeholder="예: 전자공학과",

    )

    

    st.markdown("---")

    st.markdown("**🏫 목표 대학 면접 스타일 반영 (선택)**")

    univ_uploaded = st.file_uploader(

        "대학별 면접 가이드북/기출문제 업로드",

        type=["pdf"],

        help="지원 대학의 가이드북을 업로드하면 해당 대학의 출제 경향과 평가 요소를 반영합니다.",

    )



    st.divider()

    run = st.button(

        "🚀 면접 질문지 생성 시작",

        type="primary",

        use_container_width=True,

    )

  

# ═══════════════════════════════════════════════════════════

# 실행부 (하나로 통합)

# ═══════════════════════════════════════════════════════════

if run:

    if not api_key:

        st.error("Google AI API 키를 입력해주세요.")

        st.stop()

    if not record_text.strip():

        st.error("학생부 내용을 입력해주세요.")

        st.stop()

    if not major.strip():

        st.error("희망 전공을 입력해주세요.")

        st.stop()



    # 1. 대학별 가이드북 텍스트 추출

    univ_guide_text = ""

    if univ_uploaded is not None:

        with st.spinner("대학별 면접 가이드북 분석 중..."):

            univ_guide_text = extract_text_from_pdf(univ_uploaded)



    # 2. 프롬프트 구성

    user_input = (

        f"희망 전공: {major}\n"

        f"메인 질문 수: {n_main}\n"

        f"꼬리질문 수: {n_tail}\n\n"

    )



    if univ_guide_text.strip():

        user_input += (

            f"[목표 대학 특화 면접 가이드]\n"

            f"제공된 대학 가이드북의 평가 기준과 스타일을 최우선으로 반영하여 질문을 생성할 것.\n"

            f"{univ_guide_text}\n\n"

        )



    user_input += f"[학생부]\n{record_text}"



    # 3. Gemini 호출

    with st.spinner(f"{model_name} 모델이 면접 질문지를 생성 중입니다..."):

        try:

            sheet = call_gemini(api_key, model_name, user_input)

        except Exception as e:

            st.error(f"생성 실패: {e}")

            st.stop()



    questions = sheet.get("questions", [])

    if not questions:

        st.error("질문이 생성되지 않았습니다.")

        st.stop()



    st.success(f"✅ 메인 질문 {len(questions)}개를 생성했습니다.")



    # 4. 화면 출력 및 다운로드 버튼 (이하 기존 코드 그대로 유지)

    # ... (생성된 질문 출력 및 다운로드 버튼 로직)

        st.stop()



    st.success(f"✅ 메인 질문 {len(questions)}개를 생성했습니다.")



    # ────────────────────────────────────────────────

    # 화면 출력 (A)

    # ────────────────────────────────────────────────

    st.markdown("---")

    st.markdown("## 📋 생성된 면접 질문지")



    for q in questions:

        with st.container(border=True):

            areas = " / ".join(q.get("areas", []))

            intent = " · ".join(q.get("intent", []))

            grade = q.get("grade", "-")



            st.markdown(f"### Q{q['no']}. {q['main']}")

            st.caption(

                f"활동영역: **{areas}**　|　학년: **{grade}**　|　"

                f"의도: **{intent}**　|　연계질문: "

                f"{'🔗 YES' if q.get('linked') else '–'}"

            )



            for i, fu in enumerate(q.get("followups", []), 1):

                st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;**꼬리질문 {i}.** {fu}")



            if q.get("evidence"):

                with st.expander("📎 학생부 근거 보기"):

                    st.info(q["evidence"])



    # ────────────────────────────────────────────────

    # 다운로드 버튼 (B, C)

    # ────────────────────────────────────────────────

    st.markdown("---")

    st.markdown("### 📥 다운로드")



    meta = {

        "major": major,

        "date": datetime.now().strftime("%Y-%m-%d %H:%M"),

    }

    today = datetime.now().strftime("%Y%m%d_%H%M")

    base_name = f"면접질문지_{major}_{today}"



    d1, d2, d3 = st.columns(3)

    with d1:

        pdf_bytes = build_pdf(sheet, meta)

        st.download_button(

            "📄 PDF 다운로드",

            data=pdf_bytes,

            file_name=f"{base_name}.pdf",

            mime="application/pdf",

            use_container_width=True,

        )

    with d2:

        docx_bytes = build_docx(sheet, meta)

        st.download_button(

            "📝 Word(docx) 다운로드",

            data=docx_bytes,

            file_name=f"{base_name}.docx",

            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",

            use_container_width=True,

        )

    with d3:

        json_bytes = json.dumps(sheet, ensure_ascii=False, indent=2).encode("utf-8")

        st.download_button(

            "🗂️ 원본 JSON 다운로드",

            data=json_bytes,

            file_name=f"{base_name}.json",

            mime="application/json",

            use_container_width=True,

        )



    if not FONT_LOADED:

        st.warning(

            "⚠️ `fonts/NanumGothic.ttf` 파일이 없어 PDF에 한글이 표시되지 않습니다. "

            "README의 안내에 따라 폰트를 추가해 주세요. (Word·JSON 다운로드는 정상 작동)"

        )
