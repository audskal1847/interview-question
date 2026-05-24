"""
학생부 종합 면접 어시스트 v3.0
- 동국대 학생부위주전형 가이드북 출제원리(WHAT-WHY-HOW-SO WHAT) 반영
- 화면 표시 / DOCX 다운로드 지원 (JSON, PDF 제거 버전)
- 하단 고정 푸터(만든 이) 추가
"""

import streamlit as st
import json
import io
from datetime import datetime
from pathlib import Path

from google import genai
from google.genai import types

import pypdf
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ═══════════════════════════════════════════════════════════
# 페이지 설정
# ═══════════════════════════════════════════════════════════
st.set_page_config(
    page_title="학생부 종합 면접 어시스트 v3.0",
    page_icon="🎤",
    layout="wide",
)

# ═══════════════════════════════════════════════════════════
# 하단 푸터 CSS (중앙 정렬 및 스타일링)
# ═══════════════════════════════════════════════════════════
footer_css = """
<style>
.footer {
    position: fixed;
    left: 0;
    bottom: 0;
    width: 100%;
    background-color: white;
    color: #555555;
    text-align: center;
    padding: 15px 0;
    font-size: 14px;
    border-top: 1px solid #e5e7eb;
    z-index: 1000;
    line-height: 1.6;
}
/* 푸터에 가리지 않도록 메인 컨테이너 하단 여백 추가 */
.main .block-container {
    padding-bottom: 100px; 
}
</style>
<div class="footer">
    <b>🏫 학생부 기반 면접 질문 생성기 시스템 v3.0</b><br>
    만든 이: 신선여자고등학교 김명남<br>
    🗓️ 2026.05
</div>
"""
st.markdown(footer_css, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════
# 시스템 프롬프트 불러오기
# ═══════════════════════════════════════════════════════════
PROMPT_FILE = Path("prompts/system_rules.md")
if PROMPT_FILE.exists():
    SYSTEM_PROMPT = PROMPT_FILE.read_text(encoding="utf-8")
else:
    SYSTEM_PROMPT = "당신은 면접 출제 전문가입니다. 학생부 내용을 바탕으로 동기-과정-결과-배우고느낀점(WHY-HOW-WHAT-LEARN)을 묻는 질문을 생성하세요."

# ═══════════════════════════════════════════════════════════
# 유틸리티 함수
# ═══════════════════════════════════════════════════════════
def extract_text_from_pdf(file) -> str:
    """PDF 텍스트 추출"""
    try:
        reader = pypdf.PdfReader(file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as e:
        st.error(f"PDF 읽기 실패: {e}")
        return ""

def call_gemini(api_key: str, model: str, user_input: str) -> dict:
    """Gemini API 호출 및 JSON 파싱"""
    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(
        model=model,
        contents=[{"role": "user", "parts": [{"text": user_input}]}],
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT,
            temperature=0.4,
            response_mime_type="application/json",
        ),
    )
    # 안전한 JSON 추출
    raw = response.text.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

def build_docx(sheet: dict, meta: dict) -> bytes:
    """질문지를 DOCX 파일로 변환"""
    doc = Document()
    
    # 기본 폰트 설정
    style = doc.styles["Normal"]
    style.font.name = "맑은 고딕"
    style.font.size = Pt(11)

    # 제목
    title = doc.add_heading("모의 면접 질문지", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 메타 정보
    p = doc.add_paragraph()
    p.add_run(f"희망 전공: {meta['major']} | 생성일: {meta['date']} | 메인 질문 {len(sheet['questions'])}개").italic = True
    doc.add_paragraph("─" * 50)

    # 질문 내용 작성
    for q in sheet["questions"]:
        h = doc.add_paragraph()
        run = h.add_run(f"Q{q['no']}. {q['main']}")
        run.bold, run.font.size = True, Pt(12)

        areas = " / ".join(q.get("areas", []))
        intent = " · ".join(q.get("intent", []))
        
        meta_line = doc.add_paragraph()
        meta_run = meta_line.add_run(f"   활동영역: {areas} | 학년: {q.get('grade','-')} | 의도: {intent}")
        meta_run.font.size, meta_run.font.color.rgb = Pt(9), RGBColor(0x6B, 0x72, 0x80)

        for i, fu in enumerate(q.get("followups", []), 1):
            fu_p = doc.add_paragraph(f"   └ 꼬리질문 {i}. {fu}")
            fu_p.paragraph_format.left_indent = Cm(0.5)

        if q.get("evidence"):
            ev_p = doc.add_paragraph()
            ev_run = ev_p.add_run(f"   📎 학생부 근거: {q['evidence']}")
            ev_run.italic, ev_run.font.size, ev_run.font.color.rgb = True, Pt(9), RGBColor(0x25, 0x63, 0xEB)
            
        doc.add_paragraph() # 문항 간 간격

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ═══════════════════════════════════════════════════════════
# 사이드바: 설정
# ═══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🔑 기본 설정")
    api_key = st.text_input("Google AI API 키", type="password", help="https://aistudio.google.com/apikey 에서 무료 발급")
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
        - **NEIS 원본 PDF**는 보안 처리되어 읽히지 않을 수 있습니다. 그럴 땐 텍스트를 직접 복사해 붙여넣으세요.
        - **개인정보는 반드시 마스킹** 후 입력하세요.
        """)

# ═══════════════════════════════════════════════════════════
# 메인 레이아웃
# ═══════════════════════════════════════════════════════════
st.title("🎤 학생부 면접 질문 생성기 시스템 v3.0")
st.caption("WHY-HOW-WHAT-LEARN 원리에 기반한 심층 질문 생성 시스템입니다.")

col_left, col_right = st.columns([1.2, 1])

# ─── 좌측: 학생부 입력
with col_left:
    st.subheader("1. 학생부 데이터 입력")
    uploaded = st.file_uploader("📁 학생부 PDF 업로드 (선택)", type=["pdf"], help="NEIS 보안 PDF는 읽히지 않을 수 있습니다. 그럴 땐 아래에 직접 붙여넣으세요.")
    
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
        placeholder="[1학년]\n[자율활동] ...\n[동아리활동] ...\n\n[2학년] ...\n[3학년] ..."
    )

# ─── 우측: 분석 옵션
with col_right:
    st.subheader("2. 분석 옵션")
    major = st.text_input("🎓 희망 전공 / 학과", placeholder="예: 전자공학과")
    
    st.markdown("---")
    st.markdown("**🏫 목표 대학 면접 스타일 반영 (선택)**")
    univ_uploaded = st.file_uploader(
        "대학별 면접 가이드북/기출문제 업로드", 
        type=["pdf"],
        help="지원 대학의 가이드북을 업로드하면 해당 대학의 출제 경향과 평가 요소를 반영합니다."
    )
    
    st.divider()
    run = st.button("🚀 면접 질문지 생성 시작", type="primary", use_container_width=True)

# ═══════════════════════════════════════════════════════════
# 실행부
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

    # 대학별 가이드북 텍스트 추출
    univ_guide_text = ""
    if univ_uploaded is not None:
        with st.spinner("대학별 면접 가이드북 분석 중..."):
            univ_guide_text = extract_text_from_pdf(univ_uploaded)

    # 프롬프트 구성
    user_input = f"희망 전공: {major}\n메인 질문 수: {n_main}\n꼬리질문 수: {n_tail}\n\n"

    if univ_guide_text.strip():
        user_input += (
            f"[목표 대학 특화 면접 가이드]\n"
            f"제공된 대학 가이드북의 평가 기준과 스타일을 최우선으로 반영하여 질문을 생성할 것.\n"
            f"{univ_guide_text}\n\n"
        )

    user_input += f"[학생부]\n{record_text}"

    # AI 생성 호출
    with st.spinner(f"{model_name} 모델이 면접 질문지를 생성 중입니다..."):
        try:
            sheet = call_gemini(api_key, model_name, user_input)
        except json.JSONDecodeError:
            st.error("모델이 JSON 형식으로 응답하지 않았습니다. 다시 시도해 주세요.")
            st.stop()
        except Exception as e:
            st.error(f"생성 실패: {e}")
            st.stop()

    questions = sheet.get("questions", [])
    if not questions:
        st.error("질문이 생성되지 않았습니다. 학생부 내용을 더 풍부하게 입력해 주세요.")
        st.stop()

    st.success(f"✅ 메인 질문 {len(questions)}개를 생성했습니다.")

    # ────────────────────────────────────────────────
    # 화면 출력
    # ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("## 📋 생성된 면접 질문지")
    
    for q in questions:
        with st.container(border=True):
            areas = " / ".join(q.get("areas", []))
            intent = " · ".join(q.get("intent", []))
            grade = q.get("grade", "-")

            st.markdown(f"### Q{q['no']}. {q['main']}")
            st.caption(f"활동영역: **{areas}** | 학년: **{grade}** | 의도: **{intent}** | 연계질문: {'🔗 YES' if q.get('linked') else '–'}")

            for i, fu in enumerate(q.get("followups", []), 1):
                st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;**꼬리질문 {i}.** {fu}")

            if q.get("evidence"):
                with st.expander("📎 학생부 근거 보기"):
                    st.info(q["evidence"])

    # ────────────────────────────────────────────────
    # 다운로드 버튼 (Word 1열 배치)
    # ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📥 다운로드")

    meta = {
        "major": major,
        "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    today = datetime.now().strftime("%Y%m%d_%H%M")
    base_name = f"면접질문지_{major}_{today}"

    docx_bytes = build_docx(sheet, meta)
    st.download_button(
        "📝 Word(docx) 다운로드", 
        data=docx_bytes, 
        file_name=f"{base_name}.docx", 
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
        use_container_width=True
    )
