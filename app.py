"""
학생부 종합 면접 어시스트 v3.0
- 학년/과목 간 관심사 확장을 추적하는 '연계질문' 특화 로직 반영
- 화면 표시 / DOCX 다운로드 지원 (PDF, JSON 제거 버전)
- 하단 고정 푸터(만든 이) 추가
"""

import streamlit as st
import json
import io
from datetime import datetime
from pathlib import Path

from google import genai
from google.genai import types

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pypdf

# ═══════════════════════════════════════════════════════════
# 페이지 설정
# ═══════════════════════════════════════════════════════════
st.set_page_config(
    page_title="학생부 면접 질문 생성기 시스템 v3.0",
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
    <b>🏫 학생부 면접 질문 생성기 시스템 v3.0</b><br>
    만든 이: 신선여자고등학교 김명남<br>
    🗓️ 2026.05
</div>
"""
st.markdown(footer_css, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════
# 시스템 프롬프트 (연계질문 특화)
# ═══════════════════════════════════════════════════════════
SYSTEM_PROMPT = """
당신은 대학 입학사정관 및 면접 출제 전문가입니다. 제공된 학생부 내용을 꼼꼼히 분석하여, 학생의 특정 관심사나 역량이 서로 다른 학년이나 과목(영역)으로 이어지는 '연계 지점'을 찾아 면접 질문을 생성하세요.

[출제 원칙]
1. 메인 질문: 특정 학년/과목에서 시작된 핵심 관심사나 활동을 묻습니다.
2. 연계 질문: 메인 질문의 주제가 다른 학년이나 다른 과목에서 어떻게 확장, 심화, 또는 적용되었는지 묻습니다.
3. 질문 끝에는 반드시 기재된 출처(예: 1학년 통합사회, 2학년 생명과학 I)를 괄호로 명시해야 합니다.
4. 각 질문에 대해 학생부에 기록된 사실을 바탕으로 '모범 답변(또는 활동 내용 요약)'을 함께 제시하세요.

[필수 JSON 출력 형식]
반드시 아래의 JSON 구조로만 출력하세요. 마크다운이나 다른 텍스트는 절대 포함하지 마세요.
{
  "questions": [
    {
      "no": 1,
      "main_question": "유전공학에 대한 관심으로 유전자 조작 유기체 문제를 사회 문제로 선정하여 조사하고 발표함, 어떤 내용?",
      "main_context": "(1학년 통합사회)",
      "linked_questions": [
        {
          "sub_no": "1-1",
          "question": "유전자 조작 기술에 대한 내용을 탐구하며 장기 이식용 유전자 조작 동물 생산이 장기 부족 문제 해결에 기여할 수 있다는 점에 대해 알게 됨, 어떤 내용?",
          "context": "(2학년 생명과학 I)",
          "expected_answer": "유전자 조작 돼지가 생산된 사례와 국외 대학 연구팀의 형질전환 돼지 심장 이식 수술 성공 사례를 통해 기술적 가능성과 의료적 가치 확인함."
        },
        {
          "sub_no": "1-2",
          "question": "유전자 변형 식품에 관한 찬반 토론에서 유전자 변형 식품의 다양한 장점을 근거로 제시함, 어떤 내용?",
          "context": "(2학년 생명과학 I)",
          "expected_answer": "식량 부족 문제에 도움, 농약 사용을 줄임으로써 환경 보호에 도움."
        }
      ]
    }
  ]
}
"""

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
            temperature=0.3, # 창의성보다 사실(학생부) 기반 추출을 위해 온도 낮춤
            response_mime_type="application/json",
        ),
    )
    raw = response.text.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

def build_docx(sheet: dict, meta: dict) -> bytes:
    """질문지를 DOCX 파일로 변환 (연계질문 포맷 적용)"""
    doc = Document()
    
    style = doc.styles["Normal"]
    style.font.name = "맑은 고딕"
    style.font.size = Pt(11)

    title = doc.add_heading("모의 면접 질문지", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.add_run(f"희망 전공: {meta['major']} | 생성일: {meta['date']} | 메인 세트 {len(sheet['questions'])}개").italic = True
    doc.add_paragraph("─" * 50)

    for q in sheet["questions"]:
        # 메인 질문
        mp = doc.add_paragraph()
        run = mp.add_run(f"{q['no']}. {q.get('main_question', '')} {q.get('main_context', '')}")
        run.font.size = Pt(12)

        # 연계 질문
        if q.get("linked_questions"):
            doc.add_paragraph("[연계질문]").bold = True
            
            for lq in q["linked_questions"]:
                # 서브 질문
                lp = doc.add_paragraph()
                lp.add_run(f"{lq.get('sub_no', '')}. {lq.get('question', '')} {lq.get('context', '')}")
                
                # 기대 답변 (앞에 : 붙이고 들여쓰기)
                ap = doc.add_paragraph()
                ap.add_run(f": {lq.get('expected_answer', '')}")
                # ap.paragraph_format.left_indent = Cm(0.5)
                
        doc.add_paragraph() # 문항 간 간격

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ═══════════════════════════════════════════════════════════
# 사이드바: 설정
# ═══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🔑 기본 설정")
    api_key = st.text_input("Google AI API 키", type="password")
    
    st.divider()
    st.markdown("### ⚙️ 모델 설정")
    model_name = st.selectbox(
        "Gemini 모델",
        ["gemini-2.5-flash", "gemini-2.5-pro"],
        index=0,
    )

    st.divider()
    st.markdown("### 📊 질문지 옵션")
    n_main = st.slider("메인 질문 세트 수", 3, 10, 5)

# ═══════════════════════════════════════════════════════════
# 메인 레이아웃
# ═══════════════════════════════════════════════════════════
st.title("🎤 학생부 면접 질문 생성기 시스템 v3.0")
st.caption("학생부 내 학년/영역 간 관심사 확장을 추적하는 '융합형 연계질문' 생성 시스템")

col_left, col_right = st.columns([1.2, 1])

with col_left:
    st.subheader("1. 학생부 데이터 입력")
    uploaded = st.file_uploader("📁 학생부 PDF 업로드 (선택)", type=["pdf"])
    
    extracted = ""
    if uploaded is not None:
        with st.spinner("PDF에서 텍스트 추출 중..."):
            extracted = extract_text_from_pdf(uploaded)
        if extracted.strip():
            st.success(f"PDF 추출 완료!")

    record_text = st.text_area(
        "📝 학생부 내용 (직접 붙여넣기 권장)",
        value=extracted, 
        height=400,
        placeholder="[1학년 통합사회] ...\n[2학년 생명과학I] ..."
    )

with col_right:
    st.subheader("2. 분석 옵션")
    major = st.text_input("🎓 희망 전공 / 학과", placeholder="예: 생명공학과")
    
    st.markdown("---")
    st.markdown("**🏫 목표 대학 면접 스타일 (선택)**")
    univ_uploaded = st.file_uploader("대학별 면접 가이드북 업로드", type=["pdf"])
    
    st.divider()
    run = st.button("🚀 연계형 면접 질문지 생성", type="primary", use_container_width=True)

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

    # 프롬프트 세팅
    user_input = f"희망 전공: {major}\n생성할 연계 질문 세트 수: {n_main}\n\n"
    
    univ_guide_text = ""
    if univ_uploaded is not None:
        univ_guide_text = extract_text_from_pdf(univ_uploaded)
        if univ_guide_text.strip():
            user_input += f"[대학 가이드북 참고]\n{univ_guide_text}\n\n"

    user_input += f"[학생부 원문]\n{record_text}"

    # AI 생성
    with st.spinner("학생부 전반의 맥락을 분석하여 연계질문을 추출 중입니다..."):
        try:
            sheet = call_gemini(api_key, model_name, user_input)
        except Exception as e:
            st.error(f"생성 실패: {e}")
            st.stop()

    questions = sheet.get("questions", [])
    if not questions:
        st.error("질문이 생성되지 않았습니다.")
        st.stop()

    st.success(f"✅ 총 {len(questions)}개의 연계 질문 세트를 생성했습니다.")

    # ────────────────────────────────────────────────
    # 화면 출력 (첨부 이미지 스타일로 구현)
    # ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("## 📋 생성된 연계 면접 질문지")
    
    for q in questions:
        with st.container(border=True):
            # 메인 질문
            st.markdown(f"#### {q.get('no', '')}. {q.get('main_question', '')} **{q.get('main_context', '')}**")
            
            # 연계 질문
            if q.get("linked_questions"):
                st.markdown("**[연계질문]**")
                for lq in q["linked_questions"]:
                    st.markdown(f"{lq.get('sub_no', '')}. {lq.get('question', '')} **{lq.get('context', '')}**")
                    st.caption(f": {lq.get('expected_answer', '')}")

    # ────────────────────────────────────────────────
    # 다운로드 버튼
    # ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📥 다운로드")

    meta = {
        "major": major,
        "date": datetime.now().strftime("%Y-%m-%d"),
    }
    base_name = f"연계면접질문지_{major}_{datetime.now().strftime('%H%M')}"

    docx_bytes = build_docx(sheet, meta)
    st.download_button(
        "📝 Word(docx) 다운로드", 
        data=docx_bytes, 
        file_name=f"{base_name}.docx", 
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
        use_container_width=True
    )
