# interview-question
# 🎤 학생부 기반 모의 면접 질문지 생성기

동국대학교 학생부위주전형 가이드북의 면접 출제 원리(**WHAT-WHY-HOW-SO WHAT**, 영역 연계, 활동 내부 심화)를 반영하여, 학생부 텍스트로부터 모의 면접 질문지를 자동 생성하는 Streamlit 앱.

## ✨ 주요 기능
- 학생부 PDF 업로드 또는 직접 붙여넣기
- 메인 질문 수 / 꼬리질문 수 조절
- 영역 연계형 심화 질문 자동 생성
- **화면 표시 + PDF + Word(docx) + JSON** 동시 다운로드
- 각 질문에 학생부 근거 인용 포함 (검수 용이)

## 🛠 설치
```bash
git clone https://github.com/<your-id>/interview-question-generator.git
cd interview-question-generator
pip install -r requirements.txt
```

## 🔤 한글 PDF 폰트 추가 (필수)
1. [네이버 나눔 글꼴 페이지](https://hangeul.naver.com/font)에서 **NanumGothic.ttf** 다운로드
2. `fonts/NanumGothic.ttf` 경로에 저장
3. (없어도 Word/JSON 다운로드는 정상 작동)

## 🚀 실행
```bash
streamlit run app.py
```

## 🔑 Google AI API 키 발급
- https://aistudio.google.com/apikey 에서 무료 발급
- 앱 좌측 사이드바에 입력

## 🌐 Streamlit Cloud 배포
1. 이 레포를 본인 GitHub 계정에 push
2. https://share.streamlit.io → New app → 레포 선택
3. Main file path: `app.py`
4. Deploy

## ⚠️ 개인정보 안내
- 학생 실명·학번 등은 마스킹 후 입력해 주세요.
- 입력 데이터는 Google Gemini API로 전송되며, 서버에 저장되지 않습니다.

## 📝 프롬프트 수정
면접 질문 스타일을 바꾸고 싶으면 `prompts/system_rules.md` 파일만 수정하세요. `app.py`는 손대지 않아도 됩니다.
