# 이벤트 결과보고서 자동생성 (Streamlit)

업로드한 **Excel/PDF**를 분석하고, **Markdown/Excel/PPT/ZIP** 보고서를 자동으로 만들어 주는 앱입니다.

## 바로 실행 (로컬)

```bash
pip install -r requirements.txt
streamlit run report.py
```

## 파일 설명
- `report.py` — Streamlit 앱 본문
- `requirements.txt` — 의존성 목록
- `.gitignore` — 불필요 파일 제외
- `.streamlit/config.toml` — (선택) Streamlit UI 설정

## Streamlit Cloud 배포
1. 이 저장소를 GitHub에 올립니다.
2. Streamlit Cloud에서 **New app** → 저장소/브랜치 선택 → Main file: `report.py`
3. Deploy

## 폴더/파일 업로드 가이드
- Excel: `.xlsx`, `.xls`
- PDF: 텍스트 추출은 **pdfplumber** 기반

## 라이선스
MIT
