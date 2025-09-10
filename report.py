import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import pdfplumber
import io
import zipfile
from pptx import Presentation
from pptx.util import Inches
from datetime import datetime
st.title(":짠: 이벤트 결과보고서 자동생성 프로그램")
# 여러 파일 업로드 허용
uploaded_files = st.file_uploader(
    ":열린_파일_폴더: 파일 업로드 (Excel 또는 PDF)",
    type=["xlsx", "xls", "pdf"],
    accept_multiple_files=True
)
def _fig_to_png_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()
def analyze_excel(file, file_name):
    df = pd.read_excel(file)
    st.subheader(f":막대_차트: {file_name} 분석 결과")
    # 기본 통계 요약
    st.write(":흰색_확인_표시: 데이터 요약")
    st.write(df.describe(include="all"))
    # 시각화 (숫자형 컬럼)
    num_cols = df.select_dtypes(include="number").columns
    chart_images = []  # [(title, png_bytes)]
    for col in num_cols:
        st.write(f":상승세인_차트: {col} 분포")
        fig, ax = plt.subplots()
        df[col].hist(ax=ax, bins=20)
        ax.set_title(f"{file_name} · {col} 분포")
        ax.set_xlabel(col)
        ax.set_ylabel("빈도")
        st.pyplot(fig)
        chart_images.append((f"{col} 분포", _fig_to_png_bytes(fig)))
    # 숫자형이 2개 이상이면 상관행렬
    if len(num_cols) >= 2:
        fig, ax = plt.subplots()
        corr = df[num_cols].corr(numeric_only=True)
        cax = ax.imshow(corr, aspect="auto")
        ax.set_xticks(range(len(num_cols)))
        ax.set_yticks(range(len(num_cols)))
        ax.set_xticklabels(num_cols, rotation=90)
        ax.set_yticklabels(num_cols)
        ax.set_title(f"{file_name} · 숫자형 상관관계")
        fig.colorbar(cax)
        st.pyplot(fig)
        chart_images.append(("숫자형 상관관계", _fig_to_png_bytes(fig)))
    return df, chart_images
def analyze_pdf(file, file_name):
    st.subheader(f":글씨가_쓰여진_페이지: {file_name} 텍스트 추출")
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"
    st.text_area(":책갈피_탭: 추출된 텍스트", text, height=200)
    return text, []  # PDF 차트 없음
def make_ppt_report(title: str, all_charts: dict) -> bytes:
    prs = Presentation()
    # 제목
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = f"자동 생성 · {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    # 차트 슬라이드
    for dataset_name, charts in all_charts.items():
        s = prs.slides.add_slide(prs.slide_layouts[5])
        s.shapes.title.text = f":포장: {dataset_name}"
        for chart_title, png_bytes in charts:
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = chart_title
            left = Inches(1); top = Inches(1.2); width = Inches(8)
            slide.shapes.add_picture(io.BytesIO(png_bytes), left, top, width=width)
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()
def make_excel_with_images(all_dfs, all_charts) -> bytes:
    """
    Data 시트 + Charts 시트(이미지 삽입) 형태로 엑셀 저장.
    - all_dfs: [DataFrame, ...]  (Data_1, Data_2 ...)
    - all_charts: {"파일명": [(title, png_bytes), ...], ...}
    """
    out = io.BytesIO()
    # :경고: 이미지 삽입을 위해 engine='xlsxwriter' 사용
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        # 1) 데이터 시트 저장
        for i, df in enumerate(all_dfs, start=1):
            sheet_name = f"Data_{i}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            # 가독성을 위한 약간의 서식 (선택)
            ws = writer.sheets[sheet_name]
            for col_idx, col in enumerate(df.columns):
                max_len = max([len(str(col))] + [len(str(x)) for x in df[col].head(100).astype(str).tolist()])
                ws.set_column(col_idx, col_idx, min(max_len + 2, 40))
        # 2) 차트 시트 (이미지 삽입)
        chart_ws = writer.book.add_worksheet("Charts")
        # 시트 폭 조금 넓히기
        for c in range(0, 20):
            chart_ws.set_column(c, c, 14)
        # 데이터셋 섹션 타이틀 + 이미지들
        r = 1
        c = 1
        per_row = 2  # 한 행에 이미지 2개
        idx_in_row = 0
        for ds_name, charts in all_charts.items():
            chart_ws.write(r, c, f":포장: {ds_name}")
            r += 1
            for (title, png_bytes) in charts:
                chart_ws.write(r, c, f"• {title}")
                chart_ws.insert_image(
                    r + 1, c,
                    "chart.png",
                    {"image_data": io.BytesIO(png_bytes), "x_scale": 1.0, "y_scale": 1.0}
                )
                idx_in_row += 1
                if idx_in_row % per_row == 0:
                    r += 20
                    c = 1
                else:
                    c += 8
            r += 22
            c = 1
            idx_in_row = 0
    return out.getvalue()
if uploaded_files:
    all_dfs = []           # [DataFrame, ...]
    all_texts = []         # [str, ...]
    all_charts = {}        # { "파일명": [(title, png_bytes), ...] }
    for file in uploaded_files:
        file_name = file.name
        if file_name.endswith(("xlsx", "xls")):
            df, charts = analyze_excel(file, file_name)
            all_dfs.append(df)
            all_charts[file_name] = charts
        elif file_name.endswith("pdf"):
            text, charts = analyze_pdf(file, file_name)
            all_texts.append(text)
    if st.button(":받은_편지함_트레이: 결과 보고서 생성"):
        # 1) Markdown (표/텍스트)
        md_content = "# :다트: 이벤트 결과 보고서\n\n"
        md_content += ":반짝임: 자동 생성된 요약 리포트입니다.\n\n"
        for idx, df in enumerate(all_dfs):
            md_content += f"## :막대_차트: 데이터셋 {idx+1}\n"
            md_content += df.describe(include="all").to_markdown() + "\n\n"
        for idx, txt in enumerate(all_texts):
            md_content += f"## :글씨가_쓰여진_페이지: PDF 문서 {idx+1}\n"
            md_content += (txt[:1000] + "...\n\n") if txt else "내용 없음\n\n"
        st.download_button(
            ":받은_편지함_트레이: Markdown 보고서 다운로드",
            data=md_content,
            file_name="event_report.md"
        )
        # 2) Excel (데이터 + Charts 시트에 이미지 삽입)
        excel_with_imgs = make_excel_with_images(all_dfs, all_charts)
        st.download_button(
            ":막대_차트: Excel 보고서(차트 내장) 다운로드",
            data=excel_with_imgs,
            file_name="event_report_with_charts.xlsx"
        )
        # 3) PPT (차트 포함)
        if any(len(v) > 0 for v in all_charts.values()):
            ppt_bytes = make_ppt_report("이벤트 결과 보고서 :반짝임:", all_charts)
            st.download_button(
                ":영사기: PPT 보고서(차트 포함) 다운로드",
                data=ppt_bytes,
                file_name="event_report.pptx"
            )
        # 4) 선택: 차트 PNG ZIP
        if any(len(v) > 0 for v in all_charts.values()):
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                for ds_name, charts in all_charts.items():
                    safe = ds_name.replace("/", "_")
                    for i, (ctitle, png) in enumerate(charts, start=1):
                        zf.writestr(f"{safe}/chart_{i:02d}_{ctitle}.png", png)
            st.download_button(
                ":액자에_담긴_그림: 차트 PNG 묶음(ZIP) 다운로드",
                data=zip_buf.getvalue(),
                file_name="charts.zip"
            )
