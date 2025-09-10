# report.py — Plotly 기반 / 폰트(MaruBuri) / PPT & Word 내보내기 지원
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.figure_factory as ff
import pdfplumber
import io
import zipfile
from pptx import Presentation
from pptx.util import Inches as PptxInches
from datetime import datetime

# Word(.docx)
from docx import Document
from docx.shared import Inches as DocxInches

# ===== 한글 폰트 설정 (Matplotlib 대비 및 Plotly 폰트 패밀리 지정용) =====
from matplotlib import font_manager, rcParams
import os

# 프로젝트 루트/fonts/MaruBuri-Regular.ttf 사용
FONT_PATH = os.path.join(os.path.dirname(__file__), "fonts", "MaruBuri-Regular.ttf")
if os.path.exists(FONT_PATH):
    try:
        font_manager.fontManager.addfont(FONT_PATH)
        rcParams["font.family"] = "MaruBuri"
    except Exception:
        rcParams["font.family"] = ["Malgun Gothic", "AppleGothic", "Noto Sans CJK KR", "NanumGothic", "DejaVu Sans"]
else:
    rcParams["font.family"] = ["Malgun Gothic", "AppleGothic", "Noto Sans CJK KR", "NanumGothic", "DejaVu Sans"]
rcParams["axes.unicode_minus"] = False

# Plotly에서 사용할 공통 폰트 패밀리 (브라우저에 없으면 다음 폰트로 폴백)
PLOTLY_FONT_FAMILY = "MaruBuri, NanumGothic, 'Malgun Gothic', AppleGothic, 'Noto Sans CJK KR', 'DejaVu Sans', sans-serif"


def _style_plotly(fig, title=None):
    fig.update_layout(
        template="plotly_white",
        title=title if title else fig.layout.title.text,
        title_x=0.5,
        title_font_size=18,
        font=dict(family=PLOTLY_FONT_FAMILY),
        margin=dict(l=40, r=20, t=60, b=40),
    )
    return fig


def _fig_to_png_bytes(fig):
    """
    Plotly figure -> PNG bytes (PPT/Excel/Word 삽입용).
    kaleido 필요. (pip install kaleido)
    """
    try:
        return fig.to_image(format="png", scale=2)
    except Exception:
        return None


st.title("✨ 이벤트 결과보고서 자동생성 프로그램 (Plotly 개선판)")

# ===== 차트 옵션 (사이드바) =====
st.sidebar.header("차트 옵션")
BIN_MODE = st.sidebar.radio("빈 구분 방법", ["자동", "개수 지정", "간격 지정"], index=0, horizontal=True)

nbins = st.sidebar.slider("빈 개수", 5, 100, 20) if BIN_MODE == "개수 지정" else None
binsize = st.sidebar.number_input("빈 간격(숫자)", min_value=0.0, value=0.0, step=100.0) if BIN_MODE == "간격 지정" else None

bargap = st.sidebar.slider("막대 간격(bargap)", 0.00, 0.50, 0.20, 0.01)
show_kde = st.sidebar.checkbox("밀도 곡선(KDE) 표시", value=True)
y_scale = st.sidebar.selectbox("세로축 단위", ["count", "percent", "probability density"], index=0)
# =================================

# 여러 파일 업로드 허용
uploaded_files = st.file_uploader(
    "📂 파일 업로드 (Excel 또는 PDF)",
    type=["xlsx", "xls", "pdf"],
    accept_multiple_files=True
)


def analyze_excel(file, file_name):
    df = pd.read_excel(file)
    st.subheader(f"📊 {file_name} 분석 결과")

    # 기본 통계 요약
    st.write("✅ 데이터 요약")
    try:
        st.markdown(df.describe(include="all").to_markdown())
    except Exception:
        st.text(df.describe(include="all").to_string())

    num_cols = df.select_dtypes(include="number").columns
    chart_images = []

    for col in num_cols:
        st.write(f"📈 {col} 분포")

        # KDE를 그릴 경우에는 density 스케일 권장
        histnorm = y_scale
        if show_kde and histnorm in ("count", None):
            histnorm = "probability density"

        # 히스토그램
        fig = px.histogram(
            df,
            x=col,
            nbins=nbins if BIN_MODE == "개수 지정" else None,
            color_discrete_sequence=["#4C78A8"],
            histnorm=None if histnorm == "count" else histnorm,
        )
        # 간격 지정 모드면 bin 간격 강제
        if BIN_MODE == "간격 지정" and binsize and binsize > 0:
            fig.update_traces(xbins=dict(size=binsize))

        # 막대 간격/외곽선/투명도 → 덜 답답하고 깔끔하게
        fig.update_traces(
            marker_line_color="white",
            marker_line_width=1,
            opacity=0.9,
        )
        fig.update_layout(
            bargap=bargap,          # 막대 사이 간격
            bargroupgap=0.05,       # 그룹 간 간격(여러 시리즈일 때)
        )

        # 공통 스타일
        fig.update_xaxes(title_text=col)
        fig.update_yaxes(title_text="빈도" if histnorm in (None, "count") else histnorm.title())
        _style_plotly(fig, title=f"{file_name} · {col} 분포")

        # ── KDE(밀도 곡선) 오버레이 ──
        if show_kde:
            vals = df[col].dropna().values
            if len(vals) > 1 and np.isfinite(vals).all():
                kde_fig = ff.create_distplot(
                    [vals], [col],
                    show_hist=False, show_rug=False,
                    colors=["#E45756"], curve_type="kde"
                )
                # ff가 만든 곡선(trace)을 현재 fig에 추가
                for tr in kde_fig.data:
                    if tr.type == "scatter":
                        tr.update(name="KDE", line=dict(width=2))
                        fig.add_trace(tr)

        # 출력
        st.plotly_chart(fig, use_container_width=True)

        # PPT/엑셀/워드용 PNG 저장
        png = _fig_to_png_bytes(fig)
        if png is not None:
            chart_images.append((f"{col} 분포", png))

    # 상관관계 히트맵
    if len(num_cols) >= 2:
        corr = df[num_cols].corr()
        z = corr.values
        x = corr.columns.tolist()
        y = corr.columns.tolist()
        ann = corr.round(2).values

        heat = ff.create_annotated_heatmap(
            z=z, x=x, y=y, annotation_text=ann,
            colorscale="RdBu", showscale=True, reversescale=True,
        )
        _style_plotly(heat, title=f"{file_name} · 숫자형 상관관계")
        st.plotly_chart(heat, use_container_width=True)

        png = _fig_to_png_bytes(heat)
        if png is not None:
            chart_images.append(("숫자형 상관관계", png))

    return df, chart_images


def analyze_pdf(file, file_name):
    st.subheader(f"📄 {file_name} 텍스트 추출")
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"
    st.text_area("📑 추출된 텍스트", text, height=200)
    return text, []  # Plotly 차트 없음


def make_ppt_report(title: str, all_charts: dict) -> bytes:
    prs = Presentation()
    # 제목 슬라이드
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = f"자동 생성 · {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    # 데이터셋별 차트 슬라이드
    for dataset_name, charts in all_charts.items():
        # 섹션 타이틀
        s = prs.slides.add_slide(prs.slide_layouts[5])
        s.shapes.title.text = f"📦 {dataset_name}"

        # 개별 차트
        for chart_title, png_bytes in charts:
            if not png_bytes:
                continue
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = chart_title
            left = PptxInches(0.8)
            top = PptxInches(1.2)
            width = PptxInches(8.4)
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
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        # 1) 데이터 시트 저장
        for i, df in enumerate(all_dfs, start=1):
            sheet_name = f"Data_{i}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            # 가독성 컬럼 폭
            ws = writer.sheets[sheet_name]
            for col_idx, col in enumerate(df.columns):
                max_len = max([len(str(col))] + [len(str(x)) for x in df[col].head(100).astype(str).tolist()])
                ws.set_column(col_idx, col_idx, min(max_len + 2, 40))

        # 2) 차트 시트 (이미지 삽입)
        chart_ws = writer.book.add_worksheet("Charts")
        for c in range(0, 24):
            chart_ws.set_column(c, c, 14)

        r = 1
        c = 1
        per_row = 2
        idx_in_row = 0
        for ds_name, charts in all_charts.items():
            chart_ws.write(r, c, f"📦 {ds_name}")
            r += 1
            for (title, png_bytes) in charts:
                if not png_bytes:
                    continue
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


def make_word_report(title: str, all_dfs: list[pd.DataFrame], all_texts: list[str], all_charts: dict) -> bytes:
    """
    python-docx 로 .docx 생성
    - 제목, 데이터 요약 테이블, PDF 일부 텍스트, 차트 이미지 삽입
    """
    doc = Document()
    doc.add_heading(title, level=1)
    doc.add_paragraph(f"자동 생성 · {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    # 데이터 요약 표
    for i, df in enumerate(all_dfs, start=1):
        doc.add_heading(f"데이터셋 {i} 요약", level=2)
        desc = df.describe(include="all")
        # 표로 렌더링
        table = doc.add_table(rows=1 + len(desc.index), cols=1 + len(desc.columns))
        table.style = "Light List Accent 1"
        # 헤더
        hdr = table.rows[0].cells
        hdr[0].text = ""
        for j, col in enumerate(desc.columns, start=1):
            hdr[j].text = str(col)
        # 바디
        for r_idx, idx_name in enumerate(desc.index, start=1):
            row = table.rows[r_idx].cells
            row[0].text = str(idx_name)
            for c_idx, col in enumerate(desc.columns, start=1):
                val = desc.loc[idx_name, col]
                row[c_idx].text = "" if pd.isna(val) else str(val)

    # PDF 텍스트
    for i, txt in enumerate(all_texts, start=1):
        doc.add_heading(f"PDF 문서 {i}", level=2)
        doc.add_paragraph(txt[:1500] + ("..." if len(txt) > 1500 else ""))

    # 차트 이미지
    for ds_name, charts in all_charts.items():
        doc.add_heading(f"차트 - {ds_name}", level=2)
        for (title, png_bytes) in charts:
            if not png_bytes:
                continue
            doc.add_paragraph(f"• {title}")
            doc.add_picture(io.BytesIO(png_bytes), width=DocxInches(6.5))

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ====================== 메인 로직 ======================
if uploaded_files:
    all_dfs = []           # [DataFrame, ...]
    all_texts = []         # [str, ...]
    all_charts = {}        # { "파일명": [(title, png_bytes), ...] }

    for file in uploaded_files:
        file_name = file.name
        if file_name.lower().endswith(("xlsx", "xls")):
            df, charts = analyze_excel(file, file_name)
            all_dfs.append(df)
            all_charts[file_name] = charts
        elif file_name.lower().endswith("pdf"):
            text, charts = analyze_pdf(file, file_name)
            all_texts.append(text)

    if st.button("📥 결과 보고서 생성"):
        # 1) Markdown (표/텍스트)
        md_content = "# 🎯 이벤트 결과 보고서\n\n"
        md_content += "✨ 자동 생성된 요약 리포트입니다.\n\n"
        for idx, df in enumerate(all_dfs):
            md_content += f"## 📊 데이터셋 {idx+1}\n"
            try:
                md_content += df.describe(include="all").to_markdown() + "\n\n"
            except Exception:
                md_content += df.describe(include="all").to_string() + "\n\n"

        for idx, txt in enumerate(all_texts):
            md_content += f"## 📄 PDF 문서 {idx+1}\n"
            md_content += (txt[:1000] + "...\n\n") if txt else "내용 없음\n\n"

        st.download_button(
            "📥 Markdown 보고서 다운로드",
            data=md_content,
            file_name="event_report.md"
        )

        # 2) Excel (데이터 + Charts 시트에 이미지 삽입)
        excel_with_imgs = make_excel_with_images(all_dfs, all_charts)
        st.download_button(
            "📊 Excel 보고서(차트 내장) 다운로드",
            data=excel_with_imgs,
            file_name="event_report_with_charts.xlsx"
        )

        # 3) PPT (차트 포함) - PNG 생성 실패 차트는 건너뜀
        if any(len(v) > 0 for v in all_charts.values()):
            ppt_bytes = make_ppt_report("이벤트 결과 보고서 ✨", all_charts)
            st.download_button(
                "📽 PPT 보고서(차트 포함) 다운로드",
                data=ppt_bytes,
                file_name="event_report.pptx"
            )

        # 4) Word(.docx) 보고서 (차트 포함)
        docx_bytes = make_word_report("이벤트 결과 보고서 ✨", all_dfs, all_texts, all_charts)
        st.download_button(
            "📝 Word 보고서(.docx) 다운로드",
            data=docx_bytes,
            file_name="event_report.docx"
        )

        # 5) 선택: 차트 PNG ZIP (PNG 생성된 차트만 포함)
        if any(len(v) > 0 for v in all_charts.values()):
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                for ds_name, charts in all_charts.items():
                    safe = ds_name.replace("/", "_")
                    for i, (ctitle, png) in enumerate(charts, start=1):
                        if not png:
                            continue
                        zf.writestr(f"{safe}/chart_{i:02d}_{ctitle}.png", png)
            st.download_button(
                "🖼 차트 PNG 묶음(ZIP) 다운로드",
                data=zip_buf.getvalue(),
                file_name="charts.zip"
            )
