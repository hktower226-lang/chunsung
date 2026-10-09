import re
import pandas as pd
import streamlit as st

# ==========================================
# 1. 모바일 최적화 레이아웃 및 CSS 디자인
# ==========================================
st.set_page_config(
    page_title="현장 점유율 및 채용 비율",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    /* 전체 글씨 크기 및 가독성 최적화 */
    html, body, [class*="css"] { font-size: 19px !important; }
    h1 { font-size: 2.0rem !important; font-weight: bold !important; color: #1E3A8A; }
    h2 { font-size: 1.6rem !important; font-weight: bold !important; color: #1E40AF; }
    h3 { font-size: 1.3rem !important; font-weight: bold !important; }

    /* 여백 조정 (세로 스크롤 최적화) */
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 3rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
    }

    /* 메인 세로 카드 디자인 */
    .main-total-card {
        background: linear-gradient(135deg, #1E3A8A, #3B82F6);
        color: white;
        padding: 20px;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .main-stat-card {
        background-color: #FFFFFF;
        border-left: 6px solid #3B82F6;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }

    /* 각 소속별 전용 색상 테두리 */
    .border-hanno { border-left-color: #EF4444 !important; }  /* 빨강 */
    .border-minno { border-left-color: #3B82F6 !important; }  /* 파랑 */
    .border-seomoo { border-left-color: #10B981 !important; } /* 초록 */
    .border-geunsan { border-left-color: #F59E0B !important; }/* 주황 */
    .border-jikwon { border-left-color: #8B5CF6 !important; } /* 보라 */
    .border-mijeong { border-left-color: #6B7280 !important; }/* 회색 */

    /* 요약 카드 박스 */
    .card-box {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 12px;
    }

    /* 전화걸기 버튼 */
    .phone-btn {
        display: inline-block;
        background-color: #25D366;
        color: white !important;
        padding: 12px 18px;
        font-size: 18px;
        font-weight: bold;
        text-decoration: none;
        border-radius: 8px;
        margin-top: 8px;
        text-align: center;
        width: 100%;
        box-shadow: 0 2px 5px rgba(0,0,0,0.15);
    }
    </style>
""",
    unsafe_allow_html=True,
)

EXCEL_FILE = "채용비율.xlsx"


# ==========================================
# 2. 엑셀 데이터 로드 함수
# ==========================================
@st.cache_data(ttl=5)
def load_sheets():
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        sheets = {
            sheet: pd.read_excel(xls, sheet, header=None)
            for sheet in xls.sheet_names
        }
        return sheets
    except Exception as e:
        st.error(
            f"❌ 엑셀 파일('{EXCEL_FILE}')을 읽는 중 에러가 발생했습니다: {e}"
        )
        return None


sheets = load_sheets()

if not sheets:
    st.stop()

sheet_keys = list(sheets.keys())


# 퍼센트 변환 함수 (0.28169 -> 28.2%)
def to_pct_str(val):
    try:
        if pd.isna(val) or val == "" or val == "-":
            return "0%"
        f_val = float(val)
        if f_val <= 1.0:
            return f"{f_val * 100:.1f}%"
        return f"{f_val:.1f}%"
    except:
        return str(val)


# ==========================================
# 3. 사이드바 메뉴 (요청하신 메뉴명 변경)
# ==========================================
st.sidebar.title("📌 채용 비율 메뉴 선택")
menu = st.sidebar.radio(
    "원하시는 화면을 선택하세요",
    [
        "1️⃣ 메인: 전체 소속별 점유율",
        "2️⃣ 8개 지부 현장 점유율",
        "3️⃣ 타워사별 점유 현황",
        "4️⃣ 반도체 현장 파악",
        "5️⃣ 1~9월 채용 현황 및 추이",
    ],
)


# ==========================================
# 4. 메뉴별 화면 구현
# ==========================================

# ------------------------------------------
# [메뉴 1] 메인: 전체 소속별 점유율 (세로 레이아웃 & 색상 적용)
# ------------------------------------------
if menu == "1️⃣ 메인: 전체 소속별 점유율":
    st.title("📊 2026년 전체 소속별 점유율")

    raw_df1 = sheets[sheet_keys[0]]

    # 원본 데이터 추출
    # row 2: 소속 수치 (한노 59, 민노 87, ..., 합계 209)
    # row 3: 소속 비율 (한노 0.28169, 민노 0.416, ...)
    try:
        tot_cnt = raw_df1.iloc[2, 7]  # 209
        tot_pct = raw_df1.iloc[2, 8]  # 100%

        items = [
            {
                "name": "한국노총 (한노)",
                "cnt": raw_df1.iloc[2, 1],
                "pct": to_pct_str(raw_df1.iloc[3, 1]),
                "class": "border-hanno",
                "color": "#EF4444",
            },
            {
                "name": "민주노총 (민노)",
                "cnt": raw_df1.iloc[2, 2],
                "pct": to_pct_str(raw_df1.iloc[3, 2]),
                "class": "border-minno",
                "color": "#3B82F6",
            },
            {
                "name": "섬유노조",
                "cnt": raw_df1.iloc[2, 3],
                "pct": to_pct_str(raw_df1.iloc[3, 3]),
                "class": "border-seomoo",
                "color": "#10B981",
            },
            {
                "name": "건설산업 (건산)",
                "cnt": raw_df1.iloc[2, 4],
                "pct": to_pct_str(raw_df1.iloc[3, 4]),
                "class": "border-geunsan",
                "color": "#F59E0B",
            },
            {
                "name": "직원",
                "cnt": raw_df1.iloc[2, 5],
                "pct": to_pct_str(raw_df1.iloc[3, 5]),
                "class": "border-jikwon",
                "color": "#8B5
