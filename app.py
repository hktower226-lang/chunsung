import re
import pandas as pd
import streamlit as st

# ==========================================
# 1. 페이지 기본 설정 및 모바일 CSS 스타일링
# ==========================================
st.set_page_config(
    page_title="현장 점유율 및 채용 현황", layout="wide", initial_sidebar_state="collapsed"
)

# 모바일용 가독성 증대 및 한 화면 최적화 CSS
st.markdown(
    """
    <style>
    /* 전체 글꼴 및 기본 폰트 크기 */
    html, body, [class*="css"] {
        font-size: 16px !important;
    }
    
    /* 제목 및 헤더 여백 축소 */
    h1 { font-size: 1.5rem !important; font-weight: bold !important; color: #1E3A8A; margin-bottom: 5px !important; margin-top: 0px !important; }
    h2 { font-size: 1.3rem !important; font-weight: bold !important; color: #1E40AF; margin-bottom: 5px !important; }
    h3 { font-size: 1.0rem !important; font-weight: bold !important; margin-top: 5px !important; margin-bottom: 5px !important; }
    
    /* 라벨 여백 최소화 */
    .stSelectbox label, .stRadio label, .stMultiSelect label {
        font-size: 1.0rem !important;
        font-weight: bold !important;
    }
    
    /* 여백 최소화 (위아래 세로 스크롤 최소화) */
    .block-container {
        padding-top: 0.6rem !important;
        padding-bottom: 1rem !important;
        padding-left: 0.5rem !important;
        padding-right: 0.5rem !important;
        max-width: 100% !important;
    }
    
    /* 테이블 글씨 크기 */
    .dataframe {
        font-size: 14px !important;
    }
    
    /* 전화번호 연결 버튼 스타일 */
    .phone-btn {
        display: inline-block;
        background-color: #25D366;
        color: white !important;
        padding: 8px 14px;
        font-size: 15px;
        font-weight: bold;
        text-decoration: none;
        border-radius: 6px;
        margin-top: 4px;
        margin-bottom: 6px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.15);
        width: 100%;
    }

    /* 메인 화면 슬림 카드 디자인 */
    .main-total-card {
        background: linear-gradient(135deg, #1E3A8A, #3B82F6);
        color: white;
        padding: 8px 10px;
        border-radius: 8px;
        text-align: center;
        margin-bottom: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .main-stat-card {
        background-color: #FFFFFF;
        border-left: 4px solid #3B82F6;
        border-radius: 5px;
        padding: 6px 10px;
        margin-bottom: 5px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }
    .border-hanno { border-left-color: #EF4444 !important; }   /* 빨강 */
    .border-minno { border-left-color: #3B82F6 !important; }   /* 파랑 */
    .border-seomoo { border-left-color: #10B981 !important; }  /* 초록 */
    .border-geunsan { border-left-color: #F59E0B !important; } /* 주황 */
    .border-jikwon { border-left-color: #8B5CF6 !important; }  /* 보라 */
    .border-mijeong { border-left-color: #6B7280 !important; } /* 회색 */
    </style>
""",
    unsafe_allow_html=True,
)

EXCEL_FILE = "채용비율.xlsx"


# ==========================================
# 2. 데이터 로드 및 전처리 함수
# ==========================================
@st.cache_data(ttl=5)
def load_all_sheets():
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        sheets = {sheet: pd.read_excel(xls, sheet) for sheet in xls.sheet_names}
        return sheets
    except Exception as e:
        st.error(
            f"엑셀 파일을 읽는 중 오류가 발생했습니다. '{EXCEL_FILE}' 파일이 같은 폴더에 있는지 확인해주세요.\n오류 내용: {e}"
        )
        return None


sheets = load_all_sheets()

if sheets is None:
    st.stop()

# 시트 이름 매핑
sheet_names = list(sheets.keys())
s1_name = sheet_names[0] if len(sheet_names) > 0 else "1.26년 전체 소속별 점유율"
s2_name = sheet_names[1] if len(sheet_names) > 1 else "2.각8개지부 현장점유율"
s3_name = sheet_names[2] if len(sheet_names) > 2 else "3.타워사별 점유현황"
s4_name = sheet_names[3] if len(sheet_names) > 3 else "4.반도체현장"
s5_name = sheet_names[4] if len(sheet_names) > 4 else "5.1~9월채용추이"


# 퍼센트 포맷 변환 헬퍼 함수
def fmt_pct(val):
    try:
        if pd.isna(val) or val == "" or val == "-":
            return "0%"
        f = float(val)
        if f <= 1.0:
            return f"{f * 100:.1f}%"
        return f"{f:.1f}%"
    except:
        return str(val) if pd.notna(val) else "0%"


# ==========================================
# 3. 사이드바 메뉴
# ==========================================
st.sidebar.title("📌 채용 비율 메뉴 선택")
menu = st.sidebar.radio(
    "원하시는 화면을 선택하세요",
    [
        "🏠 메인: 전체 소속별 점유율",
        "🏢 각 8개지부 현장 점유율",
        "🏗️ 타워사별 점유 현황",
        "🏭 반도체 현장 현황",
        "📅 1~9월 채용 추이",
    ],
)

# ==========================================
# 4. 메뉴별 화면 구현
# ==========================================

# ------------------------------------------
# [메뉴 1] 메인: 1.26년 전체 소속별 점유율 (안전 파싱 적용)
# ------------------------------------------
if menu == "🏠 메인: 전체 소속별 점유율":
    st.title("📊 2026년 전체 소속별 점유율")
    df1 = sheets[s1_name].copy()

    try:
        # 안전한 헤더 매핑 방식 적용 (오류 방지)
        header_row = df1.iloc[0].values
        cnt_row = df1.iloc[1].values
        pct_row = df1.iloc[2].values

        data_dict = {}
        for h, c, p in zip(header_row, cnt_row, pct_row):
            h_str = str(h).strip() if pd.notna(h) else ""
            data_dict[h_str] = {"cnt": c, "pct": p}

        tot_cnt = data_dict.get("합계", {}).get("cnt", "209")
        tot_pct = data_dict.get("비율", {}).get("cnt", "100%")

        stat_items = [
            {
                "title": "한국노총 (한노)",
                "cnt": data_dict.get("한노", {}).get("cnt", 0),
                "pct": fmt_pct(data_dict.get("한노", {}).get("pct", 0)),
                "class": "border-hanno",
                "color": "#EF4444",
            },
            {
                "title": "민주노총 (민노)",
                "cnt": data_dict.get("민노", {}).get("cnt", 0),
                "pct": fmt_pct(data_dict.get("민노", {}).get("pct", 0)),
                "class": "border-minno",
                "color": "#3B82F6",
            },
            {
                "title": "섬유노조",
                "cnt": data_dict.get("섬유", {}).get("cnt", 0),
                "pct": fmt_pct(data_dict.get("섬유", {}).get("pct", 0)),
                "class": "border-seomoo",
                "color": "#10B981",
            },
            {
                "title":
