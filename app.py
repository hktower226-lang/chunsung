import pandas as pd
import streamlit as st

# ==========================================
# 1. 페이지 기본 설정 및 모바일 CSS 스타일링
# ==========================================
st.set_page_config(
    page_title="현장 점유율 및 채용 현황", layout="wide", initial_sidebar_state="collapsed"
)

# 모바일 한 화면(Single Screen) 최적화 CSS (글자 크기, 여백, 패딩 대폭 축소)
st.markdown(
    """
    <style>
    /* 전체 글꼴 및 기본 폰트 크기 슬림화 */
    html, body, [class*="css"] {
        font-size: 15px !important;
    }
    
    /* 제목 및 헤더 크기 및 상하 여백 대폭 축소 */
    h1 { font-size: 1.4rem !important; font-weight: bold !important; color: #1E3A8A; margin-bottom: 5px !important; margin-top: 0px !important; }
    h2 { font-size: 1.2rem !important; font-weight: bold !important; color: #1E40AF; margin-bottom: 5px !important; }
    h3 { font-size: 1.0rem !important; font-weight: bold !important; margin-top: 5px !important; margin-bottom: 5px !important; }
    
    /* 선택 박스 및 라벨 여백 최소화 */
    .stSelectbox label, .stRadio label, .stMultiSelect label {
        font-size: 1.0rem !important;
        font-weight: bold !important;
    }
    
    /* 모바일 여백 최소화 (한 화면에 다 들어가도록 패딩 최소화) */
    .block-container {
        padding-top: 0.5rem !important;
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
    }

    /* 메인 화면 한 화면(Single View) 맞춤 슬림 카드 디자인 */
    .main-total-card {
        background: linear-gradient(135deg, #1E3A8A, #3B82F6);
        color: white;
        padding: 6px 10px;
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
@st.cache_data(ttl=5)  # 엑셀 수정 시 빠르게 반영되도록 설정
def load_all_sheets():
    try:
        xls = pd.Excel
