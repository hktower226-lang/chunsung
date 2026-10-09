import pandas as pd
import streamlit as st

# ==========================================
# 1. 페이지 기본 설정 및 모바일 CSS 스타일링
# ==========================================
st.set_page_config(
    page_title="현장 점유율 및 채용 현황", layout="wide", initial_sidebar_state="collapsed"
)

# 모바일용 가독성 증대 CSS (큰 글씨, 위아래 스크롤 레이아웃, 카드 색상 추가)
st.markdown(
    """
    <style>
    /* 전체 글꼴 및 기본 폰트 크기 확대 */
    html, body, [class*="css"] {
        font-size: 19px !important;
    }
    
    /* 제목 및 헤더 크기 확대 */
    h1 { font-size: 2.2rem !important; font-weight: bold !important; color: #1E3A8A; }
    h2 { font-size: 1.8rem !important; font-weight: bold !important; color: #1E40AF; }
    h3 { font-size: 1.5rem !important; font-weight: bold !important; }
    
    /* 카드/Expander 내부 텍스트 확대 */
    .stSelectbox label, .stRadio label, .stMultiSelect label {
        font-size: 1.2rem !important;
        font-weight: bold !important;
    }
    
    /* 모바일 반응형 좌우 여백 축소 (위아래 세로 스크롤 최적화) */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 3rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
        max-width: 100% !important;
    }
    
    /* 테이블 글씨 크기 확대 */
    .dataframe {
        font-size: 16px !important;
    }
    
    /* 전화번호 연결 버튼 스타일 */
    .phone-btn {
        display: inline-block;
        background-color: #25D366;
        color: white !important;
        padding: 10px 18px;
        font-size: 18px;
        font-weight: bold;
        text-decoration: none;
        border-radius: 8px;
        margin-top: 5px;
        margin-bottom: 10px;
        text-align: center;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }

    /* 메인 화면 전용 카드 스타일 & 소속별 구분색 */
    .main-total-card {
        background: linear-gradient(135deg, #
