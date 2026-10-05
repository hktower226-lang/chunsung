import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="2026 채용 현황 대시보드",
    page_icon="📱",
    layout="centered", # 모바일 레이아웃 최적화
    initial_sidebar_state="collapsed"
)

# 모바일 UI 커스텀 스타일링
st.markdown("""
<style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        padding-left: 0.8rem;
        padding-right: 0.8rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 18px;
        border-radius: 12px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
        margin-bottom: 15px;
    }
    .metric-title {
        font-size: 1.1rem;
        font-weight: bold;
        margin-bottom: 8px;
        color: #e0eafc;
    }
    .metric-main {
        font-size: 1.8rem;
        font-weight: 800;
        color: #ffffff;
    }
    .branch-card {
        background-color: #f8f9fa;
        border-left: 5px solid #2a5298;
        padding: 12px;
        border-radius: 8px;
        margin-bottom: 10px;
    }
</style>
""", unsafe_content_policy=True)

# 깃허브 원본 엑셀 RAW URL
GITHUB_EXCEL_URL = "https://raw.githubusercontent.com/chunsung/chunsung/main/%EC%B2%84%EC%9A%A9%EB%B9%84%EC%9C%A8.xlsx"

@st.cache_data(ttl=60) # 1분 간격 자동 갱신
def load_excel_data():
    try:
        xl = pd.ExcelFile(GITHUB_EXCEL_URL)
        df_total = pd.read_excel(xl, sheet_name='채용비율')
        return df_total
    except Exception as e:
        st.error(f"GitHub에서 엑셀 파일을 불러오는 중 오류가 발생했습니다: {e}")
        return None

df_raw = load_excel_data()

if df_raw is not None:
    st.title("📱 2026 채용 현황")
    st.caption("🔄 GitHub 연동 완료 (자동 업데이트)")

    # ----------------------------------------------------
    # 1. 메인 고정: 2026년 전체 소속별 점유율
    # ----------------------------------------------------
    st.markdown("### 📊 2026년 전체 소속별 점유율")
    
    try:
        cnt_vals =
