import pandas as pd
import streamlit as st

# ==========================================
# 1. 페이지 기본 설정 및 모바일 CSS 스타일링
# ==========================================
st.set_page_config(
    page_title="현장 점유율 및 채용 현황", layout="wide", initial_sidebar_state="collapsed"
)

# 모바일용 가독성 증대 CSS (큰 글씨, 위아래 스크롤 레이아웃, 버튼 강조)
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

# ==========================================
# 3. 사이드바 메뉴 (모바일 상단/좌측 메뉴)
