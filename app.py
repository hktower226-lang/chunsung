import pandas as pd
import streamlit as st

# 페이지 설정 (모바일 최적화 및 넓은 화면 레이아웃)
st.set_page_config(
    page_title="점유율 현황 모바일 대시보드",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# 모바일 가독성 및 스타일 CSS 적용
st.markdown(
    """
    <style>
    .main { font-size: 18px !important; }
    h1 { font-size: 24px !important; font-weight: bold; }
    h2 { font-size: 20px !important; font-weight: bold; }
    p, label, .stMarkdown { font-size: 16px !important; }
    
    .metric-card {
        background-color: #ffffff;
        padding: 12px;
        border-radius: 10px;
        margin-bottom: 10px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        word-break: break-all;
    }
    .subtotal-card {
        background-color: #ebf8ff;
        padding: 14px;
        border-radius: 10px;
        margin-bottom: 12px;
        border: 2px solid #3182ce;
        word-break: break-all;
    }
    
    .badge-hanno { background-color: #e53e3e; color: white; padding: 4px 8px; border-radius: 6px; font-weight: bold; display: inline-block; margin: 2px; font-size: 13px; }
    .badge-minno { background-color: #3182ce; color: white; padding: 4px 8px; border-radius: 6px; font-weight: bold; display: inline-block; margin: 2px; font-size: 13px; }
    .badge-seomvu { background-color: #dd6b20; color: white; padding: 4px 8px; border-radius: 6px; font-weight: bold; display: inline-block; margin: 2px; font-size: 13px; }
    .badge-geonsan { background-color: #38a169; color: white; padding: 4px 8px; border-radius: 6px; font-weight: bold; display: inline-block; margin: 2px; font-size: 13px; }
    .badge-staff { background-color: #805ad5; color: white; padding: 4px 8px; border-radius: 6px; font-weight: bold; display: inline-block; margin: 2px; font-size: 13px; }
    .badge-etc { background-color: #718096; color: white; padding: 4px 8px; border-radius: 6px; font-weight: bold; display: inline-block; margin: 2px; font-size: 13px; }
    
    .item-container {
        display: flex;
        flex-wrap: wrap;
        gap: 4px;
        margin-top: 6px;
        align-items: center;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

file_path = "채용비율.xlsx"

@st.cache_data
def load_data():
    xls = pd.ExcelFile(file_path)
    return xls.sheet_names

try:
    sheet_names = load_data()
except Exception as e:
    st.error(f"엑셀 파일('채용비율.xlsx')을 불러오지 못했습니다. 경로를 확인해주세요. 에러: {e}")
    st.stop()

# 상단 타이틀
st.title("📊 통합 점유율 현황 모바일 대시보드")
st.write("---")

# 3가지 탭 구성
tab1, tab2, tab3 = st.tabs([
    "1. 26년 채용비율", 
    "2. 반도체현장", 
    "3. 24년9월~25년12월 채용비율"
])

# -------------------------------------------------------------------------
# [탭 1] 26년 채용비율 시트
# -------------------------------------------------------------------------
with tab1:
    st.header("🏢 2026년 전체 소속별 점유율")
    st.markdown("맨 위 고정된 전체 소속 현황 및 지부별·임대사별 상세 조회입니다.")

    try:
        df_ratio = pd.read_excel(file_path, sheet_name="채용비율")
        
        t1 = df_ratio.iloc[0:2, :8].copy()
        t1.columns = t1.iloc[0]
        t1 = t1.drop(0).reset_index(drop=True)
        cols = [c for c in t1.columns if c not in ["소속", "합계", "비율"]]
        total_val = float(t1["합계"].values[0]) if "합계" in t1.columns else 198

        badge_map = {
            "한노": "badge-hanno", "민노": "badge-minno", "섬유": "badge-seomvu",
            "건산": "badge-geonsan", "직원": "badge-staff", "미정": "badge-etc", "기타": "badge-etc"
        }

        st.markdown("### 📌 2026년 전체 소속별 요약")
        for col in cols:
            val = float(t1[col].values[0])
            pct = (val / total_val) * 100
            b_class = badge_map.get(col, "badge-etc")
            is_hanno = (col == "한노")
            
            card_style = "border: 2px solid #e53e3e; background-color: #fff5f5;" if is_hanno else ""
            color_code = "#e53e3e" if is_hanno else "#2b6cb0"
            
            card_html = (
                '<div class="metric-card" style="' + card_style + ' display: flex; justify-content: space-between; align-items: center;">'
                '<div><span class="' + b_class + '">' + col + '</span></div>'
                '<div>'
                '<span style="font-size: 17px; font-weight: bold;">' + f"{val:,.0f}" + '명</span> &nbsp;|&nbsp; '
                '<span style="font-size: 17px; color: ' + color_code + '; font-weight: bold;">' + f"{pct:.1f}" + '%</span>'
                '</div>'
                '</div>'
            )
            st.markdown(card_html, unsafe_allow_html=True)

        st.write("---")
        st.subheader("🔍 경기지역본부 현장 및 지부 검색")

        t2 = df_ratio.iloc[7:76, :12].copy()
        t2.columns = ["지부", "현장명", "타워회사", "특이사항", "한노", "민노", "섬유", "건산", "직원", "미정", "총대수", "비고"]
        t2["지부"] = t2["지부"].ffill()

        for num_col in ["한노", "민노", "섬유", "건산", "직원", "미정", "총대수"]:
            t2[num_col] = pd.to_numeric(t2[num_col], errors="coerce").fillna(0)

        raw_branches = t2["지부"].dropna().astype(str).unique().tolist()
        clean_branches = []
        for b in raw_branches:
            cleaned = b.replace("지부 소계", "").replace("지부", "").replace(" 소계", "").strip()
            if cleaned and cleaned not in clean_branches:
                clean_branches.append(cleaned)

        selected_branch = st.selectbox("지부를 선택하세요", ["전체 보기"] + clean_branches, key="tab1_branch")

        if selected_branch != "전체 보기":
            filtered_t2 = t2[t2["지부"].astype(str).str.contains(selected_branch)]
        else:
            filtered_t2 = t2

        for _, row in filtered_t2.iterrows():
            is_
