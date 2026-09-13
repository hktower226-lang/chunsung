import os
import pandas as pd
import streamlit as st

# 1. 페이지 설정
st.set_page_config(
    page_title="경기지역본부 현장 점유율 및 통계 조회",
    page_icon="🏗️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# 2. 모바일 가독성을 위한 CSS 스타일
st.markdown(
    """
    <style>
    .main { background-color: #f3f4f6; }
    h1 { font-size: 1.4rem !important; color: #111827; text-align: center; margin-bottom: 0.2rem; font-weight: 700; }
    .subtitle { font-size: 0.8rem; color: #4b5563; text-align: center; margin-bottom: 1.2rem; }
    div.stDataFrame { border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; justify-content: center; }
    .stTabs [data-baseweb="tab"] { height: 40px; background-color: #e5e7eb; border-radius: 6px; color: #374151; font-weight: 600; font-size: 0.85rem; }
    .stTabs [aria-selected="true"] { background-color: #2563eb !important; color: white !important; }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown("🏗️ 경기지역본부 현장 점유율 및 통계", unsafe_allow_html=True)
st.markdown('<p class="subtitle">모바일 최적화 통합 조회 시스템 (조회 전용)</p>', unsafe_allow_html=True)


# [순정 로직] 코드가 숫자를 절대 건드리지 않고 엑셀 원본 그대로 표시 + 빈칸만 '-'로 처리
@st.cache_data(ttl=5)
def load_excel_raw():
    excel_files = [f for f in os.listdir(".") if f.endswith((".xlsx", ".xls")) and not f.startswith("~$")]
    if not excel_files:
        return None, None, None, "폴더에 엑셀 파일이 없습니다."

    target_file = excel_files[0]
    try:
        xl = pd.ExcelFile(target_file)
        sheet_names = xl.sheet_names

        # 각 시트의 컬럼 정리 및 빈칸만 '-'로 치환 (숫자 데이터 변조 0%)
        df1 = pd.read_excel(target_file, sheet_name=0)
        df1.columns = df1.columns.astype(str).str.strip()
        df1 = df1.fillna("-")

        df2 = pd.read_excel(target_file, sheet_name=1) if len(sheet_names) > 1 else None
        if df2 is not None:
            df2.columns = df2.columns.astype(str).str.strip()
            df2 = df2.fillna("-")

        df3 = pd.read_excel(target_file, sheet_name=2) if len(sheet_names) > 2 else None
        if df3 is not None:
            df3.columns = df3.columns.astype(str).str.strip()
            df3 = df3.fillna("-")

        return df1, df2, df3, None
    except Exception as e:
        return None, None, None, str(e)


df_main, df_sub1, df_sub2, err_msg = load_excel_raw()

if err_msg or df_main is None:
    st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {err_msg}")
else:
    tab_names = ["📋 전체 현장 조회", "📊 지부별 점유율"]
    if df_sub2 is not None:
        tab_names.append("🏗️ 타워사별 현황")

    tabs = st.tabs(tab_names)

    # --- [탭 1] 전체 현장 목록 ---
    with tabs[0]:
        st.markdown("##### 🔍 현장 통합 검색")
        search_keyword = st.text_input("검색어 입력", placeholder="현장명, 지부, 타워사 등을 입력하세요", label_visibility="collapsed")
        
        filtered_df = df_main.copy()
        if search_keyword:
            mask = filtered_df.astype(str).apply(lambda x: x.str.contains(search_keyword, case=False, na=False)).any(axis=1)
            filtered_df = filtered_df[mask]

        st.markdown(f"<p style='font-size: 0.75rem; color: #6b7280; margin-bottom: 5px;'>조회 결과: <b>{len(filtered_df)}</b>건</p>", unsafe_allow_html=True)
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)

    # --- [탭 2] 지부별 점유율 ---
    with tabs[1]:
        st.markdown("##### 📊 지부별 점유율 분석")
        if df_sub1 is not None and not df_sub1.empty:
            st.dataframe(df_sub1, use_container_width=True, hide_index=True)
        else:
            st.info("데이터가 없습니다.")

    # --- [탭 3] 타워사별 현황 ---
    if len(tabs) > 2 and df_sub2 is not None:
        with tabs[2]:
            st.markdown("##### 🏗️ 타워사별 상세 현황")
            if not df_sub2.empty:
                st.dataframe(df_sub2, use_container_width=True, hide_index=True)
            else:
                st.info("데이터가 없습니다.")

    st.markdown("---")
    st.markdown("<p style='text-align: center; font-size: 0.7rem; color: #9ca3af;'>Gyeonggi Regional Headquarters Dashboard</p>", unsafe_allow_html=True)
