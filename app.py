import os
import pandas as pd
import streamlit as st

# 1. 페이지 설정 (모바일 최적화)
st.set_page_config(
    page_title="경기지역본부 현장 점유율 및 통계 조회",
    page_icon="🏗️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# 2. 모바일 친화적 세련된 디자인 CSS 스타일 적용
st.markdown(
    """
    <style>
    .main { background-color: #f8fafc; }
    h1 { font-size: 1.35rem !important; color: #0f172a; text-align: center; margin-bottom: 0.1rem; font-weight: 800; }
    .subtitle { font-size: 0.75rem; color: #64748b; text-align: center; margin-bottom: 1rem; }
    div.stDataFrame { border-radius: 10px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .stTabs [data-baseweb="tab-list"] { gap: 6px; justify-content: center; }
    .stTabs [data-baseweb="tab"] { 
        height: 38px; background-color: #e2e8f0; border-radius: 8px; 
        color: #334155; font-weight: 600; font-size: 0.8rem; padding: 0 12px;
    }
    .stTabs [aria-selected="true"] { background-color: #2563eb !important; color: white !important; }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown("🏗️ 경기지역본부 현장 점유율 및 통계", unsafe_allow_html=True)
st.markdown('<p class="subtitle">모바일 최적화 통합 조회 시스템 (조회 전용)</p>', unsafe_allow_html=True)


# 3. 엑셀 파일 로드 및 퍼센트 자동 변환 전처리 함수
@st.cache_data(ttl=5)
def load_and_preprocess_excel():
    excel_files = [f for f in os.listdir(".") if f.endswith((".xlsx", ".xls")) and not f.startswith("~$")]
    if not excel_files:
        return None, None, None, "폴더에 엑셀 파일이 없습니다."

    target_file = excel_files[0]
    try:
        xl = pd.ExcelFile(target_file)
        sheet_names = xl.sheet_names

        dfs = []
        for i, name in enumerate(sheet_names):
            df = pd.read_excel(target_file, sheet_name=i)
            df.columns = df.columns.astype(str).str.strip()
            df = df.fillna("-")

            # 소수점 비율 데이터(0과 1 사이의 값)를 찾아 퍼센트 문자열로 자동 변환 (모바일 가독성 극대화)
            for col in df.columns:
                try:
                    converted = pd.to_numeric(df[col], errors='coerce')
                    mask = (converted > 0) & (converted < 1)
                    if mask.any():
                        df[col] = df[col].apply(
                            lambda x: f"{float(x)*100:.1f}%" if pd.notna(x) and isinstance(x, (int, float)) and 0 < float(x) < 1 else x
                        )
                except Exception:
                    pass
            dfs.append(df)

        df1 = dfs[0] if len(dfs) > 0 else None
        df2 = dfs[1] if len(dfs) > 1 else None
        df3 = dfs[2] if len(dfs) > 2 else None

        return df1, df2, df3, None
    except Exception as e:
        return None, None, None, str(e)


df_main, df_sub1, df_sub2, err_msg = load_and_preprocess_excel()

if err_msg or df_main is None:
    st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {err_msg}")
else:
    # 동적으로 시트 탭 구성
    tab_names = ["📋 지부별 현장명단", "📊 전체 현황"]
    if df_sub2 is not None:
        tab_names.append("🏗️ 타워사별 현황")

    tabs = st.tabs(tab_names)

    # --- [탭 1] 지부별 현장명단 및 통합 검색 ---
    with tabs[0]:
        st.markdown("##### 🔍 현장 통합 검색")
        search_keyword = st.text_input("검색어 입력", placeholder="현장명, 지부, 타워회사, 특이사항 입력", label_visibility="collapsed")
        
        filtered_df = df_main.copy()
        if search_keyword:
            mask = filtered_df.astype(str).apply(lambda x: x.str.contains(search_keyword, case=False, na=False)).any(axis=1)
            filtered_df = filtered_df[mask]

        st.markdown(f"<p style='font-size: 0.75rem; color: #64748b; margin-bottom: 6px;'>조회 결과: <b>{len(filtered_df)}</b>건</p>", unsafe_allow_html=True)
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)

    # --- [탭 2] 전체 현황 (점유율 분석 포함) ---
    with tabs[1]:
        st.markdown("##### 📊 전체 소속별 및 지부별 점유율 현황")
        if df_sub1 is not None and not df_sub1.empty:
            st.dataframe(df_sub1, use_container_width=True, hide_index=True)
        else:
            st.info("데이터가 없습니다.")

    # --- [탭 3] 타워사별 현황 ---
    if len(tabs) > 2 and df_sub2 is not None:
        with tabs[2]:
            st.markdown("##### 🏗️ 타워(렌탈)사별 상세 점유 현황")
            if not df_sub2.empty:
                st.dataframe(df_sub2, use_container_width=True, hide_index=True)
            else:
                st.info("데이터가 없습니다.")

    st.markdown("---")
    st.markdown("<p style='text-align: center; font-size: 0.7rem; color: #94a3b8;'>Gyeonggi Regional Headquarters Dashboard</p>", unsafe_allow_html=True)
