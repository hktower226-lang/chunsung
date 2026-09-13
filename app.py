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

# 2. 모바일 친화적 세련된 디자인 및 칸 너비 압축 CSS 스타일 적용
st.markdown(
    """
    <style>
    .main { background-color: #f8fafc; }
    h1 { font-size: 1.25rem !important; color: #0f172a; text-align: center; margin-bottom: 0.1rem; font-weight: 800; }
    .subtitle { font-size: 0.7rem; color: #64748b; text-align: center; margin-bottom: 0.8rem; }
    div.stDataFrame { border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.05); font-size: 0.75rem; }
    
    /* 지부별 현장명단 첫 번째 컬럼(지부) 폭 좁게 압축 */
    table td:nth-child(1), table th:nth-child(1) { max-width: 50px !important; width: 50px !important; text-align: center !important; }
    
    .stTabs [data-baseweb="tab-list"] { gap: 4px; justify-content: center; }
    .stTabs [data-baseweb="tab"] { 
        height: 36px; background-color: #e2e8f0; border-radius: 6px; 
        color: #334155; font-weight: 600; font-size: 0.75rem; padding: 0 10px;
    }
    .stTabs [aria-selected="true"] { background-color: #2563eb !important; color: white !important; }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown("🏗️ 경기지역본부 현장 점유율 및 통계", unsafe_allow_html=True)
st.markdown('<p class="subtitle">모바일 최적화 통합 조회 시스템</p>', unsafe_allow_html=True)


# 3. 엑셀 데이터 가공 함수
@st.cache_data(ttl=5)
def load_and_clean_data():
    excel_files = [f for f in os.listdir(".") if f.endswith((".xlsx", ".xls")) and not f.startswith("~$")]
    if not excel_files:
        return None, None, None, "폴더에 엑셀 파일이 없습니다."

    target_file = excel_files[0]
    try:
        xl = pd.ExcelFile(target_file)
        
        # --- 시트 1: 지부별 현장명단 ---
        df0 = pd.read_excel(target_file, sheet_name=0)
        # 첫 번째 행을 컬럼명으로 지정
        df0.columns = df0.iloc[0].astype(str).str.strip()
        df0 = df0.iloc[1:].reset_index(drop=True)
        # Unnamed 컬럼 제거 및 결측치 처리
        df0 = df0.loc[:, ~df0.columns.str.contains('^Unnamed')]
        df0 = df0.fillna("-")

        # --- 시트 2: 전체현황 (소속별/지부별 현황 재구성) ---
        df1_raw = pd.read_excel(target_file, sheet_name=1)
        # 지부별 소속 현황 테이블 추출 (행 5부터 끝까지)
        header_row = df1_raw.iloc[5].astype(str).str.strip()
        df1 = df1_raw.iloc[6:].copy()
        df1.columns = header_row
        df1 = df1.loc[:, df1.columns.notna() & (df1.columns != 'nan')]
        df1 = df1.fillna("-")

        # 대수와 퍼센티지를 합쳐서 "대수 (퍼센트)" 형태로 변환
        unions = ['한노', '민노', '섬유', '건산', '직원', '기타']
        for union in unions:
            col_cnt = f"{union} (대)"
            col_pct = f"{union} (%)"
            if col_cnt in df1.columns and col_pct in df1.columns:
                def combine_cnt_pct(row):
                    c_val = row[col_cnt]
                    p_val = row[col_pct]
                    if pd.notna(c_val) and c_val != "-" and pd.notna(p_val) and p_val != "-":
                        try:
                            p_num = float(p_val) * 100
                            return f"{c_val} ({p_num:.1f}%)"
                        except:
                            return str(c_val)
                    return str(c_val)
                df1[union] = df1.apply(combine_cnt_pct, axis=1)
        
        # 보기 깔끔하게 컬럼 정리 (합계, 현장수 유지)
        keep_cols = ['지부'] + [u for u in unions if u in df1.columns] + ['합계', '현장수']
        df1 = df1[[c for c in keep_cols if c in df1.columns]]

        # --- 시트 3: 타워사별 현황 ---
        df2_raw = pd.read_excel(target_file, sheet_name=2)
        df2 = df2_raw.iloc[1:].copy() if len(df2_raw) > 1 else df2_raw.copy()
        df2.columns = df2_raw.iloc[0].astype(str).str.strip()
        df2 = df2.loc[:, ~df2.columns.str.contains('^Unnamed')]
        df2 = df2.fillna("-")

        # 한노점유율을 맨 첫 번째로 배치 (타워사 이름 다음으로)
        if '한노점유율' in df2.columns and '타워사' in df2.columns:
            cols = ['타워사', '한노점유율'] + [c for c in df2.columns if c not in ['타워사', '한노점유율']]
            df2 = df2[cols]
            # 한노점유율 소수점을 백분율(%)로 변환
            def fmt_hano_ratio(val):
                try:
                    v = float(val)
                    return f"{v * 100:.1f}%"
                except:
                    return str(val)
            df2['한노점유율'] = df2['한노점유율'].apply(fmt_hano_ratio)

        # 타워사별 데이터 정렬 (기본 순서 유지 또는 가나다순)
        df2 = df2.reset_index(drop=True)

        return df0, df1, df2, None
    except Exception as e:
        return None, None, None, str(e)


df_main, df_sub1, df_sub2, err_msg = load_and_clean_data()

if err_msg or df_main is None:
    st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {err_msg}")
else:
    tabs = st.tabs(["📋 지부별 현장", "📊 전체 현황", "🏗️ 타워사별 현황"])

    # --- [탭 1] 지부별 현장명단 ---
    with tabs[0]:
        search_kw0 = st.text_input("🔍 현장 통합 검색", placeholder="현장명, 지부, 타워회사 검색", label_visibility="collapsed")
        filtered_df0 = df_main.copy()
        if search_kw0:
            mask = filtered_df0.astype(str).apply(lambda x: x.str.contains(search_kw0, case=False, na=False)).any(axis=1)
            filtered_df0 = filtered_df0[mask]

        st.markdown(f"<p style='font-size: 0.7rem; color: #64748b; margin: 4px 0;'>검색 결과: <b>{len(filtered_df0)}</b>건</p>", unsafe_allow_html=True)
        st.dataframe(filtered_df0, use_container_width=True, hide_index=True)

    # --- [탭 2] 전체 현황 ---
    with tabs[1]:
        st.markdown("<p style='font-size: 0.8rem; font-weight: 700; margin-bottom: 6px;'>📊 지부별 소속 현황 (대수 및 점유율)</p>", unsafe_allow_html=True)
        if df_sub1 is not None and not df_sub1.empty:
            st.dataframe(df_sub1, use_container_width=True, hide_index=True)
        else:
            st.info("데이터가 없습니다.")

    # --- [탭 3] 타워사별 현황 ---
    with tabs[2]:
        search_kw2 = st.text_input("🔍 타워사 검색", placeholder="타워(렌탈)사 이름 입력", label_visibility="collapsed")
        filtered_df2 = df_sub2.copy()
        if search_kw2 and filtered_df2 is not None:
            mask = filtered_df2.astype(str).apply(lambda x: x.str.contains(search_kw2, case=False, na=False)).any(axis=1)
            filtered_df2 = filtered_df2[mask]

        st.markdown(f"<p style='font-size: 0.7rem; color: #64748b; margin: 4px 0;'>검색 결과: <b>{len(filtered_df2) if filtered_df2 is not None else 0}</b>건</p>", unsafe_allow_html=True)
        if filtered_df2 is not None and not filtered_df2.empty:
            st.dataframe(filtered_df2, use_container_width=True, hide_index=True)
        else:
            st.info("데이터가 없습니다.")

    st.markdown("---")
    st.markdown("<p style='text-align: center; font-size: 0.65rem; color: #94a3b8;'>Gyeonggi Regional Headquarters Dashboard</p>", unsafe_allow_html=True)
