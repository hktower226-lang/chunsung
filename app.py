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

# 2. 모바일 친화적 디자인 및 강조 컬러 스타일 적용
st.markdown(
    """
    <style>
    .main { background-color: #f8fafc; }
    h1 { font-size: 1.25rem !important; color: #0f172a; text-align: center; margin-bottom: 0.1rem; font-weight: 800; }
    .subtitle { font-size: 0.7rem; color: #64748b; text-align: center; margin-bottom: 0.8rem; }
    
    /* 상단 메인 통계 카드 스타일 */
    .summary-card {
        background: white;
        padding: 12px;
        border-radius: 10px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
        margin-bottom: 12px;
        font-size: 0.75rem;
        text-align: center;
    }
    .summary-title { font-weight: 700; color: #1e293b; margin-bottom: 6px; font-size: 0.8rem; }
    
    /* 소속별 강조 컬러 뱃지 */
    .badge-hanno { color: #dc2626; font-weight: 700; background: #fee2e2; padding: 1px 4px; border-radius: 4px; }
    .badge-minno { color: #2563eb; font-weight: 700; background: #dbeafe; padding: 1px 4px; border-radius: 4px; }
    .badge-sumyu { color: #d97706; font-weight: 700; background: #fef3c7; padding: 1px 4px; border-radius: 4px; }
    .badge-geonsan { color: #059669; font-weight: 700; background: #d1fae5; padding: 1px 4px; border-radius: 4px; }
    .badge-staff { color: #7c3aed; font-weight: 700; background: #ede9fe; padding: 1px 4px; border-radius: 4px; }
    .badge-etc { color: #475569; font-weight: 700; background: #f1f5f9; padding: 1px 4px; border-radius: 4px; }

    div.stDataFrame { border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.05); font-size: 0.75rem; }
    
    .stTabs [data-baseweb="tab-list"] { gap: 6px; justify-content: center; }
    .stTabs [data-baseweb="tab"] { 
        height: 38px; background-color: #e2e8f0; border-radius: 6px; 
        color: #334155; font-weight: 600; font-size: 0.8rem; padding: 0 14px;
    }
    .stTabs [aria-selected="true"] { background-color: #2563eb !important; color: white !important; }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown("🏗️ 경기지역본부 현장 점유율 및 통계", unsafe_allow_html=True)
st.markdown('<p class="subtitle">모바일 최적화 통합 조회 시스템</p>', unsafe_allow_html=True)


# 3. 엑셀 데이터 가공 및 통합 함수
@st.cache_data(ttl=5)
def load_and_merge_data():
    excel_files = [f for f in os.listdir(".") if f.endswith((".xlsx", ".xls")) and not f.startswith("~$")]
    if not excel_files:
        return None, None, "폴더에 엑셀 파일이 없습니다."

    target_file = excel_files[0]
    try:
        xl = pd.ExcelFile(target_file)
        
        # --- 전체현황 데이터 읽기 (상단 메인 요약용) ---
        df1_raw = pd.read_excel(target_file, sheet_name=1)
        # 전체 소속별 점유율 추출 (행 1: 전체 대수, 행 2: 비율)
        total_counts = df1_raw.iloc[1]
        total_ratios = df1_raw.iloc[2]
        
        # 지부별 소속 현황 테이블 추출 (행 6부터)
        header_row = df1_raw.iloc[5].astype(str).str.strip()
        df_branch_stats = df1_raw.iloc[6:].copy()
        df_branch_stats.columns = header_row
        df_branch_stats = df_branch_stats.loc[:, df_branch_stats.columns.notna() & (df_branch_stats.columns != 'nan')]
        df_branch_stats = df_branch_stats.fillna("-")

        # --- 지부별 현장명단 데이터 읽기 ---
        df0 = pd.read_excel(target_file, sheet_name=0)
        df0.columns = df0.iloc[0].astype(str).str.strip()
        df0 = df0.iloc[1:].reset_index(drop=True)
        df0 = df0.loc[:, ~df0.columns.str.contains('^Unnamed')]
        df0 = df0.fillna("-")

        # 지부별 소계 행에 지부별 소속 현황(대수 + 진한 퍼센트) 반영하기
        unions = ['한노', '민노', '섬유', '건산', '직원', '기타']
        
        # 지부 이름 매핑 정리 (예: "북부지부 소계" -> "북부")
        for idx, row in df0.iterrows():
            jibu_name = str(row['지부']).strip()
            if '소계' in jibu_name:
                clean_jibu = jibu_name.replace(' 소계', '').replace('지부', '') + '지부'
                # 해당 지부의 통계 찾기
                match_stat = df_branch_stats[df_branch_stats['지부'].astype(str).str.contains(clean_jibu, na=False)]
                if not match_stat.empty:
                    for union in unions:
                        col_cnt = f"{union} (대)"
                        col_pct = f"{union} (%)"
                        if col_cnt in match_stat.columns and col_pct in match_stat.columns:
                            c_val = match_stat[col_cnt].values[0]
                            p_val = match_stat[col_pct].values[0]
                            if pd.notna(c_val) and c_val != "-" and pd.notna(p_val) and p_val != "-":
                                try:
                                    p_num = float(p_val) * 100
                                    # 진하게 강조된 HTML 스타일 적용
                                    df0.loc[idx, union] = f"<b>{c_val}대</b> <span style='color:#1d4ed8; font-weight:900;'>({p_num:.1f}%)</span>"
                                except:
                                    pass

        # --- 타워사별 현황 데이터 읽기 ---
        df2_raw = pd.read_excel(target_file, sheet_name=2)
        df2 = df2_raw.iloc[1:].copy() if len(df2_raw) > 1 else df2_raw.copy()
        df2.columns = df2_raw.iloc[0].astype(str).str.strip()
        df2 = df2.loc[:, ~df2.columns.str.contains('^Unnamed')]
        df2 = df2.fillna("-")

        # 한노점유율을 맨 첫 번째로 배치
        if '한노점유율' in df2.columns and '타워사' in df2.columns:
            cols = ['타워사', '한노점유율'] + [c for c in df2.columns if c not in ['타워사', '한노점유율']]
            df2 = df2[cols]
            def fmt_hano_ratio(val):
                try:
                    v = float(val)
                    return f"<span style='color:#dc2626; font-weight:900;'>{v * 100:.1f}%</span>"
                except:
                    return str(val)
            df2['한노점유율'] = df2['한노점유율'].apply(fmt_hano_ratio)

        # 요약 데이터 딕셔너리 생성
        summary_info = {
            "한노": f"{total_counts.get('한노', 51)}대 ({float(total_ratios.get('한노', 0.2786))*100:.1f}%)",
            "민노": f"{total_counts.get('민노', 79)}대 ({float(total_ratios.get('민노', 0.4316))*100:.1f}%)",
            "섬유": f"{total_counts.get('섬유', 6)}대 ({float(total_ratios.get('섬유', 0.0327))*100:.1f}%)",
            "건산": f"{total_counts.get('건산', 9)}대 ({float(total_ratios.get('건산', 0.0491))*100:.1f}%)",
            "직원": f"{total_counts.get('직원', 32)}대 ({float(total_ratios.get('직원', 0.1748))*100:.1f}%)",
            "기타": f"{total_counts.get('기타', 6)}대 ({float(total_ratios.get('기타', 0.0327))*100:.1f}%)",
            "합계": f"{total_counts.get('합계', 183)}대 (100.0%)"
        }

        return df0, df2, summary_info, None
    except Exception as e:
        return None, None, None, str(e)


df_main, df_sub2, summary, err_msg = load_and_merge_data()

if err_msg or df_main is None:
    st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {err_msg}")
else:
    tabs = st.tabs(["📋 지부별 현장 및 통계", "🏗️ 타워사별 현황"])

    # --- [탭 1] 지부별 현장명단 및 상단 메인 통계 카드 ---
    with tabs[0]:
        # 상단 메인 요약 박스
        if summary:
            st.markdown(
                f"""
                <div class="summary-card">
                    <div class="summary-title">📊 전체 소속별 점유 현황 (총 {summary['합계']})</div>
                    <div style="display: flex; justify-content: space-around; flex-wrap: wrap; gap: 4px; margin-top: 6px;">
                        <div>🔴 한노: <span style="font-weight:700;">{summary['한노']}</span></div>
                        <div>🔵 민노: <span style="font-weight:700;">{summary['민노']}</span></div>
                        <div>🟠 섬유: <span style="font-weight:700;">{summary['섬유']}</span></div>
                        <div>🟢 건산: <span style="font-weight:700;">{summary['건산']}</span></div>
                        <div>🟣 직원: <span style="font-weight:700;">{summary['직원']}</span></div>
                        <div>⚪ 기타: <span style="font-weight:700;">{summary['기타']}</span></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        search_kw0 = st.text_input("🔍 현장 통합 검색", placeholder="현장명, 지부, 타워회사 검색", label_visibility="collapsed")
        filtered_df0 = df_main.copy()
        if search_kw0:
            mask = filtered_df0.astype(str).apply(lambda x: x.str.contains(search_kw0, case=False, na=False)).any(axis=1)
            filtered_df0 = filtered_df0[mask]

        st.markdown(f"<p style='font-size: 0.7rem; color: #64748b; margin: 4px 0;'>검색 결과: <b>{len(filtered_df0)}</b>건</p>", unsafe_allow_html=True)
        
        # HTML 마크업이 포함된 데이터프레임을 안전하게 렌더링하기 위해 st.markdown 사용 또는 기본 dataframe
        st.dataframe(filtered_df0, use_container_width=True, hide_index=True)

    # --- [탭 2] 타워사별 현황 ---
    with tabs[1]:
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
