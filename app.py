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

# 2. 모바일 친화적 디자인 및 가로 스크롤 테이블 스타일 적용
st.markdown(
    """
    <style>
    .main { background-color: #f8fafc; }
    h1 { font-size: 1.2rem !important; color: #0f172a; text-align: center; margin-bottom: 0.1rem; font-weight: 800; }
    .subtitle { font-size: 0.7rem; color: #64748b; text-align: center; margin-bottom: 0.6rem; }
    
    /* 상단 메인 통계 카드 스타일 */
    .summary-card {
        background: white;
        padding: 10px;
        border-radius: 8px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
        margin-bottom: 10px;
        font-size: 0.75rem;
        text-align: center;
    }
    .summary-title { font-weight: 700; color: #1e293b; margin-bottom: 4px; font-size: 0.8rem; }
    
    /* 테이블을 감싸는 스크롤 컨테이너 (모바일 찌그러짐 방지) */
    .table-container {
        width: 100%;
        overflow-x: auto;
        margin-bottom: 10px;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }

    /* 모바일 테이블 스타일 */
    .mobile-table {
        width: 100%;
        min-width: 520px; /* 최소 너비를 확보하여 글자 깨짐 방지 */
        border-collapse: collapse;
        font-size: 0.62rem;
        background: white;
        table-layout: fixed;
    }
    .mobile-table th, .mobile-table td {
        padding: 5px 2px;
        text-align: center;
        border-bottom: 1px solid #f1f5f9;
        color: #334155;
        word-break: break-all;
    }
    .mobile-table th {
        background-color: #1e293b;
        color: white;
        font-weight: 600;
        font-size: 0.65rem;
    }
    .mobile-table tr:nth-child(even) { background-color: #f8fafc; }

    /* 소계 2단 분리 행 스타일 (완벽한 줄바꿈 방지) */
    .subtotal-row-1 {
        background-color: #eff6ff !important;
        font-weight: bold;
    }
    .subtotal-row-1 td {
        border-top: 1.5px solid #cbd5e1;
        border-bottom: none !important;
        color: #1e293b;
        white-space: nowrap !important;
    }
    .subtotal-row-2 {
        background-color: #eff6ff !important;
        font-weight: bold;
    }
    .subtotal-row-2 td {
        border-top: none !important;
        border-bottom: 1.5px solid #cbd5e1;
        color: #1d4ed8;
        font-size: 0.52rem;
        white-space: nowrap !important; /* 퍼센티지 찌그러짐 원천 차단 */
    }

    /* 컬럼 폭 지정 */
    .col-jibu { width: 55px; }
    .col-site { width: 110px; text-align: left !important; padding-left: 4px !important; font-size: 0.58rem; }
    .col-tower { width: 95px; text-align: left !important; padding-left: 4px !important; font-size: 0.58rem; }
    .col-etc { width: 45px; }
    .col-union { width: 42px; }
    .col-total { width: 45px; font-weight: bold; }

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
        total_counts = df1_raw.iloc[1]
        total_ratios = df1_raw.iloc[2]
        
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

        unions = ['한노', '민노', '섬유', '건산', '직원', '기타']
        subtotal_data_map = {}
        
        for idx, row in df0.iterrows():
            jibu_name = str(row['지부']).strip()
            if '소계' in jibu_name:
                clean_jibu = jibu_name.replace(' 소계', '').replace('지부', '') + '지부'
                match_stat = df_branch_stats[df_branch_stats['지부'].astype(str).str.contains(clean_jibu, na=False)]
                
                count_dict = {}
                percent_dict = {}
                
                if not match_stat.empty:
                    for union in unions:
                        col_cnt = f"{union} (대)"
                        col_pct = f"{union} (%)"
                        if col_cnt in match_stat.columns and col_pct in match_stat.columns:
                            c_val = match_stat[col_cnt].values[0]
                            p_val = match_stat[col_pct].values[0]
                            
                            if pd.notna(c_val) and c_val != "-":
                                count_dict[union] = f"{c_val}"
                            else:
                                count_dict[union] = "-"
                                
                            if pd.notna(p_val) and p_val != "-":
                                try:
                                    p_num = float(p_val) * 100
                                    percent_dict[union] = f"({p_num:.1f}%)"
                                except:
                                    percent_dict[union] = f"({str(p_val)})"
                            else:
                                percent_dict[union] = "-"
                
                subtotal_data_map[idx] = {
                    "counts": count_dict,
                    "percents": percent_dict
                }

        # --- 타워사별 현황 데이터 읽기 ---
        df2_raw = pd.read_excel(target_file, sheet_name=2)
        df2 = df2_raw.iloc[1:].copy() if len(df2_raw) > 1 else df2_raw.copy()
        df2.columns = df2_raw.iloc[0].astype(str).str.strip()
        df2 = df2.loc[:, ~df2.columns.str.contains('^Unnamed')]
        df2 = df2.fillna("-")

        if '한노점유율' in df2.columns and '타워사' in df2.columns:
            cols = ['타워사', '한노점유율'] + [c for c in df2.columns if c not in ['타워사', '한노점유율']]
            df2 = df2[cols]
            def fmt_hano_ratio(val):
                try:
                    v = float(val)
                    return f"<span style='color:#dc2626; font-weight:900; white-space:nowrap;'>{v * 100:.1f}%</span>"
                except:
                    return str(val)
            df2['한노점유율'] = df2['한노점유율'].apply(fmt_hano_ratio)

        summary_info = {
            "한노": f"{total_counts.get('한노', 51)}대 ({float(total_ratios.get('한노', 0.2786))*100:.1f}%)",
            "민노": f"{total_counts.get('민노', 79)}대 ({float(total_ratios.get('민노', 0.4316))*100:.1f}%)",
            "섬유": f"{total_counts.get('섬유', 6)}대 ({float(total_ratios.get('섬유', 0.0327))*100:.1f}%)",
            "건산": f"{total_counts.get('건산', 9)}대 ({float(total_ratios.get('건산', 0.0491))*100:.1f}%)",
            "직원": f"{total_counts.get('직원', 32)}대 ({float(total_ratios.get('직원', 0.1748))*100:.1f}%)",
            "기타": f"{total_counts.get('기타', 6)}대 ({float(total_ratios.get('기타', 0.0327))*100:.1f}%)",
            "합계": f"{total_counts.get('합계', 183)}대 (100.0%)"
        }

        return df0, df2, summary_info, subtotal_data_map, None
    except Exception as e:
        return None, None, None, None, str(e)


df_main, df_sub2, summary, subtotal_map, err_msg = load_and_merge_data()

if err_msg or df_main is None:
    st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {err_msg}")
else:
    tabs = st.tabs(["📋 지부별 현장 및 통계", "🏗️ 타워사별 현황"])

    # --- [탭 1] 지부별 현장명단 및 상단 메인 통계 카드 ---
    with tabs[0]:
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

        unique_jibus = sorted(list(df_main[df_main['지부'] != '-']['지부'].apply(lambda x: str(x).replace(' 소계', '').replace('지부', '').strip()).unique()))
        jibu_options = ["전체보기"] + [f"{j}지부" for j in unique_jibus if j and j != 'nan']
        
        selected_jibu = st.selectbox("📍 지부 선택", jibu_options, label_visibility="collapsed")
        
        filtered_df0 = df_main.copy()
        if selected_jibu != "전체보기":
            target_prefix = selected_jibu.replace('지부', '')
            filtered_df0 = filtered_df0[filtered_df0['지부'].astype(str).str.contains(target_prefix, na=False)]

        st.markdown(f"<p style='font-size: 0.7rem; color: #64748b; margin: 4px 0;'>조회 결과: <b>{len(filtered_df0)}</b>건</p>", unsafe_allow_html=True)
        
        # 스크롤 가능한 컨테이너로 테이블 감싸기
        html_table = "<div class='table-container'><table class='mobile-table'><thead><tr>"
        cols = list(filtered_df0.columns)
        for col in cols:
            if col == '지부': cls = "col-jibu"
            elif col == '현장명': cls = "col-site"
            elif col == '타워회사': cls = "col-tower"
            elif col in ['총대수', '합계']: cls = "col-total"
            elif col == '특이사항': cls = "col-etc"
            else: cls = "col-union"
            html_table += f"<th class='{cls}'>{col}</th>"
        html_table += "</tr></thead><tbody>"
        
        for idx, row in filtered_df0.iterrows():
            jibu_val = str(row['지부'])
            is_subtotal = '소계' in jibu_val
            
            if is_subtotal:
                clean_jibu_name = jibu_val.replace(' 소계', '')
                sub_info = subtotal_map.get(idx, {"counts": {}, "percents": {}})
                counts = sub_info.get("counts", {})
                percents = sub_info.get("percents", {})
                
                # [1행] 대수
                html_table += "<tr class='subtotal-row-1'>"
                for col in cols:
                    if col == '지부': cls = "col-jibu"
                    elif col == '현장명': cls = "col-site"
                    elif col == '타워회사': cls = "col-tower"
                    elif col in ['총대수', '합계']: cls = "col-total"
                    elif col == '특이사항': cls = "col-etc"
                    else: cls = "col-union"
                    
                    if col == '지부':
                        html_table += f"<td class='{cls}' rowspan='2' style='vertical-align: middle;'><b>{clean_jibu_name}</b></td>"
                    elif col == '현장명':
                        html_table += f"<td class='{cls}' style='text-align: left; padding-left: 4px; font-weight: bold;'>소계(대수)</td>"
                    elif col == '타워회사':
                        html_table += f"<td class='{cls}'>-</td>"
                    else:
                        val = counts.get(col, str(row[col]))
                        html_table += f"<td class='{cls}'>{val}</td>"
                html_table += "</tr>"
                
                # [2행] 퍼센티지 (%)
                html_table += "<tr class='subtotal-row-2'>"
                for col in cols:
                    if col == '지부': continue
                    if col == '현장명': cls = "col-site"
                    elif col == '타워회사': cls = "col-tower"
                    elif col in ['총대수', '합계']: cls = "col-total"
                    elif col == '특이사항': cls = "col-etc"
                    else: cls = "col-union"
                    
                    if col == '현장명':
                        html_table += f"<td class='{cls}' style='text-align: left; padding-left: 4px; font-weight: bold;'>점유율(%)</td>"
                    elif col == '타워회사':
                        html_table += f"<td class='{cls}'>-</td>"
                    else:
                        val = percents.get(col, "-")
                        html_table += f"<td class='{cls}'>{val}</td>"
                html_table += "</tr>"
                
            else:
                html_table += "<tr>"
                for col in cols:
                    if col == '지부': cls = "col-jibu"
                    elif col == '현장명': cls = "col-site"
                    elif col == '타워회사': cls = "col-tower"
                    elif col in ['총대수', '합계']: cls = "col-total"
                    elif col == '특이사항': cls = "col-etc"
                    else: cls = "col-union"
                    
                    val = str(row[col])
                    html_table += f"<td class='{cls}'>{val}</td>"
                html_table += "</tr>"
                
        html_table += "</tbody></table></div>"
        st.markdown(html_table, unsafe_allow_html=True)

    # --- [탭 2] 타워사별 현황 ---
    with tabs[1]:
        tower_list = ["전체보기"]
        if df_sub2 is not None and '타워사' in df_sub2.columns:
            tower_list += sorted(list(df_sub2['타워사'].astype(str).unique()))
        
        selected_tower = st.selectbox("🔍 타워사 선택/검색", tower_list, label_visibility="collapsed")
        
        filtered_df2 = df_sub2.copy() if df_sub2 is not None else None
        if selected_tower != "전체보기" and filtered_df2 is not None:
            filtered_df2 = filtered_df2[filtered_df2['타워사'].astype(str).str.contains(selected_tower, case=False, na=False)]

        st.markdown(f"<p style='font-size: 0.7rem; color: #64748b; margin: 4px 0;'>검색 결과: <b>{len(filtered_df2) if filtered_df2 is not None else 0}</b>건</p>", unsafe_allow_html=True)
        
        if filtered_df2 is not None and not filtered_df2.empty:
            html_table2 = "<div class='table-container'><table class='mobile-table'><thead><tr>"
            for col in filtered_df2.columns:
                html_table2 += f"<th>{col}</th>"
            html_table2 += "</tr></thead><tbody>"
            
            for _, row in filtered_df2.iterrows():
                html_table2 += "<tr>"
                for col in filtered_df2.columns:
                    val = str(row[col])
                    html_table2 += f"<td>{val}</td>"
                html_table2 += "</tr>"
            html_table2 += "</tbody></table></div>"
            
            st.markdown(html_table2, unsafe_allow_html=True)
        else:
            st.info("데이터가 없습니다.")

    st.markdown("---")
    st.markdown("<p style='text-align: center; font-size: 0.65rem; color: #94a3b8;'>Gyeonggi Regional Headquarters Dashboard</p>", unsafe_allow_html=True)
