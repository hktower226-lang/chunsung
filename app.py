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
        
        # 1. 전체 소속별 점유율 (상단 고정 카드)
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
            
            st.markdown(
                f"""
                <div class="metric-card" style="{card_style} display: flex; justify-content: space-between; align-items: center;">
                    <div><span class="{b_class}">{col}</span></div>
                    <div>
                        <span style="font-size: 17px; font-weight: bold;">{val:,.0f}명</span> &nbsp;|&nbsp; 
                        <span style="font-size: 17px; color: {'#e53e3e' if is_hanno else '#2b6cb0'}; font-weight: bold;">{pct:.1f}%</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.write("---")
        st.subheader("🔍 경기지역본부 현장 및 지부 검색")

        # 2. 지부별 / 현장별 데이터 가공
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
            is_subtotal = "소계" in str(row["지부"]) or pd.isna(row["현장명"])
            h, m, s, g, st_f, u, total = row["한노"], row["민노"], row["섬유"], row["건산"], row["직원"], row["미정"], row["총대수"]

            if is_subtotal:
                h_pct = (h / total * 100) if total > 0 else 0
                m_pct = (m / total * 100) if total > 0 else 0
                s_pct = (s / total * 100) if total > 0 else 0
                g_pct = (g / total * 100) if total > 0 else 0
                st_f_pct = (st_f / total * 100) if total > 0 else 0
                u_pct = (u / total * 100) if total > 0 else 0

                st.markdown(
                    f"""
                    <div class="subtotal-card">
                        <div style="font-size: 16px; font-weight: bold; color: #2b6cb0; margin-bottom: 6px;">
                            📌 [{row['지부']}] 합계 (총 대수: {total:,.0f}대) | <span style="color: #e53e3e;">한노 점유율: {h_pct:.1f}%</span>
                        </div>
                        <div class="item-container">
                            <span class="badge-hanno">한노 {h:.0f} ({h_pct:.1f}%)</span>
                            <span class="badge-minno">민노 {m:.0f} ({m_pct:.1f}%)</span>
                            <span class="badge-seomvu">섬유 {s:.0f} ({s_pct:.1f}%)</span>
                            <span class="badge-geonsan">건산 {g:.0f} ({g_pct:.1f}%)</span>
                            <span class="badge-staff">직원 {st_f:.0f} ({st_f_pct:.1f}%)</span>
                            <span class="badge-etc">미정 {u:.0f} ({u_pct:.1f}%)</span>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                h_pct = (h / total * 100) if total > 0 else 0
                st.markdown(
                    f"""
                    <div class="metric-card">
                        <div style="font-weight: bold; font-size: 15px; color: #1a202c;">{row['현장명']}</div>
                        <div style="font-size: 13px; color: #718096; margin-bottom: 4px;">타워사: {row['타워회사']} | 총 대수: <b>{total:,.0f}대</b> | <span style="color: #e53e3e; font-weight: bold;">한노 {h_pct:.1f}%</span></div>
                        <div class="item-container">
                            <span class="badge-hanno">한노 {h:.0f}</span>
                            <span class="badge-minno">민노 {m:.0f}</span>
                            <span class="badge-seomvu">섬유 {s:.0f}</span>
                            <span class="badge-geonsan">건산 {g:.0f}</span>
                            <span class="badge-staff">직원 {st_f:.0f}</span>
                            <span class="badge-etc">미정 {u:.0f}</span>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        st.write("---")
        st.subheader("🏗️️ 임대사별 한노 점유율 검색")

        t3 = df_ratio.iloc[81:, :10].copy()
        t3.columns = ["타워사", "현장수", "한노", "민노", "섬유", "건산", "직원", "기타", "합계", "한노점유율"]
        t3 = t3.dropna(subset=["타워사"])

        for col in ["현장수", "한노", "민노", "섬유", "건산", "직원", "기타", "합계", "한노점유율"]:
            t3[col] = pd.to_numeric(t3[col], errors="coerce").fillna(0)

        t3 = t3.dropna(subset=["합계"])
        t3 = t3.sort_values(by="한노점유율", ascending=False).reset_index(drop=True)

        tower_list = t3["타워사"].tolist()
        selected_tower = st.selectbox("임대사(타워사) 선택", ["전체 보기"] + tower_list, key="tab1_tower")

        if selected_tower != "전체 보기":
            filtered_t3 = t3[t3["타워사"] == selected_tower]
        else:
            filtered_t3 = t3

        for _, row in filtered_t3.iterrows():
            hanno_share = row["한노점유율"] * 100 if pd.notna(row["한노점유율"]) else 0
            st.markdown(
                f"""
                <div class="metric-card">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                        <span style="font-size: 15px; font-weight: bold; color: #2d3748;">{row['타워사']}</span>
                        <span style="font-size: 15px; font-weight: bold; color: #e53e3e;">한노 점유율: {hanno_share:.1f}%</span>
                    </div>
                    <div style="font-size: 13px; color: #718096; margin-bottom: 4px;">현장수: {row['현장수']}개 | 총합계: {row['합계']:.0f}명</div>
                    <div class="item-container">
                        <span class="badge-hanno">한노 {row['한노']:.0f}</span>
                        <span class="badge-minno">민노 {row['민노']:.0f}</span>
                        <span class="badge-seomvu">섬유 {row['섬유']:.0f}</span>
                        <span class="badge-geonsan">건산 {row['건산']:.0f}</span>
                        <span class="badge-staff">직원 {row['직원']:.0f}</span>
                        <span class="badge-etc">기타 {row['기타']:.0f}</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    except Exception as e:
        st.warning(f"데이터 처리 중 오류 발생: {e}")

# -------------------------------------------------------------------------
# [탭 2] 반도체현장 시트
# -------------------------------------------------------------------------
with tab2:
    st.header("⚡ 반도체현장 타워크레인 설치 및 노조별 대수")
    st.markdown("반도체 현장 전체 대수 비교 및 현장명 검색 기능입니다.")

    try:
        df_semi = pd.read_excel(file_path, sheet_name="반도체현장")
        
        total_summary_val = df_semi.iloc[4, 0]
        h_semi = df_semi.iloc[4, 1]
        m_semi = df_semi.iloc[4, 2]
        c_semi = df_semi.iloc[4, 3]
        b_semi = df_semi.iloc[4, 4]

        st.markdown(
            f"""
            <div class="subtotal-card">
                <div style="font-size: 16px; font-weight: bold; color: #2b6cb0; margin-bottom: 6px;">
                    📌 반도체 현장 전체 대수 비교 (총 대수: {total_summary_val}대)
                </div>
                <div class="item-container">
                    <span class="badge-hanno">한국노총 {h_semi}</span>
                    <span class="badge-minno">민주노총 {m_semi}</span>
                    <span class="badge-seomvu">건설노조 {c_semi}</span>
                    <span class="badge-etc">비노조 {b_semi}</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.write("---")
        st.subheader("🔍 반도체 현장명 검색")

        t_semi_details = df_semi.iloc[8:, :9].copy()
        t_semi_details.columns = ["No", "현장명", "전체대수", "한국노총", "민주노총", "건설노조", "비노조", "미정", "비고"]
        t_semi_details = t_semi_details.dropna(subset=["현장명"])

        for col in ["전체대수", "한국노총", "민주노총", "건설노조", "비노조", "미정"]:
            t_semi_details[col] = pd.to_numeric(t_semi_details[col], errors="coerce").fillna(0)

        semi_sites = t_semi_details["현장명"].tolist()
        selected_semi_site = st.selectbox("현장명을 선택하세요", ["전체 보기"] + semi_sites, key="tab2_site")

        if selected_semi_site != "전체 보기":
            filtered_semi = t_semi_details[t_semi_details["현장명"] == selected_semi_site]
        else:
            filtered_semi = t_semi_details

        for _, row in filtered_semi.iterrows():
            st.markdown(
                f"""
                <div class="metric-card">
                    <div style="font-weight: bold; font-size: 15px; color: #1a202c;">{row['현장명']}</div>
                    <div style="font-size: 13px; color: #718096; margin-bottom: 4px;">전체 대수: <b>{row['전체대수']:,.0f}대</b></div>
                    <div class="item-container">
                        <span class="badge-hanno">한국노총 {row['한국노총']:.0f}</span>
                        <span class="badge-minno">민주노총 {row['민주노총']:.0f}</span>
                        <span class="badge-seomvu">건설노조 {row['건설노조']:.0f}</span>
                        <span class="badge-etc">비노조 {row['비노조']:.0f}</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    except Exception as e:
        st.warning(f"데이터 처리 중 오류 발생: {e}")

# -------------------------------------------------------------------------
# [탭 3] 과거 채용비율 ('목록' 시트 데이터 활용)
# -------------------------------------------------------------------------
with tab3:
    st.header("📈 24년 9월 ~ 25년 12월 채용비율 현황")
    st.markdown("전체 채용비율 요약 및 임대사/원청사 검색 기능입니다.")

    try:
        df_list = pd.read_excel(file_path, sheet_name="목록")
        
        st.markdown(
            """
            <div class="subtotal-card">
                <div style="font-size: 16px; font-weight: bold; color: #2b6cb0; margin-bottom: 4px;">
                    📌 24년 9월 ~ 25년 12월 전체 채용비율 현황
                </div>
                <div style="font-size: 14px; color: #4a5568;">
                    해당 기간 동안의 임대사 및 원청사별 타워크레인 투입 및 노조 구성 내역입니다.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.write("---")
        st.subheader("🔍 임대사 / 원청사 검색")

        df_clean = df_list.iloc[2:, [3, 4, 1, 6]].copy()
        df_clean.columns = ["임대사", "원청사", "본부", "대형소형"]
        df_clean = df_clean.dropna(subset=["임대사"])

        search_keyword = st.text_input("임대사 또는 원청사 이름을 입력하세요")

        if search_keyword:
            searched_df = df_clean[
                df_clean["임대사"].astype(str).str.contains(search_keyword, na=False) | 
                df_clean["원청사"].astype(str).str.contains(search_keyword, na=False)
            ]
        else:
            searched_df = df_clean.head(20)

        if not searched_df.empty:
            for _, row in searched_df.iterrows():
                st.markdown(
                    f"""
                    <div class="metric-card">
                        <div style="font-weight: bold; font-size: 15px; color: #2b6cb0;">임대사: {row['임대사']}</div>
                        <div style="font-size: 14px; color: #2d3748; margin-top: 2px;">원청사: {row['원청사']}</div>
                        <div style="font-size: 12px; color: #718096; margin-top: 4px;">본부: {row.get('본부', 'N/A')} | 규격: {row.get('대형소형', 'N/A')}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
        else:
            st.info("검색 결과가 없습니다.")

    except Exception as e:
        st.warning(f"데이터 처리 중 오류 발생: {e}")

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>모바일 화면 최적화 완료</p>",
    unsafe_allow_html=True,
)
