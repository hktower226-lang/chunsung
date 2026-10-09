import pandas as pd
import streamlit as st

# ==========================================
# 1. 페이지 기본 설정 및 모바일 CSS 스타일링
# ==========================================
st.set_page_config(
    page_title="현장 점유율 및 채용 현황",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    /* 전체 글꼴 및 기본 폰트 크기 */
    html, body, [class*="css"] {
        font-size: 17px !important;
    }
    
    /* 사이드바 메뉴 1~5번 글씨 크기 크게 키우기 */
    [data-testid="stSidebar"] .stRadio label {
        font-size: 1.25rem !important;
        font-weight: bold !important;
    }
    
    /* 제목 및 헤더 크기 */
    h1 { font-size: 1.9rem !important; font-weight: bold !important; color: #1E3A8A; }
    h2 { font-size: 1.5rem !important; font-weight: bold !important; color: #1E40AF; }
    h3 { font-size: 1.3rem !important; font-weight: bold !important; }
    
    /* 카드/Expander 내부 텍스트 */
    .stSelectbox label, .stRadio label, .stMultiSelect label {
        font-size: 1.1rem !important;
        font-weight: bold !important;
    }
    
    /* 여백 설정 */
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 2.5rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 100% !important;
    }
    
    /* 테이블 글씨 크기 */
    .dataframe {
        font-size: 15px !important;
    }
    
    /* 전화번호 연결 버튼 스타일 */
    .phone-btn {
        display: inline-block;
        background-color: #25D366;
        color: white !important;
        padding: 9px 16px;
        font-size: 16px;
        font-weight: bold;
        text-decoration: none;
        border-radius: 8px;
        margin-top: 5px;
        margin-bottom: 8px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }

    /* 메인 화면 카드 스타일 */
    .main-total-card {
        background: linear-gradient(135deg, #1E3A8A, #3B82F6);
        color: white;
        padding: 14px 18px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 15px;
        box-shadow: 0 3px 5px rgba(0,0,0,0.1);
    }
    .main-stat-card {
        background-color: #FFFFFF;
        border-left: 6px solid #3B82F6;
        border-radius: 8px;
        padding: 10px 14px;
        margin-bottom: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
    }
    .border-hanno { border-left-color: #EF4444 !important; }
    .border-minno { border-left-color: #3B82F6 !important; }
    .border-seomoo { border-left-color: #10B981 !important; }
    .border-geunsan { border-left-color: #F59E0B !important; }
    .border-jikwon { border-left-color: #8B5CF6 !important; }
    .border-mijeong { border-left-color: #6B7280 !important; }
    </style>
""",
    unsafe_allow_html=True,
)

EXCEL_FILE = "채용비율.xlsx"

@st.cache_data(ttl=5)
def load_all_sheets():
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        sheets = {sheet: pd.read_excel(xls, sheet) for sheet in xls.sheet_names}
        return sheets
    except Exception as e:
        st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다. '{EXCEL_FILE}' 파일이 같은 폴더에 있는지 확인해주세요.\n오류 내용: {e}")
        return None

sheets = load_all_sheets()
if sheets is None:
    st.stop()

sheet_names = list(sheets.keys())
s1_name = sheet_names[0] if len(sheet_names) > 0 else "1.26년 전체 소속별 점유율"
s2_name = sheet_names[1] if len(sheet_names) > 1 else "2.각8개지부 현장점유율"
s3_name = sheet_names[2] if len(sheet_names) > 2 else "3.타워사별 점유현황"
s4_name = sheet_names[3] if len(sheet_names) > 3 else "4.반도체현장"
s5_name = sheet_names[4] if len(sheet_names) > 4 else "5.1~9월채용추이"

def fmt_pct(val):
    try:
        if pd.isna(val) or val == "" or val == "-":
            return "0%"
        f = float(val)
        if f <= 1.0:
            return f"{f * 100:.1f}%"
        return f"{f:.1f}%"
    except:
        return str(val) if pd.notna(val) else "0%"

# ==========================================
# 사이드바 메뉴 설정
# ==========================================
st.sidebar.title("📌 채용 비율 메뉴 선택")

menu_options = [
    "📊 1. 2026년 전체 소속별 점유율",
    "🏢 2. 각 8개지부 현장 점유율",
    "🏗️ 3. 타워사별 점유 현황",
    "🏭 4. 반도체 현장 현황",
    "📅 5. 1~9월 채용 추이",
]

menu = st.sidebar.radio("원하시는 화면을 선택하세요", menu_options)

# ==========================================
# 메뉴별 화면 구현
# ==========================================

# ------------------------------------------
# [메뉴 1] 전체 소속별 점유율
# ------------------------------------------
if menu == "📊 1. 2026년 전체 소속별 점유율":
    st.title("📊 2026년 전체 소속별 점유율")
    df1 = sheets[s1_name].copy()

    try:
        tot_cnt = df1.iloc[1, 7]
        tot_pct = df1.iloc[1, 8]

        stat_items = [
            {"title": "한국노총 (한노)", "cnt": df1.iloc[1, 1], "pct": fmt_pct(df1.iloc[2, 1]), "class": "border-hanno", "color": "#EF4444"},
            {"title": "민주노총 (민노)", "cnt": df1.iloc[1, 2], "pct": fmt_pct(df1.iloc[2, 2]), "class": "border-minno", "color": "#3B82F6"},
            {"title": "섬유노조", "cnt": df1.iloc[1, 3], "pct": fmt_pct(df1.iloc[2, 3]), "class": "border-seomoo", "color": "#10B981"},
            {"title": "건설산업 (건산)", "cnt": df1.iloc[1, 4], "pct": fmt_pct(df1.iloc[2, 4]), "class": "border-geunsan", "color": "#F59E0B"},
            {"title": "직원", "cnt": df1.iloc[1, 5], "pct": fmt_pct(df1.iloc[2, 5]), "class": "border-jikwon", "color": "#8B5CF6"},
            {"title": "미정", "cnt": df1.iloc[1, 6], "pct": fmt_pct(df1.iloc[2, 6]), "class": "border-mijeong", "color": "#6B7280"},
        ]

        st.markdown(
            f"""
        <div class="main-total-card">
            <span style="font-size: 1.05rem; font-weight:bold;">🏆 전체 현장 총 대수: </span>
            <span style="font-size: 1.5rem; font-weight:bold; margin-left:8px;">{tot_cnt}대 ({tot_pct})</span>
        </div>
        """,
            unsafe_allow_html=True,
        )

        st.markdown("<div style='font-size: 1.1rem; font-weight: bold; color: #1E40AF; margin-bottom: 8px;'>👇 소속별 점유 현황</div>", unsafe_allow_html=True)

        for item in stat_items:
            st.markdown(
                f"""
            <div class="main-stat-card {item['class']}">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-size: 1.0rem; font-weight:bold; color:{item['color']};">{item['title']}</span>
                    <span style="font-size: 1.15rem; font-weight:bold; color:#111827;">{item['pct']}</span>
                </div>
                <div style="margin-top: 4px; font-size: 0.95rem; color: #4B5563;">
                    보유 수량: <b>{item['cnt']}대</b>
                </div>
            </div>
            """,
                unsafe_allow_html=True,
            )

    except Exception:
        st.write("### 💡 전체 소속별 통계 개요")
        st.dataframe(df1.fillna(""), use_container_width=True, hide_index=True)


# ------------------------------------------
# [메뉴 2] 각 8개지부 현장 점유율 (통합 검색 기능 추가)
# ------------------------------------------
elif menu == "🏢 2. 각 8개지부 현장 점유율":
    st.title("🏢 지부별 현장 점유율")
    df2 = sheets[s2_name].copy()
    rows = df2.values.tolist()

    branches = ["북부지부", "남부지부", "남서지부", "동부지부", "서부지부", "북서지부", "용인지부", "중부지부"]

    branch_data = {}
    current_branch = "북부지부"

    for r in rows[1:]:
        branch_col = str(r[0]).strip() if pd.notna(r[0]) else ""
        for b in ["북부", "남부", "남서", "동부", "서부", "북서", "용인", "중부"]:
            if branch_col == b or branch_col == b + "지부":
                current_branch = b + "지부"
                break

        if current_branch not in branch_data:
            branch_data[current_branch] = {"summary": None, "percent": None, "sites": []}

        if "퍼센테이지" in branch_col or "소계퍼센테이지" in branch_col:
            branch_data[current_branch]["percent"] = r
        elif "소계" in branch_col:
            branch_data[current_branch]["summary"] = r
        elif pd.notna(r[1]) and str(r[1]).strip() != "현장명":
            branch_data[current_branch]["sites"].append(r)

    # 2번째 메뉴 통합 검색창
    st.subheader("🔍 타워사/현장명 통합 검색")
    search_query = st.text_input("타워사 또는 현장명을 입력하면 어느 지부인지 바로 확인 가능합니다", "", placeholder="예: 백산, 복정 등")
    
    if search_query:
        all_searched_sites = []
        for b_name, b_info in branch_data.items():
            for s in b_info["sites"]:
                site_name = str(s[1]) if len(s) > 1 and pd.notna(s[1]) else ""
                tower_name = str(s[2]) if len(s) > 2 and pd.notna(s[2]) else ""
                if search_query.lower() in site_name.lower() or search_query.lower() in tower_name.lower():
                    all_searched_sites.append({
                        "지부명": b_name,
                        "현장명": site_name,
                        "타워회사": tower_name,
                        "민노": s[5] if len(s) > 5 else 0,
                        "한노": s[4] if len(s) > 4 else 0,
                        "건산": s[7] if len(s) > 7 else 0,
                        "섬유": s[6] if len(s) > 6 else 0,
                        "총대수": s[10] if len(s) > 10 else 0
                    })
        if all_searched_sites:
            st.markdown(f"### 🔎 '{search_query}' 검색 결과")
            st.dataframe(pd.DataFrame(all_searched_sites), use_container_width=True, hide_index=True)
        else:
            st.warning("검색 결과가 없습니다.")
        st.markdown("---")

    st.subheader("📌 각 8개 지부 현장 점유율")

    for b_name in branches:
        if b_name in branch_data:
            b_info = branch_data[b_name]
            sum_row = b_info["summary"]
            pct_row = b_info["percent"]
            sites_list = b_info["sites"]

            if sum_row is not None and pct_row is not None:
                hanno_cnt, hanno_pct = sum_row[4], fmt_pct(pct_row[4])
                minno_cnt, minno_pct = sum_row[5], fmt_pct(pct_row[5])
                seomoo_cnt, seomoo_pct = sum_row[6], fmt_pct(pct_row[6])
                geunsan_cnt, geunsan_pct = sum_row[7], fmt_pct(pct_row[7])
                jikwon_cnt, jikwon_pct = sum_row[8], fmt_pct(pct_row[8])
                mijeong_cnt, mijeong_pct = sum_row[9], fmt_pct(pct_row[9])
                total_cnt = sum_row[10] if len(sum_row) > 10 else "-"

                expander_label = (
                    f"📁 **{b_name}** | "
                    f":red[**🔥 한노 {hanno_cnt}({hanno_pct})**] | "
                    f":blue[민노 {minno_cnt}({minno_pct})] | "
                    f":green[섬유 {seomoo_cnt}({seomoo_pct})] | "
                    f":orange[건산 {geunsan_cnt}({geunsan_pct})] | "
                    f":violet[직원 {jikwon_cnt}({jikwon_pct})] | "
                    f":gray[미정 {mijeong_cnt}({mijeong_pct})] | "
                    f"**총대수 {total_cnt}(100%)**"
                )
            else:
                expander_label = f"📁 **{b_name}**"

            with st.expander(expander_label):
                if sites_list:
                    site_df = pd.DataFrame(sites_list)
                    display_df = pd.DataFrame(
                        {
                            "현장명": site_df[1],
                            "타워회사": site_df[2],
                            "한노": site_df[4],
                            "민노": site_df[5],
                            "섬유": site_df[6],
                            "건산": site_df[7],
                            "직원": site_df[8],
                            "미정": site_df[9],
                            "총대수": site_df[10],
                        }
                    ).fillna(0)

                    st.markdown(f"📍 **{b_name} 상세 현장 목록**")
                    st.dataframe(display_df, use_container_width=True, hide_index=True)
                else:
                    st.info("해당 지부의 상세 현장 데이터가 없습니다.")


# ------------------------------------------
# [메뉴 3] 타워사별 점유 현황 (한노점유율 타워사 바로 다음 배치 및 빨간색 굵게 강조)
# ------------------------------------------
elif menu == "🏗️ 3. 타워사별 점유 현황":
    st.title("🏗️ 임대사(타워사)별 점유 현황")
    df3 = sheets[s3_name].copy()

    if "타워사" in df3.iloc[0].values:
        df3.columns = df3.iloc[0]
        df3 = df3[1:].reset_index(drop=True)

    df3 = df3.dropna(subset=["타워사"]).fillna(0)

    # 컬럼 재배치: '한노점유율'을 타워사 바로 다음(맨 앞)으로 이동
    cols = list(df3.columns)
    if "한노점유율" in cols and "타워사" in cols:
        cols.remove("한노점유율")
        tower_idx = cols.index("타워사")
        cols.insert(tower_idx + 1, "한노점유율")
        df3 = df3[cols]

    if "한노점유율" in df3.columns:
        df3["한노점유율"] = df3["한노점유율"].apply(fmt_pct)

    st.subheader("🔍 임대사(타워사) 검색")
    tower_list = ["전체 보기"] + list(df3["타워사"].astype(str).unique())
    selected_tower = st.selectbox("타워사를 선택하세요", tower_list)

    search_kw = st.text_input("또는 타워사 이름 일부 직접 입력", "", placeholder="예: 비엠, 대원 등")

    filtered_df = df3.copy()
    if selected_tower and selected_tower != "전체 보기":
        filtered_df = filtered_df[filtered_df["타워사"].astype(str) == selected_tower]

    if search_kw:
        filtered_df = filtered_df[filtered_df["타워사"].astype(str).str.contains(search_kw, case=False)]

    st.markdown("### 📋 타워사 점유 현황 목록")
    
    # 한노점유율 빨간색 굵게 스타일링 적용
    def style_menu3(df):
        styler = df.style
        if '한노점유율' in df.columns:
            styler = styler.applymap(lambda x: 'color: #DC2626; font-weight: bold;', subset=['한노점유율'])
        return styler

    st.dataframe(style_menu3(filtered_df), use_container_width=True, hide_index=True)


# ------------------------------------------
# [메뉴 4] 반도체 현장 현황
# ------------------------------------------
elif menu == "🏭 4. 반도체 현장 현황":
    st.title("🏭 반도체 현장 타워크레인 현황")
    df4 = sheets[s4_name].copy()

    st.subheader("📊 반도체 현장 전체 대수 비교")
    try:
        summary_data = {
            "구분": ["전체 대수", "한국노총", "민주노총", "건설노조(건산,섬유)", "비노조", "미정"],
            "대수": [df4.iloc[4, 0], df4.iloc[4, 1], df4.iloc[4, 2], df4.iloc[4, 3], df4.iloc[4, 4], df4.iloc[4, 5]],
            "점유율": ["-", fmt_pct(df4.iloc[5, 1]), fmt_pct(df4.iloc[5, 2]), fmt_pct(df4.iloc[5, 3]), fmt_pct(df4.iloc[5, 4]), fmt_pct(df4.iloc[5, 5])],
        }
        st.table(pd.DataFrame(summary_data))
    except Exception:
        st.info("상단 요약 데이터 표시 중")

    st.markdown("---")
    st.subheader("📍 각 현장별 상세 현황 및 담당자")

    detail_rows = df4.iloc[9:].dropna(how="all").copy()

    for idx, row in detail_rows.iterrows():
        site_name = str(row[1]) if pd.notna(row[1]) else ""
        if not site_name or site_name == "nan":
            continue

        total_cnt = row[2] if pd.notna(row[2]) else 0
        hanno = row[3] if pd.notna(row[3]) else 0
        minno = row[4] if pd.notna(row[4]) else 0
        gunsan = row[5] if pd.notna(row[5]) else 0
        non_union = row[6] if pd.notna(row[6]) else 0
        mijung = row[7] if pd.notna(row[7]) else 0
        contact_info = str(row[8]) if pd.notna(row[8]) else ""

        with st.expander(f"🏢 {site_name} (총 {total_cnt}대)", expanded=True):
            st.write(f"• **한국노총:** {hanno}대 | **민주노총:** {minno}대 | **건설노조:** {gunsan}대")
            st.write(f"• **비노조:** {non_union}대 | **미정:** {mijung}대")

            if contact_info:
                import re
                phone_match = re.search(r"01[016789][-\s]?\d{3,4}[-\s]?\d{4}", contact_info)
                if phone_match:
                    phone_num = phone_match.group().replace("-", "").replace(" ", "")
                    st.markdown(
                        f"👤 **담당자:** {contact_info}<br>"
                        f'<a href="tel:{phone_num}" class="phone-btn">📞 담당자 바로 전화걸기</a>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.write(f"👤 **담당자:** {contact_info}")


# ------------------------------------------
# [메뉴 5] 1~9월 채용 추이
# ------------------------------------------
elif menu == "📅 5. 1~9월 채용 추이":
    st.title("📅 월별 채용 현황 및 추이")
    df5 = sheets[s5_name].copy()

    rows = df5.values.tolist()
    month_blocks = {}
    current_month = None
    current_data = []

    for r in rows:
        title_cell = str(r[0]).strip() if pd.notna(r[0]) else ""
        if "채용 현황" in title_cell or "채용현황" in title_cell:
            if current_month and current_data:
                month_blocks[current_month] = current_data
            current_month = title_cell
            current_data = []
        elif current_month:
            current_data.append(r)

    if current_month and current_data:
        month_blocks[current_month] = current_data

    st.subheader("📌 월별 채용현황 합계 및 비율 요약")

    for m_name, m_rows in month_blocks.items():
        sum_row = None
        pct_row = None
        for r in m_rows:
            cell0 = str(r[0]).strip() if pd.notna(r[0]) else ""
            if "합계" in cell0:
                sum_row = r
            elif "비율" in cell0:
                pct_row = r

        with st.expander(f"🗓️ {m_name} 요약 보기 (클릭하여 열기/접기)"):
            if sum_row is not None:
                st.markdown(
                    f"**[합계]** 총대수: **{sum_row[3] if len(sum_row)>3 else '-'}**대 | "
                    f"한노: {sum_row[4]} | 민노: {sum_row[5]} | 기타: {sum_row[6]} | "
                    f"직원: {sum_row[7]} | 미정: {sum_row[8]}"
                )
            if pct_row is not None:
                st.markdown(
                    f"**[비율]** "
                    f"한노: **{fmt_pct(pct_row[4])}** | 민노: **{fmt_pct(pct_row[5])}** | "
                    f"기타: {fmt_pct(pct_row[6])} | 직원: {fmt_pct(pct_row[7])} | "
                    f"미정: {fmt_pct(pct_row[8])}"
                )

    st.markdown("---")
    st.subheader("🔍 상세 채용 현황 조회할 월 선택")
    selected_m = st.selectbox("월을 선택하세요", list(month_blocks.keys()))

    if selected_m:
        m_rows = month_blocks[selected_m]
        detail_list = []
        for r in m_rows:
            cell0 = str(r[0]).strip() if pd.notna(r[0]) else ""
            if cell0 and "임대사" not in cell0 and "합계" not in cell0 and "비율" not in cell0:
                detail_list.append({
                    "임대사": r[0],
                    "원청사": r[1],
                    "현장명": r[2],
                    "총대수": r[3],
                    "한노": r[4],
                    "민노": r[5],
                    "기타": r[6],
                    "직원": r[7],
                    "미정": r[8],
                    "비고": r[9] if len(r) > 9 else "",
                })

        if detail_list:
            st.markdown(f"### 📍 {selected_m} 상세 현황")
            st.dataframe(pd.DataFrame(detail_list).fillna(""), use_container_width=True, hide_index=True)
        else:
            st.info("해당 월의 상세 데이터가 없습니다.")
