import pandas as pd
import streamlit as st

# ==========================================
# 1. 페이지 기본 설정 및 모바일 CSS
# ==========================================
st.set_page_config(
    page_title="현장 점유율 및 채용 현황", layout="wide", initial_sidebar_state="collapsed"
)

st.markdown(
    """
    <style>
    html, body, [class*="css"] { font-size: 19px !important; }
    h1 { font-size: 2.2rem !important; font-weight: bold !important; color: #1E3A8A; }
    h2 { font-size: 1.8rem !important; font-weight: bold !important; color: #1E40AF; }
    h3 { font-size: 1.5rem !important; font-weight: bold !important; }
    .stSelectbox label, .stRadio label { font-size: 1.2rem !important; font-weight: bold !important; }
    .block-container { padding-top: 1.5rem !important; padding-bottom: 3rem !important; padding-left: 0.8rem !important; padding-right: 0.8rem !important; }
    .dataframe { font-size: 16px !important; }
    .phone-btn {
        display: inline-block; background-color: #25D366; color: white !important;
        padding: 10px 18px; font-size: 18px; font-weight: bold; text-decoration: none;
        border-radius: 8px; margin-top: 5px; margin-bottom: 10px; text-align: center;
    }
    </style>
""",
    unsafe_allow_html=True,
)

EXCEL_FILE = "채용비율.xlsx"


# ==========================================
# 2. 데이터 로드
# ==========================================
@st.cache_data(ttl=5)
def load_all_sheets():
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        return {sheet: pd.read_excel(xls, sheet) for sheet in xls.sheet_names}
    except Exception as e:
        st.error(f"❌ 엑셀 파일을 읽는 중 에러가 발생했습니다: {e}")
        return None


sheets = load_all_sheets()

if not sheets:
    st.warning("⚠️ 엑셀 데이터를 로드하지 못했습니다. 파일명을 확인해 주세요.")
    st.stop()

sheet_names = list(sheets.keys())
s1_name = sheet_names[0] if len(sheet_names) > 0 else "1.26년 전체 소속별 점유율"
s2_name = sheet_names[1] if len(sheet_names) > 1 else "2.각8개지부 현장점유율"
s3_name = sheet_names[2] if len(sheet_names) > 2 else "3.타워사별 점유현황"
s4_name = sheet_names[3] if len(sheet_names) > 3 else "4.반도체현장"
s5_name = sheet_names[4] if len(sheet_names) > 4 else "5.1~9월채용추이"

# ==========================================
# 3. 사이드바 메뉴
# ==========================================
st.sidebar.title("📌 메뉴 이동")
menu = st.sidebar.radio(
    "원하시는 화면을 선택하세요",
    [
        "🏠 메인: 전체 소속별 점유율",
        "🏢 각 8개지부 현장 점유율",
        "🏗️ 타워사별 점유 현황",
        "🏭 반도체 현장 현황",
        "📅 1~9월 채용 추이",
    ],
)

# ==========================================
# 4. 메뉴별 화면 구현
# ==========================================

# 1. 메인
if menu == "🏠 메인: 전체 소속별 점유율":
    st.title("📊 2026년 전체 소속별 점유율")
    df1 = sheets[s1_name].copy()
    st.write("### 💡 전체 소속별 통계 개요")
    st.dataframe(df1.fillna(""), use_container_width=True, hide_index=True)

# 2. 각 8개지부
elif menu == "🏢 각 8개지부 현장 점유율":
    st.title("🏢 지부별 현장 점유율")
    df2 = sheets[s2_name].copy()
    rows = df2.values.tolist()

    branches = [
        "북부지부",
        "남부지부",
        "남서지부",
        "동부지부",
        "서부지부",
        "북서지부",
        "용인지부",
        "중부지부",
    ]
    branch_data = {}
    current_branch = "북부"

    for r in rows[1:]:
        branch_col = str(r[0]).strip() if pd.notna(r[0]) else ""
        for b in ["북부", "남부", "남서", "동부", "서부", "북서", "용인", "중부"]:
            if branch_col == b:
                current_branch = b + "지부"
                break

        if current_branch not in branch_data:
            branch_data[current_branch] = {
                "summary": None,
                "percent": None,
                "sites": [],
            }

        if "소계" in branch_col:
            branch_data[current_branch]["summary"] = r
        elif "퍼센테이지" in branch_col or "소계퍼센테이지" in branch_col:
            branch_data[current_branch]["percent"] = r
        elif pd.notna(r[1]) and str(r[1]).strip() != "현장명":
            branch_data[current_branch]["sites"].append(r)

    st.subheader("📌 지부별 소계 및 퍼센티지")
    for b_name in branches:
        if b_name in branch_data:
            b_info = branch_data[b_name]
            sum_row = b_info["summary"]
            pct_row = b_info["percent"]

            with st.expander(f"🔹 {b_name} 요약 보기"):
                if sum_row is not None:
                    st.markdown(
                        f"**[합계]** 총대수: **{sum_row[10] if len(sum_row)>10 else '-'}**대 | "
                        f"한노: {sum_row[4]} | 민노: {sum_row[5]} | 섬유: {sum_row[6]} | "
                        f"건산: {sum_row[7]} | 직원: {sum_row[8]} | 미정: {sum_row[9]}"
                    )
                if pct_row is not None:

                    def fmt_pct(val):
                        try:
                            return f"{float(val)*100:.1f}%"
                        except:
                            return str(val) if pd.notna(val) else "0%"

                    st.markdown(
                        f"**[점유율]** "
                        f"한노: **{fmt_pct(pct_row[4])}** | 민노: **{fmt_pct(pct_row[5])}** | "
                        f"섬유: {fmt_pct(pct_row[6])} | 건산: {fmt_pct(pct_row[7])} | "
                        f"직원: {fmt_pct(pct_row[8])} | 미정: {fmt_pct(pct_row[9])}"
                    )

    st.markdown("---")
    st.subheader("🔍 상세 현황 조회할 지부 선택")
    selected_b = st.selectbox("지부를 선택하세요", branches)

    if selected_b in branch_data and branch_data[selected_b]["sites"]:
        sites_list = branch_data[selected_b]["sites"]
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
        st.markdown(f"### 📍 {selected_b} 상세 현장 목록")
        st.dataframe(display_df, use_container_width=True, hide_index=True)

# 3. 타워사별
elif menu == "🏗️ 타워사별 점유 현황":
    st.title("🏗️ 임대사(타워사)별 점유 현황")
    df3 = sheets[s3_name].copy()
    if "타워사" in df3.iloc[0].values:
        df3.columns = df3.iloc[0]
        df3 = df3[1:].reset_index(drop=True)

    df3 = df3.dropna(subset=["타워사"]).fillna(0)
    search_keyword = st.text_input("🔍 임대사 명칭 검색", "")

    if search_keyword:
        filtered_df = df3[
            df3["타워사"].astype(str).str.contains(search_keyword, case=False)
        ]
    else:
        filtered_df = df3

    if "한노점유율" in filtered_df.columns:

        def to_pct(val):
            try:
                return f"{float(val)*100:.1f}%"
            except:
                return str(val)

        filtered_df_disp = filtered_df.copy()
        filtered_df_disp["한노점유율"] = filtered_df_disp["한노점유율"].apply(to_pct)
    else:
        filtered_df_disp = filtered_df

    st.dataframe(filtered_df_disp, use_container_width=True, hide_index=True)

# 4. 반도체현장
elif menu == "🏭 반도체 현장 현황":
    st.title("🏭 반도체 현장 타워크레인 현황")
    df4 = sheets[s4_name].copy()

    st.subheader("📊 반도체 현장 전체 대수 비교")
    try:
        summary_data = {
            "구분": [
                "전체 대수",
                "한국노총",
                "민주노총",
                "건설노조(건산,섬유)",
                "비노조",
                "미정",
            ],
            "대수": [
                df4.iloc[4, 0],
                df4.iloc[4, 1],
                df4.iloc[4, 2],
                df4.iloc[4, 3],
                df4.iloc[4, 4],
                df4.iloc[4, 5],
            ],
            "점유율": [
                "-",
                f"{float(df4.iloc[5, 1])*100:.1f}%"
                if pd.notna(df4.iloc[5, 1])
                else "-",
                f"{float(df4.iloc[5, 2])*100:.1f}%"
                if pd.notna(df4.iloc[5, 2])
                else "-",
                f"{float(df4.iloc[5, 3])*100:.1f}%"
                if pd.notna(df4.iloc[5, 3])
                else "-",
                f"{float(df4.iloc[5, 4])*100:.1f}%"
                if pd.notna(df4.iloc[5, 4])
                else "-",
                f"{float(df4.iloc[5, 5])*100:.1f}%"
                if pd.notna(df4.iloc[5, 5])
                else "-",
            ],
        }
        st.table(pd.DataFrame(summary_data))
    except:
        pass

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
            st.write(
                f"• **한국노총:** {hanno}대 | **민주노총:** {minno}대 | **건설노조:** {gunsan}대"
            )
            st.write(f"• **비노조:** {non_union}대 | **미정:** {mijung}대")

            if contact_info:
                import re

                phone_match = re.search(
                    r"01[016789][-\s]?\d{3,4}[-\s]?\d{4}", contact_info
                )
                if phone_match:
                    phone_num = (
                        phone_match.group().replace("-", "").replace(" ", "")
                    )
                    st.markdown(
                        f"👤 **담당자:** {contact_info}<br>"
                        f'<a href="tel:{phone_num}" class="phone-btn">📞 담당자 바로 전화걸기</a>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.write(f"👤 **담당자:** {contact_info}")

# 5. 채용추이
elif menu == "📅 1~9월 채용 추이":
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

        with st.expander(f"🗓️ {m_name} 요약 보기"):
            if sum_row is not None:
                st.markdown(
                    f"**[합계]** 총대수: **{sum_row[3] if len(sum_row)>3 else '-'}**대 | "
                    f"한노: {sum_row[4]} | 민노: {sum_row[5]} | 기타: {sum_row[6]} | "
                    f"직원: {sum_row[7]} | 미정: {sum_row[8]}"
                )
            if pct_row is not None:

                def fmt_p(v):
                    try:
                        return f"{float(v)*100:.1f}%"
                    except:
                        return str(v) if pd.notna(v) else "0%"

                st.markdown(
                    f"**[비율]** "
                    f"한노: **{fmt_p(pct_row[4])}** | 민노: **{fmt_p(pct_row[5])}** | "
                    f"기타: {fmt_p(pct_row[6])} | 직원: {fmt_p(pct_row[7])} | "
                    f"미정: {fmt_p(pct_row[8])}"
                )

    st.markdown("---")
    st.subheader("🔍 상세 채용 현황 조회할 월 선택")
    selected_m = st.selectbox("월을 선택하세요", list(month_blocks.keys()))

    if selected_m:
        m_rows = month_blocks[selected_m]
        detail_list = []
        for r in m_rows:
            cell0 = str(r[0]).strip() if pd.notna(r[0]) else ""
            if (
                cell0
                and "임대사" not in cell0
                and "합계" not in cell0
                and "비율" not in cell0
            ):
                detail_list.append(
                    {
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
                    }
                )

        if detail_list:
            st.markdown(f"### 📍 {selected_m} 상세 현황")
            st.dataframe(
                pd.DataFrame(detail_list).fillna(""),
                use_container_width=True,
                hide_index=True,
            )
