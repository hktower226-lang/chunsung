import pandas as pd
import streamlit as st

# ==========================================
# 1. 페이지 기본 설정 및 모바일 CSS 스타일링
# ==========================================
st.set_page_config(
    page_title="현장 점유율 및 채용 현황", layout="wide", initial_sidebar_state="collapsed"
)

# 모바일용 가독성 증대 CSS (큰 글씨, 위아래 스크롤 레이아웃, 카드 색상 추가)
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

    /* 메인 화면 전용 카드 스타일 & 소속별 구분색 */
    .main-total-card {
        background: linear-gradient(135deg, #1E3A8A, #3B82F6);
        color: white;
        padding: 16px;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 18px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .main-stat-card {
        background-color: #FFFFFF;
        border-left: 6px solid #3B82F6;
        border-radius: 8px;
        padding: 12px 15px;
        margin-bottom: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06);
    }
    .border-hanno { border-left-color: #EF4444 !important; }   /* 빨강 */
    .border-minno { border-left-color: #3B82F6 !important; }   /* 파랑 */
    .border-seomoo { border-left-color: #10B981 !important; }  /* 초록 */
    .border-geunsan { border-left-color: #F59E0B !important; } /* 주황 */
    .border-jikwon { border-left-color: #8B5CF6 !important; }  /* 보라 */
    .border-mijeong { border-left-color: #6B7280 !important; } /* 회색 */
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

if len(sheet_names) > 0:
    s1_name = sheet_names[0]
else:
    s1_name = "1.26년 전체 소속별 점유율"

if len(sheet_names) > 1:
    s2_name = sheet_names[1]
else:
    s2_name = "2.각8개지부 현장점유율"

if len(sheet_names) > 2:
    s3_name = sheet_names[2]
else:
    s3_name = "3.타워사별 점유현황"

if len(sheet_names) > 3:
    s4_name = sheet_names[3]
else:
    s4_name = "4.반도체현장"

if len(sheet_names) > 4:
    s5_name = sheet_names[4]
else:
    s5_name = "5.1~9월채용추이"


# 퍼센트 포맷 변환 헬퍼 함수
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
# 3. 사이드바 메뉴 (요청 반영: 채용 비율 메뉴 선택)
# ==========================================
st.sidebar.title("📌 채용 비율 메뉴 선택")
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

# ------------------------------------------
# [메뉴 1] 메인: 1.26년 전체 소속별 점유율
# ------------------------------------------
if menu == "🏠 메인: 전체 소속별 점유율":
    st.title("📊 2026년 전체 소속별 점유율")
    df1 = sheets[s1_name].copy()

    try:
        tot_cnt = df1.iloc[1, 7]  # 209
        tot_pct = df1.iloc[1, 8]  # 100%

        stat_items = [
            {
                "title": "한국노총 (한노)",
                "cnt": df1.iloc[1, 1],
                "pct": fmt_pct(df1.iloc[2, 1]),
                "class": "border-hanno",
                "color": "#EF4444",
            },
            {
                "title": "민주노총 (민노)",
                "cnt": df1.iloc[1, 2],
                "pct": fmt_pct(df1.iloc[2, 2]),
                "class": "border-minno",
                "color": "#3B82F6",
            },
            {
                "title": "섬유노조",
                "cnt": df1.iloc[1, 3],
                "pct": fmt_pct(df1.iloc[2, 3]),
                "class": "border-seomoo",
                "color": "#10B981",
            },
            {
                "title": "건설산업 (건산)",
                "cnt": df1.iloc[1, 4],
                "pct": fmt_pct(df1.iloc[2, 4]),
                "class": "border-geunsan",
                "color": "#F59E0B",
            },
            {
                "title": "직원",
                "cnt": df1.iloc[1, 5],
                "pct": fmt_pct(df1.iloc[2, 5]),
                "class": "border-jikwon",
                "color": "#8B5CF6",
            },
            {
                "title": "미정",
                "cnt": df1.iloc[1, 6],
                "pct": fmt_pct(df1.iloc[2, 6]),
                "class": "border-mijeong",
                "color": "#6B7280",
            },
        ]

        # 1. 상단 총 대수 강조 카드
        st.markdown(
            f"""
        <div class="main-total-card">
            <h3 style="margin:0; color:white; font-size: 1.2rem !important;">🏆 전체 현장 총 대수</h3>
            <h1 style="margin:4px 0 0 0; color:white; font-size: 2.0rem !important;">{tot_cnt}대 <span style="font-size: 1.2rem;">({tot_pct})</span></h1>
        </div>
        """,
            unsafe_allow_html=True,
        )

        st.markdown("<h4 style='font-size: 1.2rem; font-weight: bold; color: #1E40AF; margin-bottom: 10px;'>👇 소속별 점유 현황 (위아래 세로 보기)</h4>", unsafe_allow_html=True)

        # 2. 위아래로 내리는 세로 배치 카드
        for item in stat_items:
            st.markdown(
                f"""
            <div class="main-stat-card {item['class']}">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-size: 1.05rem; font-weight:bold; color:{item['color']};">{item['title']}</span>
                    <span style="font-size: 1.25rem; font-weight:bold; color:#111827;">{item['pct']}</span>
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
# [메뉴 2] 2.각 8개지부 현장점유율 (요청 반영)
# ------------------------------------------
elif menu == "🏢 각 8개지부 현장 점유율":
    st.title("🏢 지부별 현장 점유율")

    df2 = sheets[s2_name].copy()

    # 데이터 파싱: 지부 / 소계 / 퍼센티지 / 현장 데이터 분리
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

    # 각 지부별 데이터 구조화
    branch_data = {}
    current_branch = "북부지부"

    for r in rows[1:]:  # 헤더 제외
        branch_col = str(r[0]).strip() if pd.notna(r[0]) else ""

        # 지부 변경 감지
        for b in ["북부", "남부", "남서", "동부", "서부", "북서", "용인", "중부"]:
            if branch_col == b or branch_col == b + "지부":
                current_branch = b + "지부"
                break

        if current_branch not in branch_data:
            branch_data[current_branch] = {
                "summary": None,
                "percent": None,
                "sites": [],
            }

        if "퍼센테이지" in branch_col or "소계퍼센테이지" in branch_col:
            branch_data[current_branch]["percent"] = r
        elif "소계" in branch_col:
            branch_data[current_branch]["summary"] = r
        elif pd.notna(r[1]) and str(r[1]).strip() != "현장명":  # 현장 데이터
            branch_data[current_branch]["sites"].append(r)

    st.subheader("📌 각 8개 지부 현장 점유율")

    for b_name in branches:
        if b_name in branch_data:
            b_info = branch_data[b_name]
            sum_row = b_info["summary"]
            pct_row = b_info["percent"]
            sites_list = b_info["sites"]

            # 외부 제목 만들기 (요청대로 소속별 색상 지정 및 한 줄 표출)
            if sum_row is not None and pct_row is not None:
                hanno_cnt, hanno_pct = sum_row[4], fmt_pct(pct_row[4])
                minno_cnt, minno_pct = sum_row[5], fmt_pct(pct_row[5])
                seomoo_cnt, seomoo_pct = sum_row[6], fmt_pct(pct_row[6])
                geunsan_cnt, geunsan_pct = sum_row[7], fmt_pct(pct_row[7])
                jikwon_cnt, jikwon_pct = sum_row[8], fmt_pct(pct_row[8])
                mijeong_cnt, mijeong_pct = sum_row[9], fmt_pct(pct_row[9])
                total_cnt = sum_row[10] if len(sum_row) > 10 else "-"

                expander_label = (
                    f"{b_name}📊 "
                    f":red[한노 {hanno_cnt}({hanno_pct})] | "
                    f":blue[민노 {minno_cnt}({minno_pct})] | "
                    f":green[섬유 {seomoo_cnt}({seomoo_pct})] | "
                    f":orange[건산 {geunsan_cnt}({geunsan_pct})] | "
                    f":violet[직원 {jikwon_cnt}({jikwon_pct})] | "
                    f":gray[미정 {mijeong_cnt}({mijeong_pct})] | "
                    f"**총대수 {total_cnt}(100%)**"
                )
            elif sum_row is not None:
                total_cnt = sum_row[10] if len(sum_row) > 10 else "-"
                expander_label = f"{b_name}📊 **총대수 {total_cnt}대**"
            else:
                expander_label = f"{b_name}"

            # 펼치기 항목
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
# [메뉴 3] 3.타워사별 점유현황
# ------------------------------------------
elif menu == "🏗️ 타워사별 점유 현황":
    st.title("🏗️ 임대사(타워사)별 점유 현황")

    df3 = sheets[s3_name].copy()

    # 데이터 정돈
    if "타워사" in df3.iloc[0].values:
        df3.columns = df3.iloc[0]
        df3 = df3[1:].reset_index(drop=True)

    df3 = df3.dropna(subset=["타워사"]).fillna(0)

    # 타워사 검색 기능
    search_keyword = st.text_input(
        "🔍 임대사(타워사) 명칭 검색", "", placeholder="예: 비엠케이, 국영 등"
    )

    if search_keyword:
        filtered_df = df3[
            df3["타워사"].astype(str).str.contains(search_keyword, case=False)
        ]
    else:
        filtered_df = df3

    # 한노점유율 퍼센트 변환
    if "한노점유율" in filtered_df.columns:
        filtered_df_disp = filtered_df.copy()
        filtered_df_disp["한노점유율"] = filtered_df_disp["한노점유율"].apply(
            fmt_pct
        )
    else:
        filtered_df_disp = filtered_df

    st.markdown("### 📋 타워사 점유 현황 목록")
    st.dataframe(filtered_df_disp, use_container_width=True, hide_index=True)


# ------------------------------------------
# [메뉴 4] 4.반도체현장
# ------------------------------------------
elif menu == "🏭 반도체 현장 현황":
    st.title("🏭 반도체 현장 타워크레인 현황")

    df4 = sheets[s4_name].copy()

    # 1. 전체 대수 요약 정보 찾기
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
                fmt_pct(df4.iloc[5, 1]),
                fmt_pct(df4.iloc[5, 2]),
                fmt_pct(df4.iloc[5, 3]),
                fmt_pct(df4.iloc[5, 4]),
                fmt_pct(df4.iloc[5, 5]),
            ],
        }
        st.table(pd.DataFrame(summary_data))
    except Exception:
        st.info("상단 요약 데이터 표시 중")

    st.markdown("---")
    st.subheader("📍 각 현장별 상세 현황 및 담당자")

    # 현장별 상세 테이블 추출 (8번 행부터)
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

            # 담당자 전화번호 연결 버튼 생성
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


# ------------------------------------------
# [메뉴 5] 5.1~9월채용추이
# ------------------------------------------
elif menu == "📅 1~9월 채용 추이":
    st.title("📅 월별 채용 현황 및 추이")

    df5 = sheets[s5_name].copy()

    # 월별 블록 분리 파싱
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

    # 1. 월별 합계 및 비율 한눈에 보기
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

    # 2. 월 선택 후 상세 현황 보기
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
        else:
            st.info("해당 월의 상세 데이터가 없습니다.")
