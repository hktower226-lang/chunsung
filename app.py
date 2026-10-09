import re
import pandas as pd
import streamlit as st

# ==========================================
# 1. 모바일 최적화 레이아웃 및 디자인 CSS
# ==========================================
st.set_page_config(
    page_title="현장 점유율 어플", layout="wide", initial_sidebar_state="expanded"
)

st.markdown(
    """
    <style>
    /* 전체 글씨 크기 확대 */
    html, body, [class*="css"] { font-size: 19px !important; }
    
    /* 헤더 및 타이틀 */
    h1 { font-size: 2.1rem !important; font-weight: bold !important; color: #1E3A8A; }
    h2 { font-size: 1.7rem !important; font-weight: bold !important; color: #1E40AF; }
    h3 { font-size: 1.4rem !important; font-weight: bold !important; }
    
    /* 모바일 세로 스크롤 레이아웃 */
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 3rem !important;
        padding-left: 0.6rem !important;
        padding-right: 0.6rem !important;
    }
    
    /* 버튼 및 카드 디자인 */
    .stSelectbox label, .stRadio label { font-size: 1.2rem !important; font-weight: bold !important; }
    
    /* 전화걸기 버튼 스타일 */
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
        width: 100%;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }
    
    /* 카드형 박스 */
    .card-box {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 12px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

EXCEL_FILE = "채용비율.xlsx"


# ==========================================
# 2. 엑셀 데이터 로드 (헤더 없는 원본 파싱)
# ==========================================
@st.cache_data(ttl=5)
def load_sheets():
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        # header=None으로 읽어 원본 데이터 유실 방지
        sheets = {
            sheet: pd.read_excel(xls, sheet, header=None)
            for sheet in xls.sheet_names
        }
        return sheets
    except Exception as e:
        st.error(
            f"❌ 엑셀 파일('{EXCEL_FILE}')을 읽는 중 에러가 발생했습니다: {e}"
        )
        return None


sheets = load_sheets()

if not sheets:
    st.stop()

sheet_keys = list(sheets.keys())

# ==========================================
# 3. 사이드바 어플 메뉴 (모바일 메뉴)
# ==========================================
st.sidebar.title("📱 어플 메뉴")
menu = st.sidebar.radio(
    "원하시는 화면을 선택하세요",
    [
        "1️⃣ 메인: 전체 소속별 점유율",
        "2️⃣ 8개 지부 현장 점유율",
        "3️⃣ 타워사별 점유 현황",
        "4️⃣ 반도체 현장 파악",
        "5️⃣ 1~9월 채용 현황 및 추이",
    ],
)

# 퍼센트 변환 함수
def to_pct_str(val):
    try:
        if pd.isna(val) or val == "" or val == "-":
            return "0%"
        f_val = float(val)
        if f_val <= 1.0:
            return f"{f_val * 100:.1f}%"
        return f"{f_val:.1f}%"
    except:
        return str(val)


# ==========================================
# 4. 메뉴별 화면 구현
# ==========================================

# ------------------------------------------
# [메뉴 1] 어플 켰을 때 첫번째 시트 (메인)
# ------------------------------------------
if menu == "1️⃣ 메인: 전체 소속별 점유율":
    st.title("📊 전체 소속별 점유율 (메인)")

    raw_df1 = sheets[sheet_keys[0]].fillna("")

    # 메인 시트 예쁘게 테이블로 표시
    st.markdown("### 💡 전체 현장 통계 개요")

    # 헤더와 데이터 분리
    header1 = raw_df1.iloc[1].tolist()
    data1 = raw_df1.iloc[2:].copy()
    data1.columns = header1

    # 비율 퍼센트 처리
    if "비율" in data1.columns:
        data1["비율"] = data1["비율"].apply(to_pct_str)

    st.dataframe(data1, use_container_width=True, hide_index=True)


# ------------------------------------------
# [메뉴 2] 2번째 시트: 각 8개지부 현장점유율
# ------------------------------------------
elif menu == "2️⃣ 8개 지부 현장 점유율":
    st.title("🏢 지부별 현장 점유율")

    raw_df2 = sheets[sheet_keys[1]]

    # 1. 지부별 데이터 분리 파싱
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
    branch_map = {
        "북부": "북부지부",
        "남부": "남부지부",
        "남서": "남서지부",
        "동부": "동부지부",
        "서부": "서부지부",
        "북서": "북서지부",
        "용인": "용인지부",
        "중부": "중부지부",
    }

    b_data = {b: {"summary": None, "pct": None, "sites": []} for b in branches}

    curr_b = "북부지부"

    for idx, row in raw_df2.iterrows():
        if idx == 0:
            continue  # 컬럼명 행 제외

        c0 = str(row[0]).strip() if pd.notna(row[0]) else ""

        # 지부 이름 감지
        for short_k, full_k in branch_map.items():
            if c0 == short_k:
                curr_b = full_k
                break

        if "소계" in c0 and "퍼센테이지" not in c0:
            b_data[curr_b]["summary"] = row
        elif "퍼센테이지" in c0:
            b_data[curr_b]["pct"] = row
        elif pd.notna(row[1]) and str(row[1]).strip() != "현장명":
            b_data[curr_b]["sites"].append(row)

    # 2. 먼저 각 지부 소계 및 퍼센티지 한눈에 쫙 출력
    st.subheader("📌 1. 각 지부 소계 및 퍼센티지 요약")

    for idx_b, b_name in enumerate(branches, 1):
        info = b_data[b_name]
        s_row = info["summary"]
        p_row = info["pct"]

        with st.container():
            st.markdown(f"#### {idx_b}. {b_name} 소계 및 퍼센티지")

            total_cnt = s_row[10] if s_row is not None and len(s_row) > 10 else 0
            hanno = s_row[4] if s_row is not None else 0
            minno = s_row[5] if s_row is not None else 0
            seomoo = s_row[6] if s_row is not None else 0
            geunsan = s_row[7] if s_row is not None else 0
            jikwon = s_row[8] if s_row is not None else 0
            mijeong = s_row[9] if s_row is not None else 0

            hanno_p = to_pct_str(p_row[4]) if p_row is not None else "0%"
            minno_p = to_pct_str(p_row[5]) if p_row is not None else "0%"
            seomoo_p = to_pct_str(p_row[6]) if p_row is not None else "0%"
            geunsan_p = to_pct_str(p_row[7]) if p_row is not None else "0%"
            jikwon_p = to_pct_str(p_row[8]) if p_row is not None else "0%"
            mijeong_p = to_pct_str(p_row[9]) if p_row is not None else "0%"

            st.markdown(
                f"""
            <div class="card-box">
                <b>[소계]</b> 총대수: <b>{total_cnt}대</b> | 한노: {hanno} | 민노: {minno} | 섬유: {seomoo} | 건산: {geunsan} | 직원: {jikwon} | 미정: {mijeong}<br>
                <b>[비율]</b> 한노: <b>{hanno_p}</b> | 민노: <b>{minno_p}</b> | 섬유: {seomoo_p} | 건산: {geunsan_p} | 직원: {jikwon_p} | 미정: {mijeong_p}
            </div>
            """,
                unsafe_allow_html=True,
            )

    st.markdown("---")

    # 3. 클릭 시 해당 지부 상세 정보 출력 (버튼 / 선택 박스)
    st.subheader("📌 2. 지부 클릭 시 상세 현황 보기")
    selected_branch = st.selectbox(
        "상세 내용을 확인할 지부를 선택하세요", branches
    )

    if selected_branch and b_data[selected_branch]["sites"]:
        sites_rows = b_data[selected_branch]["sites"]
        site_df = pd.DataFrame(sites_rows)

        # 표시할 항목: 현장명, 타워회사, 한노, 민노, 섬유, 건산, 직원, 미정, 총대수
        res_df = pd.DataFrame(
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

        st.markdown(f"### 📍 [{selected_branch}] 상세 현장 목록")
        st.dataframe(res_df, use_container_width=True, hide_index=True)


# ------------------------------------------
# [메뉴 3] 3번째 시트: 타워사별 점유현황
# ------------------------------------------
elif menu == "3️⃣ 타워사별 점유 현황":
    st.title("🏗️ 임대사 카테고리 검색 및 점유현황")

    raw_df3 = sheets[sheet_keys[2]].fillna(0)

    # 헤더 정리
    cols = raw_df3.iloc[1].tolist()
    df3 = raw_df3.iloc[2:].copy()
    df3.columns = cols

    # 임대사 카테고리 검색 기능
    search_q = st.text_input(
        "🔍 임대사(타워사) 검색", "", placeholder="타워사명을 입력하세요 (예: 비엠케이)"
    )

    if search_q:
        df3_filtered = df3[
            df3["타워사"].astype(str).str.contains(search_q, case=False)
        ]
    else:
        df3_filtered = df3

    # 한노점유율 퍼센트 변환
    if "한노점유율" in df3_filtered.columns:
        df3_filtered["한노점유율"] = df3_filtered["한노점유율"].apply(
            to_pct_str
        )

    st.markdown("### 📋 임대사 점유현황 목록")
    st.dataframe(df3_filtered, use_container_width=True, hide_index=True)


# ------------------------------------------
# [메뉴 4] 4번째 시트: 반도체현장 (전화 연결)
# ------------------------------------------
elif menu == "4️⃣ 반도체 현장 파악":
    st.title("🏭 반도체 현장 타워크레인 현황")

    raw_df4 = sheets[sheet_keys[3]]

    # 1. 요약 현황
    st.subheader("📊 반도체 현장 전체 대수 비교")
    try:
        sum_table = pd.DataFrame(
            {
                "구분": [
                    "전체대수",
                    "한국노총",
                    "민주노총",
                    "건설노조(건산,섬유)",
                    "비노조",
                    "미정",
                ],
                "대수": [
                    raw_df4.iloc[5, 0],
                    raw_df4.iloc[5, 1],
                    raw_df4.iloc[5, 2],
                    raw_df4.iloc[5, 3],
                    raw_df4.iloc[5, 4],
                    raw_df4.iloc[5, 5],
                ],
                "점유율": [
                    "-",
                    to_pct_str(raw_df4.iloc[6, 1]),
                    to_pct_str(raw_df4.iloc[6, 2]),
                    to_pct_str(raw_df4.iloc[6, 3]),
                    to_pct_str(raw_df4.iloc[6, 4]),
                    to_pct_str(raw_df4.iloc[6, 5]),
                ],
            }
        )
        st.table(sum_table)
    except:
        pass

    st.markdown("---")
    st.subheader("📍 각 현장별 상세 내용 및 담당자 통화")

    # 현장 데이터 파싱 (10번 행부터)
    sites_rows = raw_df4.iloc[10:].dropna(how="all").copy()

    for idx, row in sites_rows.iterrows():
        site_name = str(row[1]) if pd.notna(row[1]) else ""
        if not site_name or site_name == "nan" or site_name == "현장명":
            continue

        total_cnt = row[2] if pd.notna(row[2]) else 0
        hanno = row[3] if pd.notna(row[3]) else 0
        minno = row[4] if pd.notna(row[4]) else 0
        gunsan = row[5] if pd.notna(row[5]) else 0
        non_union = row[6] if pd.notna(row[6]) else 0
        mijung = row[7] if pd.notna(row[7]) else 0
        contact_str = str(row[8]) if pd.notna(row[8]) else ""

        with st.expander(f"🏢 {site_name} (총 {total_cnt}대)", expanded=True):
            st.write(
                f"• **한국노총:** {hanno}대 | **민주노총:** {minno}대 | **건설노조:** {gunsan}대"
            )
            st.write(f"• **비노조:** {non_union}대 | **미정:** {mijung}대")

            # 전화번호 자동 감지 및 즉시 통화 연결 버튼
            if contact_str:
                phone_match = re.search(
                    r"01[016789][-\s]?\d{3,4}[-\s]?\d{4}", contact_str
                )
                if phone_match:
                    clean_phone = (
                        phone_match.group().replace("-", "").replace(" ", "")
                    )
                    st.markdown(
                        f"👤 **담당자:** {contact_str}<br>"
                        f'<a href="tel:{clean_phone}" class="phone-btn">📞 담당자에게 전화걸기 ({clean_phone})</a>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.write(f"👤 **담당자 비고:** {contact_str}")


# ------------------------------------------
# [메뉴 5] 5번째 시트: 1~9월 채용 추이
# ------------------------------------------
elif menu == "5️⃣ 1~9월 채용 현황 및 추이":
    st.title("📅 월별 채용 현황 및 비율 추이")

    raw_df5 = sheets[sheet_keys[4]]

    # 월별 블록 파싱
    months_data = {}
    curr_m = None
    curr_rows = []

    for idx, row in raw_df5.iterrows():
        c0 = str(row[0]).strip() if pd.notna(row[0]) else ""

        if "채용 현황" in c0 or "채용현황" in c0:
            if curr_m and curr_rows:
                months_data[curr_m] = curr_rows
            curr_m = c0
            curr_rows = []
        elif curr_m:
            curr_rows.append(row)

    if curr_m and curr_rows:
        months_data[curr_m] = curr_rows

    # 1. 1월~9월 채용현황 합계 및 비율 먼저 쫙 출력
    st.subheader("📌 1. 월별 채용현황 합계 및 비율 요약")

    for m_title, rows_list in months_data.items():
        sum_r = None
        pct_r = None

        for r in rows_list:
            r0 = str(r[0]).strip() if pd.notna(r[0]) else ""
            if "합계" in r0:
                sum_r = r
            elif "비율" in r0:
                pct_r = r

        st.markdown(f"#### 🗓️ {m_title}")

        total_cnt = sum_r[3] if sum_r is not None and len(sum_r) > 3 else 0
        hanno = sum_r[4] if sum_r is not None else 0
        minno = sum_r[5] if sum_r is not None else 0
        etc = sum_r[6] if sum_r is not None else 0
        jikwon = sum_r[7] if sum_r is not None else 0
        mijeong = sum_r[8] if sum_r is not None else 0

        hanno_p = to_pct_str(pct_r[4]) if pct_r is not None else "0%"
        minno_p = to_pct_str(pct_r[5]) if pct_r is not None else "0%"
        etc_p = to_pct_str(pct_r[6]) if pct_r is not None else "0%"
        jikwon_p = to_pct_str(pct_r[7]) if pct_r is not None else "0%"
        mijeong_p = to_pct_str(pct_r[8]) if pct_r is not None else "0%"

        st.markdown(
            f"""
        <div class="card-box">
            <b>[합계]</b> 총대수: <b>{total_cnt}대</b> | 한노: {hanno} | 민노: {minno} | 기타: {etc} | 직원: {jikwon} | 미정: {mijeong}<br>
            <b>[비율]</b> 한노: <b>{hanno_p}</b> | 민노: <b>{minno_p}</b> | 기타: {etc_p} | 직원: {jikwon_p} | 미정: {mijeong_p}
        </div>
        """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # 2. 클릭 시 해당 월 상세 수치(임대사, 원청사, 현장명, 총대수, 수치들) 출력
    st.subheader("📌 2. 클릭하여 상세 현황 수치 보기")
    selected_month = st.selectbox(
        "상세 수치를 볼 월을 선택하세요", list(months_data.keys())
    )

    if selected_month:
        m_rows_list = months_data[selected_month]

        detail_records = []
        for r in m_rows_list:
            r0 = str(r[0]).strip() if pd.notna(r[0]) else ""
            if r0 and "임대사" not in r0 and "합계" not in r0 and "비율" not in r0:
                detail_records.append(
                    {
                        "임대사": r[0],
                        "원청사": r[1],
                        "현장명": r[2],
                        "총대수": r[3],
                        "한노": r[4],
                        "민노": r[5],
                        "기타(건산,섬유)": r[6],
                        "직원": r[7],
                        "미정": r[8],
                        "비고": r[9] if len(r) > 9 else "",
                    }
                )

        if detail_records:
            st.markdown(f"### 📍 [{selected_month}] 상세 채용 현황")
            st.dataframe(
                pd.DataFrame(detail_records).fillna(""),
                use_container_width=True,
                hide_index=True,
            )
