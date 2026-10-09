import pandas as pd
import streamlit as st

# ==========================================
# 1. 페이지 기본 설정 및 모바일 CSS 스타일링
# ==========================================
st.set_page_config(
    page_title="현장 점유율 및 채용 현황", layout="wide", initial_sidebar_state="collapsed"
)

# 모바일용 가독성 증대 CSS (큰 글씨, 위아래 스크롤 레이아웃, 버튼 강조)
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
s1_name = sheet_names[0] if len(sheet_names) > 0 else "1.전체소속별점유율"
s2_name = sheet_names[1] if len(sheet_names) > 1 else "2.각8개지부현장점유율"
s3_name = sheet_names[2] if len(sheet_names) > 2 else "3.타워사별점유현황"
s4_name = sheet_names[3] if len(sheet_names) > 3 else "4.반도체현장"
s5_name = sheet_names[4] if len(sheet_names) > 4 else "5.1~9월채용추이"

# ==========================================
# 3. 사이드바 메뉴 (모바일 상단/좌측 메뉴)
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

# ------------------------------------------
# [메뉴 1] 메인: 1.26년 전체 소속별 점유율
# ------------------------------------------
if menu == "🏠 메인: 전체 소속별 점유율":
    st.title("📊 2026년 전체 소속별 점유율")
    df1 = sheets[s1_name].copy()

    # 데이터 정돈 및 출력
    st.write("### 💡 전체 소속별 통계 개요")
    st.dataframe(df1.fillna(""), use_container_width=True, hide_index=True)


# ------------------------------------------
# [메뉴 2] 2.각 8개지부 현장점유율
# ------------------------------------------
elif menu == "🏢 각 8개지부 현장 점유율":
    st.title("🏢 지부별 현장 점유율")

    df2 = sheets[s2_name].copy()

    # 데이터 파싱: 지부 / 소계 / 현장 데이터 분리
    # 컬럼 정의: [지부, 현장명, 타워회사, 특이사항, 한노, 민노, 섬유, 건산, 직원, 미정, 총대수, 비고]
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
    current_branch = "북부"

    for r in rows[1:]:  # 헤더 제외
        branch_col = str(r[0]).strip() if pd.notna(r[0]) else ""

        # 지부 변경 감지
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
        elif pd.notna(r[1]) and str(r[1]).strip() != "현장명":  # 현장 데이터
            branch_data[current_branch]["sites"].append(r)

    # 1. 지부별 소계 및 퍼센티지 한눈에 보기
    st.subheader("📌 지부별 소계 및 퍼센티지")

    for b_name in branches:
        if b_name in branch_data:
            b_info = branch_data[b_name]
            sum_row = b_info["summary"]
            pct_row = b_info["percent"]

            with st.expander(f"🔹 {b_name} 요약 보기 (클릭하여 열기/접기)"):
                if sum_row is not None:
                    st.markdown(
                        f"**[합계]** 총대수: **{sum_row[10] if len(sum_row)>10 else '-'}**대 | "
                        f"한노: {sum_row[4]} | 민노: {sum_row[5]} | 섬유: {sum_row[6]} | "
                        f"건산: {sum_row[7]} | 직원: {sum_row[8]} | 미정: {sum_row[9]}"
                    )
                if pct_row is not None:
                    # 퍼센트 변환 표시
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

    # 2. 지부 선택 후 상세 현황 보기
    st.subheader("🔍 상세 현황 조회할 지부 선택")
    selected_b = st.selectbox("지부를 선택하세요", branches)

    if selected_b in branch_data and branch_data[selected_b]["sites"]:
        sites_list = branch_data[selected_b]["sites"]
        site_df = pd.DataFrame(sites_list)

        # 필요한 컬럼 추출 [현장명, 타워회사, 한노, 민노, 섬유, 건산, 직원, 미정, 총대수]
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
    else:
        st.info("해당 지부의 상세 현장 데이터가 없습니다.")


# ------------------------------------------
# [메뉴 3] 3.타워사별 점유현황
# ------------------------------------------
elif menu == "🏗️ 타워사별 점유 현황":
    st.title("🏗️ 임대사(타워사)별 점유 현황")

    df3 = sheets[s3_name].copy()

    # 데이터 정돈
    # 0번 행이 헤더인 경우 처리
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

        def to_pct(val):
            try:
                return f"{float(val)*100:.1f}%"
            except:
                return str(val)

        filtered_df_disp = filtered_df.copy()
        filtered_df_disp["한노점유율"] = filtered_df_disp["한노점유율"].apply(to_pct)
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
        # 전체 대수 요약 테이블 추출
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
        st.info("상단 요약 데이터
