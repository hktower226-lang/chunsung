import os
import pandas as pd
import streamlit as st

# 1. 페이지 설정 (모바일 최적화 및 넓은 화면 사용)
st.set_page_config(
    page_title="경기지역본부 현장 점유율 및 통계 조회",
    page_icon="🏗️",
    layout="centered",  # 모바일에서 보기 편하도록 가운데 정렬 레이아웃
    initial_sidebar_state="collapsed",
)

# 2. 모바일 가독성을 위한 커스텀 CSS 스타일 적용
st.markdown(
    """
    <style>
    /* 전체 폰트 및 여백 조정 */
    .main {
        background-color: #f8f9fa;
    }
    /* 타이틀 스타일 */
    h1 {
        font-size: 1.5rem !important;
        color: #1f2937;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    /* 설명글 스타일 */
    .subtitle {
        font-size: 0.85rem;
        color: #6b7280;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    /* 카드 디자인 (모바일에서 박스형태로 정보 강조) */
    .metric-card {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        text-align: center;
        margin-bottom: 10px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# 3. 타이틀 영역
st.markdown("🏗️ 경기지역본부 현장 점유율 및 통계 조회", unsafe_allow_html=True)
st.markdown(
    '<p class="subtitle">임직원 및 타워사 점유율 현황 조회 전용 프로그램 (조회만 가능)</p>',
    unsafe_allow_html=True,
)


# 4. 엑셀 파일 자동 탐색 및 불러오기 함수
@st.cache_data(ttl=5)  # 데이터가 갱신되면 빠르게 반영되도록 캐시 설정
def load_excel_data():
    excel_files = [
        f
        for f in os.listdir(".")
        if f.endswith((".xlsx", ".xls")) and not f.startswith("~$")
    ]
    if not excel_files:
        return None, None, "폴더에 엑셀 파일이 없습니다."

    target_file = excel_files[0]

    try:
        # 첫 번째 시트 (현장 목록 데이터)
        df1 = pd.read_excel(target_file, sheet_name=0)
        df1.columns = df1.columns.astype(str).str.strip()

        # 두 번째 시트 (지부별 점유율 데이터) 존재 여부 확인
        xl = pd.ExcelFile(target_file)
        df2 = None
        if len(xl.sheet_names) > 1:
            df2 = pd.read_excel(target_file, sheet_name=1)
            df2.columns = df2.columns.astype(str).str.strip()

        return df1, df2, None
    except Exception as e:
        return None, None, str(e)


df, df_share, err_msg = load_excel_data()

# 엑셀 파일 로드 실패 시 에러 메시지 출력
if err_msg or df is None:
    st.error(f"엑셀 파일을 읽지 못했습니다. 파일을 확인해주세요: {err_msg}")
else:
    # 5. 탭 구성 (1번 탭: 현장 조회, 2번 탭: 지부별 점유율 통계)
    tab1, tab2 = st.tabs(["📋 현장 목록 조회", "📊 지부별 점유율 통계"])

    # --- [탭 1] 현장 목록 조회 화면 (모바일 친화적 검색) ---
    with tab1:
        st.markdown("### 🔍 현장 검색")

        # 검색창 입력 (모바일 키보드 사용 편리)
        search_keyword = st.text_input(
            "현장명, 지부 또는 타워사 검색",
            placeholder="검색어를 입력하세요 (예: 수원, A타워 등)",
        )

        filtered_df = df.copy()
        if search_keyword:
            # 전체 컬럼에서 검색어 포함된 행 필터링
            mask = (
                filtered_df.astype(str)
                .apply(
                    lambda x: x.str.contains(search_keyword, case=False, na=False)
                )
                .any(axis=1)
            )
            filtered_df = filtered_df[mask]

        st.markdown(
            f"<p style='font-size: 0.8rem; color: #555;'>검색 결과: <b>{len(filtered_df)}</b>건</p>",
            unsafe_allow_html=True,
        )

        # 데이터 테이블 출력 (모바일 스크롤 최적화)
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)

    # --- [탭 2] 지부별 점유율 통계 화면 ---
    with tab2:
        st.markdown("### 📊 지부별 점유율 현황")

        if df_share is not None and not df_share.empty:
            # 깔끔하게 테이블 및 차트 형태로 시각화 제공
            st.dataframe(df_share, use_container_width=True, hide_index=True)

            # 만약 숫자 데이터가 있다면 간단한 바 차트 추가 제공 (선택사항)
            numeric_cols = df_share.select_dtypes(
                include=["number"]
            ).columns.tolist()
            if len(numeric_cols) > 0 and len(df_share) > 0:
                st.markdown("---")
                st.markdown("📈 **시각화 통계 그래프**")
                # 첫 번째 문자열/범주형 열을 인덱스로 잡고 바 차트 표시 시도
                try:
                    string_cols = df_share.select_dtypes(
                        include=["object"]
                    ).columns.tolist()
                    if string_cols:
                        chart_data = df_share.set_index(string_cols[0])[
                            numeric_cols[0]
                        ]
                        st.bar_chart(chart_data)
                except Exception:
                    pass
        else:
            st.info(
                "엑셀 파일의 두 번째 시트에 지부별 점유율 데이터가 비어있거나 감지되지 않았습니다."
            )

    # 하단 안내 문구
    st.markdown("---")
    st.markdown(
        "<p style='text-align: center; font-size: 0.75rem; color: #9ca3af;'>© Gyeonggi Regional Headquarters - Mobile Occupancy Dashboard</p>",
        unsafe_allow_html=True,
    )
