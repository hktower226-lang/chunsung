import os
import pandas as pd
import streamlit as st

# 1. 페이지 설정 (모바일 최적화 및 넓은 화면 사용)
st.set_page_config(
    page_title="경기지역본부 현장 점유율 및 통계 조회",
    page_icon="🏗️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# 2. 모바일 가독성 및 디자인을 위한 고급 CSS 스타일 적용
st.markdown(
    """
    <style>
    /* 전체 배경 및 폰트 부드럽게 */
    .main {
        background-color: #f3f4f6;
    }
    /* 타이틀 스타일 */
    h1 {
        font-size: 1.4rem !important;
        color: #111827;
        text-align: center;
        margin-bottom: 0.2rem;
        font-weight: 700;
    }
    .subtitle {
        font-size: 0.8rem;
        color: #4b5563;
        text-align: center;
        margin-bottom: 1.2rem;
    }
    /* 데이터 테이블/카드 여백 및 테두리 깔끔하게 */
    div.stDataFrame {
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    /* 탭 디자인 강조 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        justify-content: center;
    }
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        background-color: #e5e7eb;
        border-radius: 6px;
        color: #374151;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2563eb !important;
        color: white !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# 3. 타이틀 영역
st.markdown("🏗️ 경기지역본부 현장 점유율 및 통계", unsafe_allow_html=True)
st.markdown(
    '<p class="subtitle">모바일 최적화 통합 조회 시스템 (조회 전용)</p>',
    unsafe_allow_html=True,
)


# 4. 엑셀 데이터 정밀 로드 및 퍼센트(%) 서식 복구 함수
@st.cache_data(ttl=5)
def load_and_clean_excel():
    excel_files = [
        f
        for f in os.listdir(".")
        if f.endswith((".xlsx", ".xls")) and not f.startswith("~$")
    ]
    if not excel_files:
        return None, None, None, "폴더에 엑셀 파일이 없습니다."

    target_file = excel_files[0]

    try:
        xl = pd.ExcelFile(target_file)
        sheet_names = xl.sheet_names

        # 첫 번째 시트 (현장 목록)
        df1 = pd.read_excel(target_file, sheet_name=0)
        df1.columns = df1.columns.astype(str).str.strip()
        df1 = df1.fillna("-")  # 빈칸이나 None을 깔끔한 '-'로 치환

        # 두 번째 시트 이상이 존재할 경우 모두 탐색 (지부별 / 타워사별 분리 또는 통합 처리)
        df2 = None
        df3 = None

        if len(sheet_names) > 1:
            df2 = pd.read_excel(target_file, sheet_name=1)
            df2.columns = df2.columns.astype(str).str.strip()
            df2 = df2.fillna("-")

            # 만약 세 번째 시트도 있다면 추가로 로드
            if len(sheet_names) > 2:
                df3 = pd.read_excel(target_file, sheet_name=2)
                df3.columns = df3.columns.astype(str).str.strip()
                df3 = df3.fillna("-")

        # 소수점 데이터(예: 0.24)를 퍼센트 형태("24.0%")로 보기 좋게 자동 변환해 주는 로직
        for df_target in [df1, df2, df3]:
            if df_target is not None:
                for col in df_target.columns:
                    # 숫자들이 소수점 형태(0과 1 사이)로 되어 있는 경우 퍼센트로 변환
                    if pd.api.types.is_numeric_dtype(df_target[col]):
                        # 값이 모두 0 이상 1 이레이고 정수가 아닌 경우에만 비율로 판단
                        if (
                            df_target[col].min() >= 0
                            and df_target[col].max() <= 1
                            and not (df_target[col] % 1 == 0).all()
                        ):
                            df_target[col] = (df_target[col] * 100).round(
                                1
                            ).astype(str) + "%"

        return df1, df2, df3, None
    except Exception as e:
        return None, None, None, str(e)


df_main, df_sub1, df_sub2, err_msg = load_and_clean_excel()

if err_msg or df_main is None:
    st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {err_msg}")
else:
    # 5. 탭 구성 (시트 개수나 내용에 맞게 동적 구성)
    # 시트가 3개 이상이거나 타워사별 데이터가 별도로 있다면 탭을 유연하게 구성
    tab_names = ["📋 전체 현장 조회", "📊 지부별 점유율"]
    if df_sub2 is not None:
        tab_names.append("🏗️ 타워사별 현황")

    tabs = st.tabs(tab_names)

    # --- [탭 1] 전체 현장 목록 조회 화면 ---
    with tabs[0]:
        st.markdown("##### 🔍 현장 통합 검색")
        search_keyword = st.text_input(
            "검색어 입력",
            placeholder="현장명, 지부, 타워사 등을 입력하세요",
            label_visibility="collapsed",
        )

        filtered_df = df_main.copy()
        if search_keyword:
            mask = (
                filtered_df.astype(str)
                .apply(
                    lambda x: x.str.contains(search_keyword, case=False, na=False)
                )
                .any(axis=1)
            )
            filtered_df = filtered_df[mask]

        st.markdown(
            f"<p style='font-size: 0.75rem; color: #6b7280; margin-bottom: 5px;'>조회 결과: <b>{len(filtered_df)}</b>건</p>",
            unsafe_allow_html=True,
        )
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)

    # --- [탭 2] 지부별 점유율 통계 화면 ---
    with tabs[1]:
        st.markdown("##### 📊 지부별 점유율 분석")
        if df_sub1 is not None and not df_sub1.empty:
            st.dataframe(df_sub1, use_container_width=True, hide_index=True)
        else:
            st.info("지부별 점유율 데이터가 비어있습니다.")

    # --- [탭 3] 타워사별 현황 화면 (시트가 있을 경우 자동 활성화) ---
    if len(tabs) > 2 and df_sub2 is not None:
        with tabs[2]:
            st.markdown("##### 🏗️ 타워사별 상세 점유 현황")
            if not df_sub2.empty:
                st.dataframe(
                    df_sub2, use_container_width=True, hide_index=True
                )
            else:
                st.info("타워사별 데이터가 비어있습니다.")

    # 하단 푸터
    st.markdown("---")
    st.markdown(
        "<p style='text-align: center; font-size: 0.7rem; color: #9ca3af;'>Gyeonggi Regional Headquarters Dashboard</p>",
        unsafe_allow_html=True,
    )
