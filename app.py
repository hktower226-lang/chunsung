import pandas as pd
import streamlit as st

# 페이지 설정 (모바일 최적화 및 넓은 화면 레이아웃)
st.set_page_config(
    page_title="점유율 현황 모바일 대시보드",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# 모바일 가독성 향상 및 표/카드 스타일 CSS 적용
st.markdown(
    """
    <style>
    .main { font-size: 20px !important; }
    h1 { font-size: 26px !important; font-weight: bold; }
    h2 { font-size: 22px !important; font-weight: bold; }
    p, label, .stMarkdown { font-size: 16px !important; }
    
    .metric-card {
        background-color: #ffffff;
        padding: 14px;
        border-radius: 10px;
        margin-bottom: 10px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .subtotal-card {
        background-color: #f0f4f8;
        padding: 16px;
        border-radius: 10px;
        margin-bottom: 12px;
        border: 2px solid #2b6cb0;
    }
    /* 모바일 테이블 디자인 최적화 */
    table {
        width: 100% !important;
        font-size: 14px !important;
        text-align: center !important;
    }
    th {
        background-color: #f1f3f5 !important;
        text-align: center !important;
        font-size: 14px !important;
    }
    td {
        text-align: center !important;
        font-size: 14px !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)


@st.cache_data
def load_sheet_names():
    file_path = "data.xlsx"
    xls = pd.ExcelFile(file_path)
    return xls.sheet_names


try:
    sheet_names = load_sheet_names()
    file_path = "data.xlsx"
except Exception as e:
    st.error(
        f"엑셀 파일('data.xlsx')을 불러오지 못했습니다. 경로를 확인해주세요. 에러: {e}"
    )
    st.stop()

# 상단 타이틀
st.title("📊 점유율 현황 모바일 대시보드")
st.write("---")

# 3가지 화면(탭) 구성
tab1, tab2, tab3 = st.tabs(
    ["1. 2026 전체 소속", "2. 경기지역본부", "3. 타워사 별"]
)

# -------------------------------------------------------------------------
# 첫 번째 화면: 2026 전체 소속별 점유율
# -------------------------------------------------------------------------
with tab1:
    st.header("🏢 2026년 전체 소속별 점유율")
    st.markdown("전체 소속별 수치 및 퍼센테이지 현황입니다.")

    try:
        df_ratio = pd.read_excel(file_path, sheet_name="채용비율")
        t1 = df_ratio.iloc[0:2, :8].copy()
        t1.columns = t1.iloc[0]
        t1 = t1.drop(0).reset_index(drop=True)

        cols = [c for c in t1.columns if c not in ["소속", "합계", "비율"]]
        total_val = float(t1["합계"].values[0]) if "합계" in t1.columns else 198

        for col in cols:
            val = float(t1[col].values[0])
            pct = (val / total_val) * 100

            st.markdown(
                f"""
                <div class="metric-card" style="display: flex; justify-content: space-between; align-items: center;">
                    <div style="font-weight: bold; font-size: 18px;">{col}</div>
                    <div>
                        <span style="font-size: 18px; font-weight: bold;">{val:,.0f}명</span> &nbsp;|&nbsp; 
                        <span style="font-size: 18px; color: #e53e3e; font-weight: bold;">{pct:.1f}%</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
    except Exception as e:
        st.warning(f"데이터 처리 중 오류 발생: {e}")

# -------------------------------------------------------------------------
# 두 번째 화면: 경기지역본부 현장 점유율 (표 형태로 깔끔하게 정돈)
# -------------------------------------------------------------------------
with tab2:
    st.header("📍 경기지역본부 현장 점유율")
    st.markdown("원하시는 지부를 선택해 현장별 및 소계 현황을 확인하세요.")

    try:
        df_ratio = pd.read_excel(file_path, sheet_name="채용비율")
        t2 = df_ratio.iloc[7:76, :12].copy()
        t2.columns = [
            "지부",
            "현장명",
            "타워회사",
            "특이사항",
            "한노",
            "민노",
            "섬유",
            "건산",
            "직원",
            "미정",
            "총대수",
            "비고",
        ]
        t2["지부"] = t2["지부"].ffill()

        # 숫자형 변환
        for num_col in ["한노", "민노", "섬유", "건산", "직원", "미정", "총대수"]:
            t2[num_col] = (
                pd.to_numeric(t2[num_col], errors="coerce").fillna(0)
            )

        # 중복 소계 행 제거
        t2 = t2[
            ~(
                t2["지부"].astype(str).str.contains("소계")
                & (t2["총대수"] == 1)
                & t2["현장명"].isna()
            )
        ]

        # 지부 이름만 추출
        raw_branches = t2["지부"].dropna().astype(str).unique().tolist()
        clean_branches = []
        for b in raw_branches:
            cleaned = (
                b.replace("지부 소계", "")
                .replace("지부", "")
                .replace(" 소계", "")
                .strip()
            )
            if cleaned and cleaned not in clean_branches:
                clean_branches.append(cleaned)

        selected_branch = st.selectbox(
            "🔍 지부 선택", ["전체 보기"] + clean_branches
        )

        if selected_branch != "전체 보기":
            filtered_df = t2[
                t2["지부"].astype(str).str.contains(selected_branch)
            ]
        else:
            filtered_df = t2

        for _, row in filtered_df.iterrows():
            is_subtotal = "소계" in str(row["지부"]) or pd.isna(row["현장명"])

            h = row["한노"]
            m = row["민노"]
            s = row["섬유"]
            g = row["건산"]
            st_f = row["직원"]
            u = row["미정"]
            total = row["총대수"]

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
                        <div style="font-size: 18px; font-weight: bold; color: #2b6cb0; margin-bottom: 6px;">
                            📌 [{row['지부']}] 합계 (총 대수: {total:,.0f}대)
                        </div>
                        <table style="width:100%; margin-top:4px;">
                            <tr>
                                <th>한노</th><th>민노</th><th>섬유</th><th>건산</th><th>직원</th><th>미정</th>
                            </tr>
                            <tr>
                                <td><b>{h:.0f}</b><br><span style="color:#e53e3e; font-size:12px;">({h_pct:.1f}%)</span></td>
                                <td><b>{m:.0f}</b><br><span style="color:#3182ce; font-size:12px;">({m_pct:.1f}%)</span></td>
                                <td><b>{s:.0f}</b><br><span style="font-size:12px;">({s_pct:.1f}%)</span></td>
                                <td><b>{g:.0f}</b><br><span style="font-size:12px;">({g_pct:.1f}%)</span></td>
                                <td><b>{st_f:.0f}</b><br><span style="font-size:12px;">({st_f_pct:.1f}%)</span></td>
                                <td><b>{u:.0f}</b><br><span style="font-size:12px;">({u_pct:.1f}%)</span></td>
                            </tr>
                        </table>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f"""
                    <div class="metric-card">
                        <div style="font-weight: bold; font-size: 16px; color: #1a202c;">{row['현장명']}</div>
                        <div style="font-size: 13px; color: #718096; margin-bottom: 6px;">타워사: {row['타워회사']}</div>
                        <table style="width:100%;">
                            <tr>
                                <th>한노</th><th>민노</th><th>섬유</th><th>건산</th><th>직원</th><th>미정</th><th>총합</th>
                            </tr>
                            <tr>
                                <td style="color: #e53e3e; font-weight: bold;">{h:.0f}</td>
                                <td style="color: #3182ce; font-weight: bold;">{m:.0f}</td>
                                <td>{s:.0f}</td>
                                <td>{g:.0f}</td>
                                <td>{st_f:.0f}</td>
                                <td>{u:.0f}</td>
                                <td style="font-weight: bold;">{total:,.0f}</td>
                            </tr>
                        </table>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
    except Exception as e:
        st.warning(f"데이터 처리 중 오류 발생: {e}")

# -------------------------------------------------------------------------
# 세 번째 화면: 타워사 별 점유율
# -------------------------------------------------------------------------
with tab3:
    st.header("🏗️ 타워사 별 점유율 현황")
    st.markdown("한노 점유율을 맨 앞에 배치한 타워사별 현황입니다.")

    try:
        df_ratio = pd.read_excel(file_path, sheet_name="채용비율")
        t3 = df_ratio.iloc[81:, :10].copy()
        t3.columns = [
            "타워사",
            "현장수",
            "한노",
            "민노",
            "섬유",
            "건산",
            "직원",
            "기타",
            "합계",
            "한노점유율",
        ]
        t3 = t3.dropna(subset=["타워사"])

        for col in [
            "현장수",
            "한노",
            "민노",
            "섬유",
            "건산",
            "직원",
            "기타",
            "합계",
            "한노점유율",
        ]:
            t3[col] = pd.to_numeric(t3[col], errors="coerce").fillna(0)

        t3 = t3.dropna(subset=["합계"])
        t3 = t3.sort_values(by="한노점유율", ascending=False).reset_index(
            drop=True
        )

        for _, row in t3.iterrows():
            hanno_share = (
                row["한노점유율"] * 100 if pd.notna(row["한노점유율"]) else 0
            )

            st.markdown(
                f"""
                <div class="metric-card">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                        <span style="font-size: 16px; font-weight: bold; color: #2d3748;">{row['타워사']}</span>
                        <span style="font-size: 16px; font-weight: bold; color: #e53e3e;">한노 점유율: {hanno_share:.1f}%</span>
                    </div>
                    <div style="font-size: 13px; color: #718096; margin-bottom: 6px;">현장수: {row['현장수']}개 | 총합계: {row['합계']:.0f}명</div>
                    <table style="width:100%;">
                        <tr>
                            <th>한노</th><th>민노</th><th>섬유</th><th>건산</th><th>직원</th><th>기타</th>
                        </tr>
                        <tr>
                            <td style="color: #e53e3e; font-weight: bold;">{row['한노']:.0f}</td>
                            <td style="color: #3182ce; font-weight: bold;">{row['민노']:.0f}</td>
                            <td>{row['섬유']:.0f}</td>
                            <td>{row['건산']:.0f}</td>
                            <td>{row['직원']:.0f}</td>
                            <td>{row['기타']:.0f}</td>
                        </tr>
                    </table>
                </div>
                """,
                unsafe_allow_html=True,
            )
    except Exception as e:
        st.warning(f"데이터 처리 중 오류 발생: {e}")

# 하단 안내
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>모바일 화면 최적화 완료</p>",
    unsafe_allow_html=True,
)
